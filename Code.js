/**
 * Shipping Tracker Auto-Sync
 * Scans emails from shipping carriers and sends tracking info to Parcel API
 */

/**
 * HELPER: Generate direct tracking URL for the carrier
 */
function getTrackingUrl(carrier, number) {
  const urls = {
    'ups': `https://www.ups.com/track?tracknum=${number}`,
    'usps': `https://tools.usps.com/go/TrackConfirmAction?tLabels=${number}`,
    'fedex': `https://www.fedex.com/fedextrack/?tracknumbers=${number}`,
    'ontrac': `https://www.ontrac.com/tracking?number=${number}`
  };
  return urls[carrier.toLowerCase()] || '#';
}

// Configuration
const CONFIG = {
  PARCEL_API_KEY: PropertiesService.getScriptProperties().getProperty('PARCEL_API_KEY'),
  PARCEL_API_URL: PropertiesService.getScriptProperties().getProperty('PARCEL_API_URL') || 'https://api.parcel.app/external/add-delivery/',
  HOURS_TO_SCAN: 336, // 14 days
  LABEL_PROCESSED: 'Tracking/Processed',
  LABEL_DELIVERED: 'Tracking/Delivered', 
  SEND_PUSH_NOTIFICATION: false, 
  DAILY_RATE_LIMIT: 20, 
  API_ACTIVITY_KEY: 'PARCEL_API_ACTIVITY_V2', 
  SENT_TRACKING_KEY: 'SENT_TRACKING_NUMBERS', 
  DAILY_SUMMARY_KEY: 'DAILY_SUMMARY_DATA', 
  SEND_EMAIL_SUMMARY: true, 
  SEND_DAILY_SUMMARY: true, 
  DAILY_SUMMARY_HOUR: 8, 
  SEND_ERROR_NOTIFICATIONS: true, 
  EMAIL_ADDRESS: Session.getActiveUser().getEmail() 
};

// Universal tracking number patterns
const TRACKING_PATTERNS = [
  { pattern: /\b(1Z[0-9A-Z]{16})\b/gi, carrier: 'ups', priority: 1 },
  { pattern: /\b((?:94|93|92|95)\d{20})\b/g, carrier: 'usps', priority: 1 },
  { pattern: /\b((?:EA|EC|CP|RA|LK|LN|LM|RH|RB|RD|RE)\d{9}US)\b/gi, carrier: 'usps', priority: 1 },
  { pattern: /\b((?:420\d{5})?(?:91|92|93|94|95)\d{20})\b/g, carrier: 'usps', priority: 2 },
  { pattern: /(C\d{14})/gi, carrier: 'ontrac', priority: 3 },
  { pattern: /\b(\d{12})\b/g, carrier: 'fedex', priority: 5 },
  { pattern: /\b(\d{15})\b/g, carrier: 'fedex', priority: 5 },
  { pattern: /\b(\d{20})\b/g, carrier: 'fedex', priority: 5 }
];

/**
 * Main background task
 */
function scanShippingEmails() {
  const summary = {
    emailsScanned: 0,
    trackingNumbersFound: 0,
    successfullySent: 0,
    failed: 0,
    trackingDetails: [],
    errors: [],
    apiCallsBeforeScan: getTodayApiCallCount()
  };

  try {
    if (summary.apiCallsBeforeScan >= CONFIG.DAILY_RATE_LIMIT) {
      summary.rateLimitReached = true;
      if (CONFIG.SEND_EMAIL_SUMMARY) sendEmailSummary(summary);
      return;
    }

    const timestamp = new Date(Date.now() - CONFIG.HOURS_TO_SCAN * 60 * 60 * 1000);
    const query = `after:${Math.floor(timestamp.getTime() / 1000)} -label:${CONFIG.LABEL_PROCESSED} -label:${CONFIG.LABEL_DELIVERED} (tracking OR shipment OR shipped OR delivery OR "order confirmation" OR "your order" OR "Informed Delivery")`;

    const threads = GmailApp.search(query, 0, 50);

    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        summary.emailsScanned++;
        const from = message.getFrom();
        const senderEmail = extractEmail(from);
        const subject = message.getSubject();
        const body = message.getPlainBody() + '\n' + message.getBody();

        if (shouldSkipEmail(senderEmail, subject, body)) return;

        let trackingResults = [];
        if (senderEmail.includes('informeddelivery.usps.com') && subject.includes('Digest')) {
          trackingResults = extractUspsDigestData(message.getBody());
        } else {
          trackingResults = extractAllTrackingNumbers(subject + '\n' + body, senderEmail);
        }

        if (trackingResults.length > 0) {
          const generalDescription = extractDescription(subject, body, senderEmail);

          trackingResults.forEach(result => {
            const trackingNumber = result.trackingNumber;
            if (hasBeenSent(trackingNumber)) return;
            if (getTodayApiCallCount() >= CONFIG.DAILY_RATE_LIMIT) {
               summary.rateLimitReached = true;
               return;
            }

            const cleanDesc = (result.description || generalDescription || "").substring(0, 100);
            const success = sendToParcelAPI({
              tracking_number: trackingNumber,
              carrier_code: result.carrier.toLowerCase(), 
              description: cleanDesc, 
              send_push_confirmation: CONFIG.SEND_PUSH_NOTIFICATION
            });

            if (success) {
              markAsSent(trackingNumber);
              summary.successfullySent++;
            } else {
              summary.failed++;
            }

            summary.trackingDetails.push({
              trackingNumber: trackingNumber,
              carrier: result.carrier,
              description: cleanDesc,
              emailDate: message.getDate(),
              success: success
            });
          });
        }
      });
      createLabelIfNeeded(CONFIG.LABEL_PROCESSED);
      thread.addLabel(GmailApp.getUserLabelByName(CONFIG.LABEL_PROCESSED));
    });
  } catch (error) {
    summary.errors.push(error.message);
    if(CONFIG.SEND_ERROR_NOTIFICATIONS) sendErrorNotification(error, summary);
  }

  accumulateDailySummary(summary);
  if (summary.successfullySent > 0 || summary.errors.length > 0 || summary.rateLimitReached) {
    sendEmailSummary(summary);
  }
}

/**
 * Universal tracking number parser
 */
function extractAllTrackingNumbers(text, senderEmail) {
  const found = new Map(); 
  const isFedExEmail = senderEmail && senderEmail.toLowerCase().includes('fedex.com');

  TRACKING_PATTERNS.forEach(patternConfig => {
    const { pattern, carrier, priority } = patternConfig;
    const matches = text.matchAll(pattern);

    for (const match of matches) {
      const trackingNum = match[1].trim().replace(/\s+/g, '');
      if (trackingNum.length < 10) continue;
      if (found.has(trackingNum) && found.get(trackingNum).priority <= priority) continue;

      if (carrier === 'fedex') {
        if (text.includes('1Z' + trackingNum) || text.includes('1ZXH' + trackingNum)) continue;
        if (!isFedExEmail && !isLikelyFedExTracking(text, trackingNum)) continue;
        if (isLikelyNotTrackingNumber(trackingNum, true)) continue;
      } else {
        if (isLikelyNotTrackingNumber(trackingNum, false)) continue;
      }
      found.set(trackingNum, { carrier, priority });
    }
  });

  return Array.from(found.entries()).map(([trackingNumber, info]) => ({
    trackingNumber,
    carrier: info.carrier
  }));
}

/**
 * Specialized USPS parser
 */
function extractUspsDigestData(htmlBody) {
  const results = [];
  const digestRegex = /FROM:\s*<b><span[^>]*>([^<]{3,100}?)<\/span><\/b>(?:(?!FROM:)[\s\S])*?<span[^>]*>(\d{20,})<\/span>/gi;
  const matches = htmlBody.matchAll(digestRegex);
  
  for (const match of matches) {
    let shipperName = cleanHtml(match[1]);
    const trackingNumber = match[2].trim();
    shipperName = shipperName.replace(trackingNumber, '').replace(/\d{10,}/g, '').trim();
    shipperName = shipperName.toLowerCase().split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
    if (shipperName.length > 50) shipperName = shipperName.substring(0, 47) + '...';

    results.push({
      trackingNumber: trackingNumber,
      carrier: 'usps',
      description: shipperName
    });
  }
  return results;
}

/**
 * Filter helpers
 */
function shouldSkipEmail(senderEmail, subject, body) {
  if (senderEmail === CONFIG.EMAIL_ADDRESS) return true;
  if (subject.includes('Parcel Tracker')) return true;
  if (isDelivered(subject, body)) return true;
  return false;
}

function isDelivered(subject, body) {
  const combined = `${subject} ${body}`.toLowerCase();
  const notDeliveredYet = [/can't be delivered/i, /delivery error/i, /couldn't be delivered/i, /delayed/i];
  if (notDeliveredYet.some(p => p.test(combined))) return false;
  const keywords = [/was delivered/i, /delivery complete/i, /successfully delivered/i, /^delivered: /i];
  return keywords.some(p => p.test(combined));
}

function isLikelyNotTrackingNumber(num, isFedExContext) {
  if (/^1?[2-9]\d{9}$/.test(num)) return true;
  if (/^\d{11}$/.test(num) && /^(1|800|888|877|866|855|844|833|822|88[0-79])/.test(num)) return true;
  if (!isFedExContext && /^[1-9]\d{10,14}$/.test(num) && new Set(num.split('')).size >= 6) return true;
  if (!isFedExContext && /^\d{11,13}$/.test(num) && !/^1Z/.test(num) && !/^(94|93|92|95|420)/.test(num)) return true;
  if (/^(\d)\1{9,}$/.test(num) || /^(?:0123456789|1234567890|9876543210|0987654321)/.test(num)) return true;
  return false;
}

function isLikelyFedExTracking(text, num) {
  const idx = text.indexOf(num);
  if (idx === -1) return false;
  const context = text.substring(Math.max(0, idx - 300), Math.min(text.length, idx + num.length + 300)).toLowerCase();
  let score = 0;
  if (context.includes('fedex') || context.includes('federal express')) score += 3;
  if (context.includes('tracking') || context.includes('shipment')) score += 2;
  return score >= 3;
}

function cleanHtml(text) {
  if (!text) return '';
  return text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '').replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
}

function extractEmail(from) {
  const match = from.match(/<([^>]+)>/);
  return match ? match[1].toLowerCase() : from.toLowerCase();
}

function extractDescription(subject, body, senderEmail) {
  const cleanBody = cleanHtml(body);
  const cleanSubject = cleanHtml(subject);
  if (senderEmail.includes('usps.com')) {
    const match = cleanBody.match(/From: ([^:]+?)(?: Expected|$)/i);
    if (match) return match[1].trim();
  }
  if (senderEmail.includes('ups.com')) {
    const match = cleanBody.match(/From\s+([A-Z0-9\s\.\-]{3,30})/i);
    if (match) return match[1].trim();
  }
  const merchantPatterns = [/order from\s+(.+?)(?:\s+has|\s+is|\s+-|$)/i, /shipment from\s+(.+?)(?:\s+has|\s+is|\s+-|$)/i, /(.+?)\s+order\s+(?:confirmation|has shipped)/i, /Your\s+(.+?)\s+order/i];
  for (const p of merchantPatterns) {
    const m = cleanSubject.match(p);
    if (m && m[1].length < 50) return m[1].trim();
  }
  if (senderEmail) {
    const parts = senderEmail.split('@')[1].split('.');
    const len = parts.length;
    let mainPart = len >= 2 ? parts[len - 2] : parts[0];
    const slds = ['co', 'com', 'org', 'net', 'edu', 'gov', 'ac'];
    if (len >= 3 && slds.includes(mainPart.toLowerCase())) mainPart = parts[len - 3];
    return mainPart.charAt(0).toUpperCase() + mainPart.slice(1);
  }
  return cleanSubject;
}

/**
 * API and Quota Logic
 */
function sendToParcelAPI(data) {
  try {
    const options = { method: 'post', contentType: 'application/json', headers: { 'api-key': CONFIG.PARCEL_API_KEY }, payload: JSON.stringify(data), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(CONFIG.PARCEL_API_URL, options);
    const success = response.getResponseCode() >= 200 && response.getResponseCode() < 300;
    recordApiActivity(success);
    return success;
  } catch (e) {
    recordApiActivity(false);
    return false;
  }
}

function recordApiActivity(success) {
  const props = PropertiesService.getScriptProperties();
  const activity = JSON.parse(props.getProperty(CONFIG.API_ACTIVITY_KEY) || '[]');
  const now = Date.now();
  const filtered = activity.filter(a => a.t > now - 86400000);
  filtered.push({ t: now, s: success ? 1 : 0 });
  props.setProperty(CONFIG.API_ACTIVITY_KEY, JSON.stringify(filtered));
}

function getTodayApiCallCount() {
  const activity = JSON.parse(PropertiesService.getScriptProperties().getProperty(CONFIG.API_ACTIVITY_KEY) || '[]');
  const today = new Date().setHours(0,0,0,0);
  return activity.filter(a => a.t >= today).length;
}

function showApiQuotaStatus() {
  const activity = JSON.parse(PropertiesService.getScriptProperties().getProperty(CONFIG.API_ACTIVITY_KEY) || '[]').sort((a,b) => a.t-b.t);
  const count = activity.filter(a => a.t > Date.now() - 86400000).length;
  Logger.log(`Used: ${count} / ${CONFIG.DAILY_RATE_LIMIT}`);
  if (count > 0) Logger.log(`Next slot: ${new Date(activity[0].t + 86400000).toLocaleString()}`);
}

function hasBeenSent(num) {
  const list = JSON.parse(PropertiesService.getScriptProperties().getProperty(CONFIG.SENT_TRACKING_KEY) || '[]');
  return list.includes(num);
}

function markAsSent(num) {
  const props = PropertiesService.getScriptProperties();
  const list = JSON.parse(props.getProperty(CONFIG.SENT_TRACKING_KEY) || '[]');
  if (!list.includes(num)) {
    list.push(num);
    props.setProperty(CONFIG.SENT_TRACKING_KEY, JSON.stringify(list.slice(-200)));
  }
}

function accumulateDailySummary(summary) {
  const props = PropertiesService.getScriptProperties();
  let daily = JSON.parse(props.getProperty(CONFIG.DAILY_SUMMARY_KEY) || '{"date":"","totalEmailsScanned":0,"totalSuccessfullySent":0,"trackingDetails":[],"errors":[]}');
  if (daily.date !== new Date().toDateString()) { daily = { date: new Date().toDateString(), totalEmailsScanned: 0, totalSuccessfullySent: 0, trackingDetails: [], errors: [] }; }
  daily.totalEmailsScanned += summary.emailsScanned;
  daily.totalSuccessfullySent += summary.successfullySent;
  if (summary.trackingDetails) daily.trackingDetails.push(...summary.trackingDetails);
  props.setProperty(CONFIG.DAILY_SUMMARY_KEY, JSON.stringify(daily));
}

function sendEmailSummary(summary) {
  if (!CONFIG.SEND_EMAIL_SUMMARY || !summary) return;
  const props = PropertiesService.getScriptProperties();
  const today = new Date().toDateString();
  if (summary.rateLimitReached) {
    if (props.getProperty('LAST_LIMIT_ALERT_SENT') === today) return;
    props.setProperty('LAST_LIMIT_ALERT_SENT', today);
  }
  let html = `<div style="font-family: Arial; max-width: 600px;"><h2 style="color: #1a73e8;">Parcel Tracker Sync Report</h2><p><b>Successfully Added:</b> ${summary.successfullySent}</p>`;
  if (summary.trackingDetails.length > 0) {
    html += `<table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%;"><tr style="background: #eee;"><th>Carrier</th><th>Number</th><th>Description</th></tr>`;
    summary.trackingDetails.forEach(d => {
      html += `<tr><td>${d.carrier.toUpperCase()}</td><td><a href="${getTrackingUrl(d.carrier, d.trackingNumber)}">${d.trackingNumber}</a></td><td>${d.description}</td></tr>`;
    });
    html += `</table>`;
  }
  html += `<hr><p style="font-size: 12px;">Quota: ${getTodayApiCallCount()} / ${CONFIG.DAILY_RATE_LIMIT} used.</p></div>`;
  GmailApp.sendEmail(CONFIG.EMAIL_ADDRESS, summary.rateLimitReached ? 'Parcel Tracker - Limit Reached' : `Parcel Tracker Sync - ${summary.successfullySent} Added`, '', { htmlBody: html });
}

function checkAndSendDailySummary() {
  const now = new Date();
  if (now.getHours() !== CONFIG.DAILY_SUMMARY_HOUR) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAST_DAILY_SUMMARY_SENT') === now.toDateString()) return;
  const dailyData = JSON.parse(props.getProperty(CONFIG.DAILY_SUMMARY_KEY) || 'null');
  if (!dailyData || dailyData.totalSuccessfullySent === 0) return;
  
  let html = `<div style="font-family: Arial; max-width: 600px;"><h2 style="color: #1a73e8;">Daily Shipping Summary</h2><p><b>Total Packages:</b> ${dailyData.totalSuccessfullySent}</p>`;
  html += `<table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%;"><tr style="background: #eee;"><th>Carrier</th><th>Number</th><th>Description</th></tr>`;
  dailyData.trackingDetails.forEach(d => { html += `<tr><td>${d.carrier.toUpperCase()}</td><td><a href="${getTrackingUrl(d.carrier, d.trackingNumber)}">${d.trackingNumber}</a></td><td>${d.description}</td></tr>`; });
  html += `</table></div>`;
  
  GmailApp.sendEmail(CONFIG.EMAIL_ADDRESS, `Daily Parcel Tracker Summary - ${dailyData.totalSuccessfullySent} Added`, '', { htmlBody: html });
  props.setProperty('LAST_DAILY_SUMMARY_SENT', now.toDateString());
  props.deleteProperty(CONFIG.DAILY_SUMMARY_KEY);
}

function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scanShippingEmails').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('checkAndSendDailySummary').timeBased().everyHours(1).create();
  Logger.log('✅ Triggers updated.');
}

function testEmailSummary() {
  sendEmailSummary({ successfullySent: 1, trackingDetails: [{ carrier: 'ups', trackingNumber: '1Z12345TEST', description: 'Test Package' }], rateLimitReached: false, emailsScanned: 5 });
}

function sendDailySummaryNow() {
  const dailyData = JSON.parse(PropertiesService.getScriptProperties().getProperty(CONFIG.DAILY_SUMMARY_KEY) || 'null');
  if (!dailyData) { Logger.log('No data.'); return; }
  let html = `<h2>Manual Daily Summary</h2><p><b>Total:</b> ${dailyData.totalSuccessfullySent}</p>`;
  GmailApp.sendEmail(CONFIG.EMAIL_ADDRESS, 'Manual Daily Summary', '', { htmlBody: html });
}

function resetApiQuotaLog() { PropertiesService.getScriptProperties().deleteProperty(CONFIG.API_ACTIVITY_KEY); Logger.log('Reset.'); }
function clearSentTrackingNumbers() { PropertiesService.getScriptProperties().deleteProperty(CONFIG.SENT_TRACKING_KEY); Logger.log('Cleared.'); }

function manualSyncRecentEmails() {
  const threads = GmailApp.search(`after:${Math.floor((Date.now() - 14 * 24 * 60 * 60 * 1000) / 1000)} (tracking OR shipment OR shipped OR delivery)`, 0, 50);
  threads.forEach(t => {
    t.getMessages().forEach(m => {
      const email = extractEmail(m.getFrom());
      const sub = m.getSubject();
      const body = m.getPlainBody() + '\n' + m.getBody();
      if (shouldSkipEmail(email, sub, body)) return;
      const results = extractAllTrackingNumbers(sub + '\n' + body, email);
      results.forEach(res => {
        if (hasBeenSent(res.trackingNumber)) return;
        const desc = extractDescription(sub, body, email).substring(0, 100);
        const success = sendToParcelAPI({ tracking_number: res.trackingNumber, carrier_code: res.carrier.toLowerCase(), description: desc });
        if (success) markAsSent(res.trackingNumber);
      });
    });
  });
  Logger.log('Manual sync complete.');
}
