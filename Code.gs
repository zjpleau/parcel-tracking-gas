/**
 * Shipping Tracker Auto-Sync
 * Scans emails from shipping carriers and sends tracking info to Parcel API
 */

// Configuration - Set these in Script Properties (File > Project Properties > Script Properties)
const CONFIG = {
  PARCEL_API_KEY: PropertiesService.getScriptProperties().getProperty('PARCEL_API_KEY'),
  PARCEL_API_URL: PropertiesService.getScriptProperties().getProperty('PARCEL_API_URL') || 'https://api.parcel.app/external/add-delivery/',
  HOURS_TO_SCAN: 72, // How far back to scan for emails (2 hours with buffer)
  LABEL_PROCESSED: 'Tracking/Processed',
  LABEL_DELIVERED: 'Tracking/Delivered', // Label for delivered packages
  SEND_PUSH_NOTIFICATION: false, // Set to true to get push notifications when deliveries are added
  DAILY_RATE_LIMIT: 20, // Parcel API limit
  API_LOG_KEY: 'PARCEL_API_CALL_LOG', // Key for storing API call history
  SENT_TRACKING_KEY: 'SENT_TRACKING_NUMBERS', // Key for tracking numbers we've already sent
  DAILY_SUMMARY_KEY: 'DAILY_SUMMARY_DATA', // Key for storing daily summary data
  API_ACTIVITY_KEY: 'PARCEL_API_ACTIVITY_V2',
  SEND_EMAIL_SUMMARY: true, // Set to false to disable immediate email summaries after each run
  SEND_DAILY_SUMMARY: true, // Set to false to disable daily email summaries
  DAILY_SUMMARY_HOUR: 8, // Hour to send daily summary (0-23, 8 = 8am)
  SEND_ERROR_NOTIFICATIONS: true, // Set to false to disable immediate error emails
  EMAIL_ADDRESS: Session.getActiveUser().getEmail() // Your email address
};

// Universal tracking number patterns for all carriers
const TRACKING_PATTERNS = [
  // UPS (highest priority - very distinctive format)
  { pattern: /\b(1Z[0-9A-Z]{16})\b/gi, carrier: 'ups', priority: 1 },

  // USPS (high priority - distinctive prefixes)
  { pattern: /\b((?:94|93|92|95)\d{20})\b/g, carrier: 'usps', priority: 1 }, // 22-digit starting with 94/93/92/95
  { pattern: /\b((?:EA|EC|CP|RA|LK|LN|LM|RH|RB|RD|RE)\d{9}US)\b/gi, carrier: 'usps', priority: 1 }, // International
  { pattern: /\b((?:420\d{5})?(?:91|92|93|94|95)\d{20})\b/g, carrier: 'usps', priority: 2 }, // With or without 420 prefix

  // OnTrac (starts with C followed by 14 digits)
  { pattern: /\b(C\d{14})\b/gi, carrier: 'ontrac', priority: 3 },

  // FedEx (12, 15, or 20 digits - now with context-based detection)
  { pattern: /\b(\d{12})\b/g, carrier: 'fedex', priority: 5 }, // 12 digits
  { pattern: /\b(\d{15})\b/g, carrier: 'fedex', priority: 5 }, // 15 digits
  { pattern: /\b(\d{20})\b/g, carrier: 'fedex', priority: 5 }  // 20 digits
];

/**
 * HELPER: Generate direct tracking URL for the carrier
 */
function getTrackingUrl(carrier, number) {
  if (!carrier) return '#';
  const urls = {
    'ups': `https://www.ups.com/track?tracknum=${number}`,
    'usps': `https://tools.usps.com/go/TrackConfirmAction?tLabels=${number}`,
    'fedex': `https://www.fedex.com/fedextrack/?tracknumbers=${number}`,
    'ontrac': `https://www.ontrac.com/tracking?number=${number}`,
    'on-trac': `https://www.ontrac.com/tracking?number=${number}`
  };
  return urls[carrier.toLowerCase()] || '#';
}

/**
 * Main function to be triggered on a schedule (e.g., every 15 minutes)
 */
function scanShippingEmails() {
  const startTime = new Date();
  const summary = {
    startTime: startTime,
    emailsScanned: 0,
    trackingNumbersFound: 0,
    successfullySent: 0,
    failed: 0,
    trackingDetails: [],
    errors: []
  };

  try {
    Logger.log('Starting email scan...');
    
    // Check rate limit
    if (getTodayApiCallCount() >= CONFIG.DAILY_RATE_LIMIT) {
      Logger.log('⚠️ Daily API limit reached. Skipping.');
      return;
    }

    const hoursAgo = CONFIG.HOURS_TO_SCAN;
    const timestamp = new Date(Date.now() - hoursAgo * 60 * 60 * 1000);
    const query = `after:${Math.floor(timestamp.getTime() / 1000)} -label:${CONFIG.LABEL_PROCESSED} -label:${CONFIG.LABEL_DELIVERED} (tracking OR shipment OR shipped OR delivery OR "order confirmation" OR "your order" OR "Informed Delivery")`;

    const threads = GmailApp.search(query, 0, 50);

    threads.forEach(thread => {
      const messages = thread.getMessages();

      messages.forEach(message => {
        summary.emailsScanned++;
        const from = message.getFrom();
        const senderEmail = extractEmail(from);
        const subject = message.getSubject();
        const htmlBody = message.getBody(); 
        const combinedBody = message.getPlainBody() + '\n' + htmlBody;
        const date = message.getDate();

        // 1. SKIP CHECKS
        if (senderEmail === CONFIG.EMAIL_ADDRESS || subject.includes('Parcel Tracker')) return;
        if (senderEmail.includes('amazon')) return;
        if (isDelivered(subject, combinedBody)) {
           // Mark delivered threads so we don't scan them again
           createLabelIfNeeded(CONFIG.LABEL_DELIVERED);
           thread.addLabel(GmailApp.getUserLabelByName(CONFIG.LABEL_DELIVERED));
           return;
        }

        // 2. TRACKING EXTRACTION STRATEGY
        let trackingResults = [];

        // STRATEGY A: Specialized USPS Digest Parser
        // We check if this is an Informed Delivery email to use the advanced parser
        if (senderEmail.includes('informeddelivery.usps.com') && subject.includes('Digest')) {
          Logger.log('💌 Detected USPS Daily Digest - Running specialized parser...');
          trackingResults = extractUspsDigestData(htmlBody);
        } 
        // STRATEGY B: Standard Universal Parser (for everything else)
        else {
          trackingResults = extractAllTrackingNumbers(subject + '\n' + combinedBody, senderEmail);
        }

        // 3. PROCESS RESULTS
        if (trackingResults.length > 0) {
          Logger.log(`Found ${trackingResults.length} numbers in email from ${senderEmail}`);
          summary.trackingNumbersFound += trackingResults.length;

          // If we are using Strategy B, we need to calculate a general description for the whole email
          const generalDescription = extractDescription(subject, combinedBody, senderEmail);

          trackingResults.forEach(result => {
            const trackingNumber = result.trackingNumber;
            const carrier = result.carrier;
            
            // Use the specific description found in the Digest (Strategy A), 
            // otherwise fallback to the general email description (Strategy B)
            const finalDescription = result.description || generalDescription;

            if (hasBeenSent(trackingNumber)) {
              Logger.log(`Skipping ${trackingNumber} - already sent`);
              return;
            }

            markAsSent(trackingNumber);

            const success = sendToParcelAPI({
              tracking_number: trackingNumber,
              carrier_code: carrier,
              description: finalDescription, 
              send_push_confirmation: CONFIG.SEND_PUSH_NOTIFICATION
            });

            summary.trackingDetails.push({
              trackingNumber: trackingNumber,
              carrier: carrier,
              description: finalDescription,
              emailDate: date,
              success: success
            });

            if (success) summary.successfullySent++;
            else summary.failed++;
          });
        }
      });

      // Mark processed
      createLabelIfNeeded(CONFIG.LABEL_PROCESSED);
      thread.addLabel(GmailApp.getUserLabelByName(CONFIG.LABEL_PROCESSED));
    });

  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    summary.errors.push(error.message);
    if(CONFIG.SEND_ERROR_NOTIFICATIONS) sendErrorNotification(error, summary);
  }

  summary.apiCallsAfterScan = getTodayApiCallCount();
  accumulateDailySummary(summary);
  if (summary.successfullySent > 0 || summary.errors.length > 0 || summary.rateLimitReached) {
    sendEmailSummary(summary);
  }
}

/**
 * Extract ALL tracking numbers
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
        
        const hasGoodContext = isLikelyFedExTracking(text, trackingNum);
        if (!isFedExEmail && !hasGoodContext) continue;
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
 * SPECIALIZED PARSER: Handles USPS Informed Delivery Digest HTML
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

function isLikelyNotTrackingNumber(num, isFedExContext = false) {
  if (/^1?[2-9]\d{9}$/.test(num)) return true;
  if (/^\d{11}$/.test(num)) {
    if (/^(1|800|888|877|866|855|844|833|822|88[0-79])/.test(num)) return true;
  }
  if (!isFedExContext && /^[1-9]\d{10,14}$/.test(num) && num.length >= 11) {
    const uniqueDigits = new Set(num.split('')).size;
    if (uniqueDigits >= 6) return true;
  }
  if (!isFedExContext && /^\d{11,13}$/.test(num)) {
    if (!/^1Z/.test(num) && !/^(94|93|92|95|420)/.test(num)) return true;
  }
  if (/^(\d)\1{9,}$/.test(num)) return true;
  if (/^(?:0123456789|1234567890|9876543210|0987654321)/.test(num)) return true;
  return false;
}

function isLikelyFedExTracking(text, trackingNumber) {
  if (!text || !trackingNumber) return false;
  const index = text.indexOf(trackingNumber);
  if (index === -1) return false;
  const context = text.substring(Math.max(0, index - 300), Math.min(text.length, index + trackingNumber.length + 300)).toLowerCase();
  let score = 0;
  if (context.includes('fedex') || context.includes('federal express')) score += 3;
  if (context.includes('tracking') || context.includes('shipment')) score += 2;
  return score >= 3;
}

function extractEmail(fromField) {
  const match = fromField.match(/<([^>]+)>/);
  return match ? match[1].toLowerCase() : fromField.toLowerCase();
}

function cleanHtml(text) {
  if (!text) return '';
  return text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '').replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
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
    const match = cleanSubject.match(p);
    if (match && match[1].length < 50) return match[1].trim();
  }
  if (senderEmail) {
    const domain = senderEmail.split('@')[1];
    if (domain) {
       const parts = domain.split('.');
       const len = parts.length;
       let mainPart = len >= 2 ? parts[len - 2] : parts[0];
       const commonSLDs = ['co', 'com', 'org', 'net', 'edu', 'gov', 'ac'];
       if (len >= 3 && commonSLDs.includes(mainPart.toLowerCase())) mainPart = parts[len - 3];
       return mainPart.charAt(0).toUpperCase() + mainPart.slice(1);
    }
  }
  return cleanSubject;
}

function sendToParcelAPI(data) {
  try {
    const options = { method: 'post', contentType: 'application/json', headers: { 'api-key': CONFIG.PARCEL_API_KEY }, payload: JSON.stringify(data), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(CONFIG.PARCEL_API_URL, options);
    const success = response.getResponseCode() >= 200 && response.getResponseCode() < 300;
    recordApiActivity(success);
    return success;
  } catch (error) {
    recordApiActivity(false);
    return false;
  }
}

function recordApiActivity(success) {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty(CONFIG.API_ACTIVITY_KEY);
  let activity = json ? JSON.parse(json) : [];
  const now = Date.now();
  activity = activity.filter(a => a.t > now - 86400000);
  activity.push({ t: now, s: success ? 1 : 0 });
  props.setProperty(CONFIG.API_ACTIVITY_KEY, JSON.stringify(activity));
}

function getTodayApiCallCount() {
  const json = PropertiesService.getScriptProperties().getProperty(CONFIG.API_ACTIVITY_KEY);
  if (!json) return 0;
  const activity = JSON.parse(json);
  const today = new Date().setHours(0,0,0,0);
  return activity.filter(a => a.t >= today).length;
}

function createLabelIfNeeded(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);
  return label;
}

function isDelivered(subject, body) {
  const keywords = [/delivered/i, /delivery complete/i, /successfully delivered/i];
  const combined = `${subject} ${body}`;
  return keywords.some(p => p.test(combined));
}

function hasBeenSent(trackingNumber) {
  const json = PropertiesService.getScriptProperties().getProperty(CONFIG.SENT_TRACKING_KEY);
  if (!json) return false;
  return JSON.parse(json).includes(trackingNumber);
}

function markAsSent(trackingNumber) {
  const props = PropertiesService.getScriptProperties();
  let list = JSON.parse(props.getProperty(CONFIG.SENT_TRACKING_KEY) || '[]');
  if (!list.includes(trackingNumber)) {
    list.push(trackingNumber);
    props.setProperty(CONFIG.SENT_TRACKING_KEY, JSON.stringify(list.slice(-200)));
  }
}

function accumulateDailySummary(summary) {
  const props = PropertiesService.getScriptProperties();
  let daily = JSON.parse(props.getProperty(CONFIG.DAILY_SUMMARY_KEY) || '{"date":"","totalEmailsScanned":0,"totalSuccessfullySent":0,"totalFailed":0,"trackingDetails":[],"errors":[]}');
  if (daily.date !== new Date().toDateString()) {
    daily = { date: new Date().toDateString(), totalEmailsScanned: 0, totalSuccessfullySent: 0, totalFailed: 0, trackingDetails: [], errors: [] };
  }
  daily.totalEmailsScanned += summary.emailsScanned;
  daily.totalSuccessfullySent += summary.successfullySent;
  daily.totalFailed += summary.failed;
  if (summary.trackingDetails) daily.trackingDetails.push(...summary.trackingDetails.filter(d => d.success));
  if (daily.trackingDetails.length > 30) daily.trackingDetails = daily.trackingDetails.slice(-30);
  if (summary.errors) daily.errors.push(...summary.errors);
  props.setProperty(CONFIG.DAILY_SUMMARY_KEY, JSON.stringify(daily));
}

function shouldSkipEmail(senderEmail, subject, body) {
  if (senderEmail === CONFIG.EMAIL_ADDRESS) return true;
  if (subject.includes('Parcel Tracker')) return true;
  if (isDelivered(subject, body)) return true;
  return false;
}

function sendEmailSummary(summary) {
  if (!CONFIG.SEND_EMAIL_SUMMARY || !summary) return;
  const subject = `Parcel Tracker Sync - ${summary.successfullySent} Added`;
  let html = `<div style="font-family: Arial; max-width: 600px;"><h2>Sync Report</h2><p>Scanned: ${summary.emailsScanned}</p><p>Added: ${summary.successfullySent}</p>`;
  if (summary.trackingDetails && summary.trackingDetails.length > 0) {
    html += `<table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%;"><tr style="background: #eee;"><th>Carrier</th><th>Number</th><th>Description</th></tr>`;
    summary.trackingDetails.forEach(d => {
      html += `<tr><td>${(d.carrier || 'UPS').toUpperCase()}</td><td><a href="${getTrackingUrl(d.carrier, d.trackingNumber)}">${d.trackingNumber}</a></td><td>${d.description || 'N/A'}</td></tr>`;
    });
    html += `</table>`;
  }
  html += `</div>`;
  GmailApp.sendEmail(CONFIG.EMAIL_ADDRESS, subject, '', { htmlBody: html });
}

function checkAndSendDailySummary() {
  if (!CONFIG.SEND_DAILY_SUMMARY) return;
  const now = new Date();
  if (now.getHours() !== CONFIG.DAILY_SUMMARY_HOUR) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAST_DAILY_SUMMARY_SENT') === now.toDateString()) return;
  const data = JSON.parse(props.getProperty(CONFIG.DAILY_SUMMARY_KEY) || 'null');
  if (!data) return;
  let html = `<div style="font-family: Arial; max-width: 600px;"><h2>Daily Summary</h2><p>Total: ${data.totalSuccessfullySent}</p></div>`;
  GmailApp.sendEmail(CONFIG.EMAIL_ADDRESS, `Daily Parcel Tracker Summary - ${data.totalSuccessfullySent} Added`, '', { htmlBody: html });
  props.setProperty('LAST_DAILY_SUMMARY_SENT', now.toDateString());
  props.deleteProperty(CONFIG.DAILY_SUMMARY_KEY);
}

function sendErrorNotification(error, summary) {
  if (!CONFIG.SEND_ERROR_NOTIFICATIONS) return;
  const html = `<h3>Error Alert</h3><p>Error: ${error.message}</p><p>Stack: ${error.stack}</p>`;
  GmailApp.sendEmail(CONFIG.EMAIL_ADDRESS, 'ERROR - Parcel Tracker Auto-Sync', '', { htmlBody: html });
}

function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scanShippingEmails').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('checkAndSendDailySummary').timeBased().everyHours(1).create();
}

function testEmailSummary() {
  sendEmailSummary({ emailsScanned: 10, successfullySent: 1, trackingDetails: [{ trackingNumber: '1Z12345TEST', carrier: 'ups', description: 'Test' }] });
}

function sendDailySummaryNow() {
  const data = JSON.parse(PropertiesService.getScriptProperties().getProperty(CONFIG.DAILY_SUMMARY_KEY) || 'null');
  if (!data) return;
  GmailApp.sendEmail(CONFIG.EMAIL_ADDRESS, 'Manual Summary', '', { htmlBody: `<h2>Manual Summary</h2><p>Total: ${data.totalSuccessfullySent}</p>` });
}

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
        if (!hasBeenSent(res.trackingNumber)) {
          markAsSent(res.trackingNumber);
          sendToParcelAPI({ tracking_number: res.trackingNumber, carrier_code: res.carrier, description: extractDescription(sub, body, email).substring(0, 100) });
        }
      });
    });
  });
}
