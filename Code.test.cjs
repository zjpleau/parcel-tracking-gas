const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync('Code.gs', 'utf8');
const context = vm.createContext({
  PropertiesService: {
    getScriptProperties: () => ({ getProperty: () => null })
  },
  Session: { getActiveUser: () => ({ getEmail: () => 'owner@example.com' }) }
});

vm.runInContext(source, context);

test('UPS estimated-delivery email remains eligible and exposes its tracking number', () => {
  const subject = 'UPS: Get Ready for Your Package!';
  const body = [
    'Your package is on the way.',
    'Your TARGET.COM package now has an estimated delivery date,',
    'and may be delivered by our trusted delivery partner.',
    'Tracking number: 1ZH4F859YW15767104'
  ].join(' ');

  assert.equal(context.isDelivered(subject, body), false);
  assert.deepEqual(
    JSON.parse(JSON.stringify(context.extractAllTrackingNumbers(`${subject}\n${body}`, 'mcinfo@ups.com'))),
    [{ trackingNumber: '1ZH4F859YW15767104', carrier: 'ups' }]
  );
});

test('future delivery wording is not treated as completed delivery', () => {
  const examples = [
    'Your package may be delivered by a trusted partner.',
    'Your package will be delivered tomorrow.',
    'Your package is expected to be delivered Friday.'
  ];

  examples.forEach(body => assert.equal(context.isDelivered('', body), false));
});

test('definitive delivery confirmations are still recognized', () => {
  const examples = [
    ['Your package was delivered', ''],
    ['', 'Your shipment has been delivered.'],
    ['', 'Delivery complete'],
    ['', 'Status: Delivered'],
    ['', "We've delivered your package."]
  ];

  examples.forEach(([subject, body]) => assert.equal(context.isDelivered(subject, body), true));
});

test('B&H merchant confirmation recognizes a labeled 12-digit FedEx number', () => {
  const text = [
    'B&H Photo Order #1131889923 Shipped',
    'Tracking #1: 540081458990',
    'Shipping Method: Expedited Delivery'
  ].join('\n');

  assert.deepEqual(
    JSON.parse(JSON.stringify(context.extractAllTrackingNumbers(text, 'ord-status@bhphotovideo.com'))),
    [{ trackingNumber: '540081458990', carrier: 'fedex' }]
  );
});

test("Sam's Club tracking-widget URL recognizes a 12-digit FedEx number", () => {
  const text = [
    'Your SamsClub.com order has shipped',
    '<img src="https://www.samsclub.com/api/node/vivaldi/v2/tracking-widget/535076984170"',
    'alt="tracking">'
  ].join('\n');

  assert.deepEqual(
    JSON.parse(JSON.stringify(context.extractAllTrackingNumbers(text, 'transaction@info.samsclub.com'))),
    [{ trackingNumber: '535076984170', carrier: 'fedex' }]
  );
});
