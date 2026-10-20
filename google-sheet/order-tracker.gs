/**
 * Keepsake Song Co. order tracker (Google Apps Script).
 *
 * Paste this into the Sheet: Extensions > Apps Script. Then reload the Sheet
 * and use the "Keepsake" menu > "1. Set up this sheet".
 *
 * What it does:
 *  - New orders arrive as rows automatically (sent by the Stripe webhook).
 *  - Rush and extra style add-ons tick the matching box on that customer's order.
 *  - Ticking "Delivered" moves the customer from Keepsake – Customers to
 *    Keepsake – Delivered in Sender, which starts the delivered emails.
 *  - "Due by" turns red when an order is late and not yet delivered.
 */

const SHEET_NAME = 'Orders';
const SENDER_CUSTOMERS = 'azNrj5'; // Keepsake – Customers
const SENDER_DELIVERED = 'erNJXw'; // Keepsake – Delivered

const HEADERS = [
  'Order date', 'Due by', 'Name', 'Email', 'Phone', 'Package', 'Song for',
  'Occasion', 'Needed by (their words)', 'Rush', 'Extra style', 'Paid £',
  'Story received', 'Preview sent', 'Final sent', 'Delivered', 'Delivered on',
  'Sender', 'Notes', 'Stripe ref',
];
const COL = {};
HEADERS.forEach(function (h, i) { COL[h] = i + 1; });
const CHECKBOX_COLS = ['Rush', 'Extra style', 'Story received', 'Preview sent', 'Final sent', 'Delivered'];
const NORMAL_DAYS = 7;
const RUSH_DAYS = 2;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Keepsake')
    .addItem('1. Set up this sheet', 'setup')
    .addItem('2. Show details for Netlify', 'showNetlifyDetails')
    .addItem('Add a test order', 'addTestOrder')
    .addToUi();
}

/** One-off set-up: Sender key, secret, layout and the "Delivered" trigger. */
function setup() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const answer = ui.prompt(
    'Sender API key',
    'Paste your Sender API access token (Sender > Settings > API access tokens).' +
      (props.getProperty('SENDER_TOKEN') ? '\nLeave blank to keep the one already saved.' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const token = answer.getResponseText().trim();
  if (token) props.setProperty('SENDER_TOKEN', token);
  if (!props.getProperty('SENDER_TOKEN')) {
    ui.alert('No Sender key saved, so set-up stopped. Run it again when you have the key.');
    return;
  }
  if (!props.getProperty('SHEET_SECRET')) {
    props.setProperty('SHEET_SECRET', Utilities.getUuid().replace(/-/g, ''));
  }

  layOutSheet_();

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'handleEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();

  ui.alert('Set-up done. Next: Extensions > Apps Script > Deploy > New deployment ' +
    '(type Web app, execute as Me, access Anyone). Then use Keepsake > 2. Show details for Netlify.');
}

function layOutSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
  }
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#2F4D46').setFontColor('#FFFFFF').setWrap(true);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  sh.setColumnWidths(1, HEADERS.length, 110);
  sh.setColumnWidth(COL['Email'], 200);
  sh.setColumnWidth(COL['Notes'], 240);
  sh.setColumnWidth(COL['Sender'], 170);
  sh.getRange(2, COL['Order date'], sh.getMaxRows() - 1, 1).setNumberFormat('ddd d mmm yyyy');
  sh.getRange(2, COL['Due by'], sh.getMaxRows() - 1, 1).setNumberFormat('ddd d mmm');
  sh.getRange(2, COL['Delivered on'], sh.getMaxRows() - 1, 1).setNumberFormat('ddd d mmm');
  sh.getRange(2, COL['Paid £'], sh.getMaxRows() - 1, 1).setNumberFormat('£0.00');

  // Late and not delivered: red "Due by". Delivered: whole row goes green.
  const last = columnLetter_(HEADERS.length);
  const due = columnLetter_(COL['Due by']);
  const delivered = columnLetter_(COL['Delivered']);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + delivered + '2=TRUE')
      .setBackground('#E3EEE7')
      .setRanges([sh.getRange('A2:' + last)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + due + '2<>"",$' + due + '2<TODAY(),$' + delivered + '2<>TRUE)')
      .setBackground('#F4C7B8').setFontColor('#8A2E0B')
      .setRanges([sh.getRange(due + '2:' + due)]).build(),
  ];
  sh.setConditionalFormatRules(rules);

  let sum = ss.getSheetByName('Summary');
  if (!sum) sum = ss.insertSheet('Summary');
  const o = SHEET_NAME + '!';
  sum.getRange('A1:B6').setValues([
    ['At a glance', ''],
    ['Intro offer spots sold (of 15)', '=COUNTIF(' + o + 'F2:F,"Intro Bundle")'],
    ['Orders in progress', '=COUNTIFS(' + o + 'D2:D,"<>",' + o + 'P2:P,"<>TRUE")'],
    ['Late (past due, not delivered)', '=COUNTIFS(' + o + 'B2:B,"<"&TODAY(),' + o + 'P2:P,"<>TRUE")'],
    ['Waiting for their story', '=COUNTIFS(' + o + 'D2:D,"<>",' + o + 'F2:F,"<>*add-on*",' + o + 'M2:M,"<>TRUE")'],
    ['Delivered', '=COUNTIF(' + o + 'P2:P,TRUE)'],
  ]);
  sum.getRange('A1').setFontWeight('bold').setFontSize(13);
  sum.setColumnWidth(1, 260);
}

/** Shows the two values to paste into Netlify. */
function showNetlifyDetails() {
  const props = PropertiesService.getScriptProperties();
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(
    'Paste these into Netlify (Site configuration > Environment variables):\n\n' +
    'SHEET_WEBHOOK_URL\n' + (url || '(not deployed yet: Deploy > New deployment > Web app)') + '\n\n' +
    'SHEET_SECRET\n' + (props.getProperty('SHEET_SECRET') || '(run 1. Set up this sheet first)'));
}

/** Receives orders from the Stripe webhook on Netlify. */
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return text_('Bad request');
  }
  const secret = PropertiesService.getScriptProperties().getProperty('SHEET_SECRET');
  if (!secret || data.secret !== secret) return text_('Forbidden');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (data.session && findRow_(sh, COL['Stripe ref'], data.session)) return text_('OK duplicate');
    if (data.type === 'addon') return text_(recordAddon_(sh, data));
    addOrderRow_(sh, data);
    return text_('OK added');
  } finally {
    lock.releaseLock();
  }
}

function addOrderRow_(sh, d) {
  const row = sh.getLastRow() + 1;
  const values = HEADERS.map(function () { return ''; });
  values[COL['Order date'] - 1] = d.created ? new Date(d.created * 1000) : new Date();
  values[COL['Name'] - 1] = d.name || '';
  values[COL['Email'] - 1] = (d.email || '').toLowerCase();
  values[COL['Phone'] - 1] = d.phone ? "'" + d.phone : '';
  values[COL['Package'] - 1] = d.package || '';
  values[COL['Song for'] - 1] = d.songFor || '';
  values[COL['Occasion'] - 1] = d.occasion || '';
  values[COL['Needed by (their words)'] - 1] = d.neededBy || '';
  values[COL['Paid £'] - 1] = typeof d.amount === 'number' ? d.amount : '';
  values[COL['Sender'] - 1] = d.type === 'order' ? 'In Customers' : '';
  values[COL['Notes'] - 1] = d.note || '';
  values[COL['Stripe ref'] - 1] = d.session || '';
  sh.getRange(row, 1, 1, values.length).setValues([values]);
  CHECKBOX_COLS.forEach(function (h) { sh.getRange(row, COL[h]).insertCheckboxes(); });
  const a = columnLetter_(COL['Order date']) + row;
  const r = columnLetter_(COL['Rush']) + row;
  sh.getRange(row, COL['Due by']).setFormula(
    '=IF(' + a + '="","",' + a + '+IF(' + r + ',' + RUSH_DAYS + ',' + NORMAL_DAYS + '))');
  return row;
}

/** Ticks Rush or Extra style on the customer's latest undelivered order. */
function recordAddon_(sh, d) {
  const email = (d.email || '').toLowerCase();
  const col = d.package === 'Rush' ? COL['Rush'] : COL['Extra style'];
  const last = sh.getLastRow();
  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][COL['Email'] - 1]).toLowerCase() === email && rows[i][COL['Delivered'] - 1] !== true) {
        const row = i + 2;
        sh.getRange(row, col).setValue(true);
        const notes = sh.getRange(row, COL['Notes']);
        notes.setValue([notes.getValue(), d.package + ' paid ' + Utilities.formatDate(new Date(), 'Europe/London', 'd MMM') +
          ' (' + d.session + ')'].filter(String).join('; '));
        return 'OK add-on ticked';
      }
    }
  }
  // No matching order (different email used?): add its own row so nothing is lost.
  const row = addOrderRow_(sh, Object.assign({}, d, {
    package: d.package + ' add-on',
    note: 'No order found for this email. Match it to their order by hand.',
  }));
  sh.getRange(row, col).setValue(true);
  return 'OK add-on row';
}

/** Installable on-edit trigger: ticking "Delivered" moves them in Sender. */
function handleEdit(e) {
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;
  const col = COL['Delivered'];
  if (range.getColumn() > col || range.getLastColumn() < col) return;

  for (let row = Math.max(2, range.getRow()); row <= range.getLastRow(); row++) {
    const ticked = sh.getRange(row, col).getValue() === true;
    const status = sh.getRange(row, COL['Sender']);
    const email = String(sh.getRange(row, COL['Email']).getValue()).trim();
    if (!email) continue;
    if (!ticked) {
      if (String(status.getValue()).indexOf('Delivered') !== -1) {
        status.setValue('Unticked: still in Delivered in Sender, change there if needed');
      }
      continue;
    }
    if (String(status.getValue()) === 'Moved to Delivered') continue;
    const result = moveToDelivered_(email);
    status.setValue(result);
    if (result === 'Moved to Delivered') sh.getRange(row, COL['Delivered on']).setValue(new Date());
  }
}

function moveToDelivered_(email) {
  // Same call the Stripe webhook uses: adds the group (keeping their other
  // groups) and starts the Delivered automation.
  const add = sender_('post', '/subscribers',
    { email: email, groups: [SENDER_DELIVERED], trigger_automation: true });
  if (add.code >= 300) return 'Error ' + add.code + ': ' + add.body.slice(0, 120);
  const remove = sender_('delete', '/subscribers/groups/' + SENDER_CUSTOMERS, { subscribers: [email] });
  if (remove.code >= 300 && remove.code !== 404) {
    return 'In Delivered, but still in Customers (error ' + remove.code + ')';
  }
  return 'Moved to Delivered';
}

function sender_(method, path, body) {
  const token = PropertiesService.getScriptProperties().getProperty('SENDER_TOKEN');
  if (!token) return { code: 401, body: 'No Sender key saved. Run Keepsake > 1. Set up this sheet.' };
  const res = UrlFetchApp.fetch('https://api.sender.net/v2' + path, {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  return { code: res.getResponseCode(), body: res.getContentText() };
}

/** Adds a pretend order so you can see how a row looks. Delete it after. */
function addTestOrder() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  addOrderRow_(sh, {
    type: 'test', package: 'Song', name: 'Test Order', email: 'test@example.com',
    songFor: 'Mum', occasion: 'Birthday', neededBy: '1 Dec', amount: 34.99,
    note: 'TEST ROW: delete me. Do not tick Delivered on this one.',
  });
}

function findRow_(sh, col, value) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === value) return i + 2;
  return 0;
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function text_(s) {
  return ContentService.createTextOutput(s);
}
