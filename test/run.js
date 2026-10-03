'use strict';
/**
 * Offline tests. Run with: node test/run.js
 * No dependencies and no Google account needed; see harness.js for the mocks.
 */

const assert = require('node:assert/strict');
const { createEnvironment } = require('./harness');

const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

/** Values that cross the vm boundary have foreign prototypes; compare them as plain data. */
function plain(v) { return JSON.parse(JSON.stringify(v)); }
function same(actual, expected, message) { assert.deepEqual(plain(actual), plain(expected), message); }

/** Fresh environment with sheets and demo customers ready. */
function ready() {
  const env = createEnvironment();
  env.script.setupSheets();
  env.script.addDemoCustomers();
  return env;
}

/** Convenience: the 31 day cells of one Grid row, as a plain array. */
function dayCells(env, rowNumber) {
  const grid = env.sheet('Grid');
  return grid.getRange(rowNumber, 6, 1, 31).getValues()[0];
}

/** Finds a Grid row by customer, vehicle and month label. Returns the 1-based row number. */
function findRow(env, customer, vehicle, month) {
  const rows = env.rows('Grid');
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === customer && rows[i][2] === vehicle && rows[i][3] === month) return i + 1;
  }
  return -1;
}

function submission(overrides) {
  return Object.assign({
    responseId: 'r-' + Math.random().toString(36).slice(2, 8),
    timestamp: new Date('2026-10-03T14:00:00Z'),
    customer: 'Acme Logistics',
    vehicle: 'VAN-101',
    startDate: '2026-04-01',
    endDate: '2026-04-10',
    value: 100,
    note: ''
  }, overrides);
}

// ---------------------------------------------------------------------------
// Sheets and config
// ---------------------------------------------------------------------------

test('setupSheets creates the five sheets with headers and defaults', () => {
  const env = createEnvironment();
  env.script.setupSheets();
  ['Config', 'Customers', 'Grid', 'Log', 'Summary'].forEach(name => assert.ok(env.sheet(name), name + ' exists'));
  same(env.rows('Customers')[0], ['Customer', 'Vehicle', 'Active']);
  const gridHeader = env.rows('Grid')[0];
  assert.equal(gridHeader.length, 37);
  assert.equal(gridHeader[5], 1);
  assert.equal(gridHeader[35], 31);
  assert.equal(gridHeader[36], 'Total');
  const cfg = env.script.getConfig(env.ss);
  assert.equal(cfg.FY_START_MONTH, 4);
  assert.equal(cfg.OVERWRITE_EXISTING, false);
  assert.equal(cfg.MAX_DAYS_PER_SUBMISSION, 92);
  assert.ok(String(env.sheet('Summary').getRange(2, 1).getValue()).startsWith('=IFERROR(QUERY(Grid!A2:AK'));
});

test('setupSheets is safe to run twice', () => {
  const env = createEnvironment();
  env.script.setupSheets();
  env.script.setupSheets();
  assert.equal(env.ss.getSheets().length, 5);
  assert.equal(env.rows('Config').length, 1 + Object.keys(env.script.CONFIG_DEFAULTS).length);
});

test('config values are read from the sheet and coerced', () => {
  const env = ready();
  env.script.setConfigValue(env.ss, 'FY_START_MONTH', '7');
  env.script.setConfigValue(env.ss, 'OVERWRITE_EXISTING', 'TRUE');
  env.script.setConfigValue(env.ss, 'MAX_DAYS_PER_SUBMISSION', 10);
  const cfg = env.script.getConfig(env.ss);
  assert.equal(cfg.FY_START_MONTH, 7);
  assert.equal(cfg.OVERWRITE_EXISTING, true);
  assert.equal(cfg.MAX_DAYS_PER_SUBMISSION, 10);
});

test('onOpen adds the menu', () => {
  const env = createEnvironment();
  env.script.onOpen();
  assert.equal(env.ui.menus.length, 1);
  assert.equal(env.ui.menus[0].name, 'Period Logger');
  assert.ok(env.ui.menus[0].items.some(i => i.fn === 'runFullDemoSetup'));
});

// ---------------------------------------------------------------------------
// Date and financial-year helpers
// ---------------------------------------------------------------------------

test('parseIsoDate accepts real dates and rejects the rest', () => {
  const env = createEnvironment();
  same(env.script.parseIsoDate('2026-02-28'), { y: 2026, m: 2, d: 28 });
  same(env.script.parseIsoDate('2028-02-29'), { y: 2028, m: 2, d: 29 });
  assert.equal(env.script.parseIsoDate('2026-02-30'), null);
  assert.equal(env.script.parseIsoDate('2026-13-01'), null);
  assert.equal(env.script.parseIsoDate('March 3'), null);
  assert.equal(env.script.parseIsoDate(''), null);
});

test('fyLabel follows the configured start month', () => {
  const env = createEnvironment();
  const fy = env.script.fyLabel;
  assert.equal(fy(2026, 3, 4), '2025-26');
  assert.equal(fy(2026, 4, 4), '2026-27');
  assert.equal(fy(2027, 3, 4), '2026-27');
  assert.equal(fy(2026, 6, 7), '2025-26');
  assert.equal(fy(2026, 7, 7), '2026-27');
  assert.equal(fy(2026, 12, 1), '2026');
  assert.equal(fy(1999, 12, 4), '1999-00');
});

test('enumerateDays covers month ends and leap days', () => {
  const env = createEnvironment();
  const days = env.script.enumerateDays({ y: 2028, m: 2, d: 27 }, { y: 2028, m: 3, d: 2 });
  same(days.map(p => env.script.isoString(p)), ['2028-02-27', '2028-02-28', '2028-02-29', '2028-03-01', '2028-03-02']);
});

test('parseNumber handles strings, commas and junk', () => {
  const env = createEnvironment();
  assert.equal(env.script.parseNumber(95), 95);
  assert.equal(env.script.parseNumber(' 1,200.50 '), 1200.5);
  assert.equal(env.script.parseNumber('0'), 0);
  assert.equal(env.script.parseNumber('ninety'), null);
  assert.equal(env.script.parseNumber(''), null);
});

// ---------------------------------------------------------------------------
// Form builder
// ---------------------------------------------------------------------------

test('buildForm creates one page per active customer with the right vehicles and navigation', () => {
  const env = ready();
  const form = env.script.buildForm();
  const items = form.getItems();

  const customerItem = items[0];
  assert.equal(customerItem.getType(), 'MULTIPLE_CHOICE');
  same(customerItem.choices.map(c => c.value), ['Acme Logistics', 'Birch Farms', 'Cedar Builders']);

  const pages = items.filter(i => i.getType() === 'PAGE_BREAK');
  same(pages.map(p => p.getTitle()), ['Acme Logistics', 'Birch Farms', 'Cedar Builders', 'Period']);

  // each customer choice jumps to that customer's page
  customerItem.choices.forEach((c, i) => assert.equal(c.navigation, pages[i]));

  // the inactive vehicle DMP-9 is not offered
  const lists = items.filter(i => i.getType() === 'LIST');
  same(lists.map(l => l.getTitle()), ['Vehicle (Acme Logistics)', 'Vehicle (Birch Farms)', 'Vehicle (Cedar Builders)']);
  same(lists[2].choiceValues, ['EXC-3']);

  // after the Acme and Birch pages, jump to Period; Cedar reaches Period in normal order
  const period = pages[3];
  assert.equal(pages[1].goToPage, period);
  assert.equal(pages[2].goToPage, period);
  assert.equal(pages[0].goToPage, null);

  // period page questions
  const titles = items.slice(items.indexOf(period) + 1).map(i => i.getTitle());
  same(titles, ['Start date', 'End date', 'Daily value', 'Note']);
  assert.equal(items.find(i => i.getTitle() === 'Daily value').validation.min, 0);

  // form id and url stored in Config
  const cfg = env.script.getConfig(env.ss);
  assert.equal(cfg.FORM_ID, form.getId());
  assert.ok(cfg.FORM_URL.includes(form.getId()));
  assert.equal(form.destination.id, env.ss.getId());
});

test('buildForm rebuilds the same form in place after a customer change', () => {
  const env = ready();
  const first = env.script.buildForm();
  env.sheet('Customers').appendRow(['Delta Dairy', 'MLK-1', true]);
  const second = env.script.buildForm();
  assert.equal(second.getId(), first.getId(), 'same form, same URL');
  const pages = second.getItems().filter(i => i.getType() === 'PAGE_BREAK').map(p => p.getTitle());
  same(pages, ['Acme Logistics', 'Birch Farms', 'Cedar Builders', 'Delta Dairy', 'Period']);
});

test('buildForm refuses to run without customers', () => {
  const env = createEnvironment();
  env.script.setupSheets();
  assert.throws(() => env.script.buildForm(), /No active customers/);
});

test('installTrigger replaces any earlier submit trigger', () => {
  const env = ready();
  env.script.buildForm();
  env.script.installTrigger();
  env.script.installTrigger();
  assert.equal(env.triggers.length, 1);
  assert.equal(env.triggers[0].spec.handler, 'onFormSubmit');
  assert.equal(env.triggers[0].spec.event, 'ON_FORM_SUBMIT');
  assert.equal(env.triggers[0].spec.formId, env.script.getConfig(env.ss).FORM_ID);
});

// ---------------------------------------------------------------------------
// Writing periods into the grid
// ---------------------------------------------------------------------------

test('a period inside one month writes one row with the right cells and a total formula', () => {
  const env = ready();
  const out = env.script.processSubmission(submission({ startDate: '2026-04-05', endDate: '2026-04-09', value: 120 }));
  assert.equal(out.status, 'OK');
  assert.equal(out.written, 5);

  const row = findRow(env, 'Acme Logistics', 'VAN-101', 'Apr 2026');
  assert.ok(row > 0);
  const cells = dayCells(env, row);
  same(cells.slice(0, 10), ['', '', '', '', 120, 120, 120, 120, 120, '']);
  const rows = env.rows('Grid');
  assert.equal(rows[row - 1][0], '2026-27');
  assert.equal(rows[row - 1][4], 202604);
  assert.equal(rows[row - 1][36], '=SUM(RC[-31]:RC[-1])');

  const log = env.rows('Log');
  assert.equal(log.length, 2);
  assert.equal(log[1][9], 'OK');
  assert.equal(log[1][7], 5);
});

test('a period crossing a month and a financial-year boundary writes two rows with two FY labels', () => {
  const env = ready();
  const out = env.script.processSubmission(submission({ startDate: '2026-03-28', endDate: '2026-04-03', value: 50 }));
  assert.equal(out.status, 'OK');
  assert.equal(out.written, 7);

  const marRow = findRow(env, 'Acme Logistics', 'VAN-101', 'Mar 2026');
  const aprRow = findRow(env, 'Acme Logistics', 'VAN-101', 'Apr 2026');
  assert.ok(marRow > 0 && aprRow > 0);
  assert.equal(marRow + 1, aprRow, 'March row comes directly before April');
  assert.equal(env.rows('Grid')[marRow - 1][0], '2025-26');
  assert.equal(env.rows('Grid')[aprRow - 1][0], '2026-27');
  same(dayCells(env, marRow).slice(27), [50, 50, 50, 50]);
  same(dayCells(env, aprRow).slice(0, 4), [50, 50, 50, '']);
});

test('reversed dates are swapped and the log says so', () => {
  const env = ready();
  const out = env.script.processSubmission(submission({ startDate: '2026-05-20', endDate: '2026-05-10', value: 7 }));
  assert.equal(out.status, 'OK');
  assert.equal(out.written, 11);
  assert.match(out.message, /reversed/);
  const log = env.rows('Log')[1];
  assert.equal(log[4], '2026-05-10');
  assert.equal(log[5], '2026-05-20');
});

test('overlapping days are skipped by default and the log lists them', () => {
  const env = ready();
  env.script.processSubmission(submission({ responseId: 'a', startDate: '2026-04-10', endDate: '2026-04-12', value: 300 }));
  const out = env.script.processSubmission(submission({ responseId: 'b', startDate: '2026-04-12', endDate: '2026-04-15', value: 999 }));
  assert.equal(out.status, 'PARTIAL');
  assert.equal(out.written, 3);
  assert.equal(out.skipped, 1);
  assert.match(out.message, /2026-04-12/);
  const row = findRow(env, 'Acme Logistics', 'VAN-101', 'Apr 2026');
  same(dayCells(env, row).slice(9, 15), [300, 300, 300, 999, 999, 999]);
});

test('a fully overlapping period is logged as SKIPPED and changes nothing', () => {
  const env = ready();
  env.script.processSubmission(submission({ responseId: 'a', startDate: '2026-04-10', endDate: '2026-04-12', value: 300 }));
  const out = env.script.processSubmission(submission({ responseId: 'b', startDate: '2026-04-11', endDate: '2026-04-11', value: 1 }));
  assert.equal(out.status, 'SKIPPED');
  assert.equal(out.written, 0);
  const row = findRow(env, 'Acme Logistics', 'VAN-101', 'Apr 2026');
  same(dayCells(env, row).slice(9, 12), [300, 300, 300]);
});

test('OVERWRITE_EXISTING = TRUE replaces values and reports it', () => {
  const env = ready();
  env.script.setConfigValue(env.ss, 'OVERWRITE_EXISTING', 'TRUE');
  env.script.processSubmission(submission({ responseId: 'a', startDate: '2026-04-10', endDate: '2026-04-12', value: 300 }));
  const out = env.script.processSubmission(submission({ responseId: 'b', startDate: '2026-04-12', endDate: '2026-04-13', value: 5 }));
  assert.equal(out.status, 'OK');
  assert.equal(out.written, 2);
  assert.match(out.message, /Overwrote 1 day/);
  const row = findRow(env, 'Acme Logistics', 'VAN-101', 'Apr 2026');
  same(dayCells(env, row).slice(9, 13), [300, 300, 5, 5]);
});

test('the same form response is processed once', () => {
  const env = ready();
  const first = env.script.processSubmission(submission({ responseId: 'same', value: 10 }));
  const second = env.script.processSubmission(submission({ responseId: 'same', value: 10 }));
  assert.equal(first.status, 'OK');
  assert.equal(second.status, 'DUPLICATE');
  assert.equal(second.written, 0);
  assert.equal(env.rows('Grid').length, 2, 'still one grid row');
});

test('bad submissions are rejected with a reason and never touch the grid', () => {
  const env = ready();
  const cases = [
    [{ customer: 'Nobody Inc' }, /Unknown customer/],
    [{ vehicle: 'TRC-07' }, /not listed for Acme Logistics/],
    [{ startDate: '2026-02-30' }, /Start date/],
    [{ endDate: 'soon' }, /End date/],
    [{ value: 'lots' }, /Daily value/],
    [{ value: -4 }, /Daily value/],
    [{ startDate: '2026-01-01', endDate: '2026-06-30' }, /limit is 92/],
    [{ customer: '' }, /Missing customer/]
  ];
  cases.forEach(([overrides, pattern]) => {
    const out = env.script.processSubmission(submission(overrides));
    assert.equal(out.status, 'ERROR');
    assert.match(out.message, pattern);
  });
  assert.equal(env.rows('Grid').length, 1, 'header only');
  assert.equal(env.rows('Log').length, 1 + cases.length);
});

test('grid rows stay sorted by customer, vehicle and month however submissions arrive', () => {
  const env = ready();
  const order = [
    ['Cedar Builders', 'EXC-3', '2026-05-01', '2026-05-02'],
    ['Acme Logistics', 'VAN-102', '2026-04-01', '2026-04-02'],
    ['Acme Logistics', 'VAN-101', '2026-06-01', '2026-06-02'],
    ['Birch Farms', 'TRC-07', '2026-04-01', '2026-04-02'],
    ['Acme Logistics', 'VAN-101', '2026-04-01', '2026-04-02'],
    ['Acme Logistics', 'VAN-101', '2026-05-30', '2026-05-31']
  ];
  order.forEach(([customer, vehicle, startDate, endDate]) => {
    const out = env.script.processSubmission(submission({ customer, vehicle, startDate, endDate, value: 1 }));
    assert.equal(out.status, 'OK');
  });
  const keys = env.rows('Grid').slice(1).map(r => `${r[1]} | ${r[2]} | ${r[3]}`);
  same(keys, [
    'Acme Logistics | VAN-101 | Apr 2026',
    'Acme Logistics | VAN-101 | May 2026',
    'Acme Logistics | VAN-101 | Jun 2026',
    'Acme Logistics | VAN-102 | Apr 2026',
    'Birch Farms | TRC-07 | Apr 2026',
    'Cedar Builders | EXC-3 | May 2026'
  ]);
  // the row that was inserted in the middle kept its total formula and its values
  const junRow = findRow(env, 'Acme Logistics', 'VAN-101', 'Jun 2026');
  same(dayCells(env, junRow).slice(0, 3), [1, 1, '']);
  assert.equal(env.rows('Grid')[junRow - 1][36], '=SUM(RC[-31]:RC[-1])');
});

test('an inactive vehicle still accepts a late submission from an older form', () => {
  const env = ready();
  const out = env.script.processSubmission(submission({ customer: 'Cedar Builders', vehicle: 'DMP-9', startDate: '2026-04-01', endDate: '2026-04-01', value: 2 }));
  assert.equal(out.status, 'OK');
});

// ---------------------------------------------------------------------------
// End to end through the trigger, demo data and notifications
// ---------------------------------------------------------------------------

test('onFormSubmit reads the form response, including the per-customer vehicle question', () => {
  const env = ready();
  env.script.buildForm();
  const response = env.makeFormResponse('resp-1', new Date('2026-10-03T09:30:00Z'), {
    'Customer': 'Birch Farms',
    'Vehicle (Birch Farms)': 'PKP-12',
    'Start date': '2026-04-20',
    'End date': '2026-04-22',
    'Daily value': '1,250',
    'Note': 'harvest week'
  });
  const out = env.script.onFormSubmit({ response });
  assert.equal(out.status, 'OK');
  assert.equal(out.written, 3);
  const row = findRow(env, 'Birch Farms', 'PKP-12', 'Apr 2026');
  same(dayCells(env, row).slice(19, 22), [1250, 1250, 1250]);
  const log = env.rows('Log')[1];
  assert.equal(log[0], '2026-10-03 09:30:00');
  assert.equal(log[1], 'resp-1');
});

test('onFormSubmit without a response logs the problem and rethrows', () => {
  const env = ready();
  assert.throws(() => env.script.onFormSubmit({}), /No form response/);
  const log = env.rows('Log')[1];
  assert.equal(log[9], 'ERROR');
  assert.match(log[10], /No form response/);
});

test('addDemoSubmissions produces the expected mix of outcomes', () => {
  const env = ready();
  const results = env.script.addDemoSubmissions();
  same(results.map(r => r.status), ['OK', 'OK', 'OK', 'PARTIAL', 'OK', 'OK']);
  assert.equal(env.rows('Grid').length, 1 + 5, 'Mar+Apr for VAN-101, Apr for VAN-102, Apr for TRC-07, May for EXC-3');
});

test('runFullDemoSetup wires everything together in one call', () => {
  const env = createEnvironment();
  const results = env.script.runFullDemoSetup();
  assert.equal(results.length, 6);
  assert.equal(env.triggers.length, 1);
  assert.equal(Object.keys(env.forms).length, 1);
  assert.ok(env.ss.toasts.some(t => t.message.startsWith('Demo ready')));
});

test('resetDemo clears the grid and log but keeps customers and the form', () => {
  const env = createEnvironment();
  env.script.runFullDemoSetup();
  env.script.resetDemo();
  assert.equal(env.rows('Grid').length, 1);
  assert.equal(env.rows('Log').length, 1);
  assert.equal(env.rows('Customers').length, 8);
  assert.ok(env.script.getConfig(env.ss).FORM_ID);
});

test('NOTIFY_EMAIL sends one summary email per submission', () => {
  const env = ready();
  env.script.setConfigValue(env.ss, 'NOTIFY_EMAIL', 'ops@example.com');
  env.script.processSubmission(submission({ value: 42 }));
  assert.equal(env.mail.length, 1);
  assert.equal(env.mail[0].to, 'ops@example.com');
  assert.match(env.mail[0].subject, /OK for Acme Logistics \/ VAN-101/);
  assert.match(env.mail[0].body, /Days written: 10/);
});

// ---------------------------------------------------------------------------
// Runner
// ---------------------------------------------------------------------------

let passed = 0;
let failed = 0;
for (const t of tests) {
  try {
    t.fn();
    passed++;
    console.log('  ok   ' + t.name);
  } catch (err) {
    failed++;
    console.log('  FAIL ' + t.name);
    console.log('       ' + String(err.message).split('\n').join('\n       '));
  }
}
console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
