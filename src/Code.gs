/**
 * Period Logger for Google Sheets
 *
 * A Google Form feeds a calendar-style grid in a spreadsheet:
 *   Form (customer -> that customer's vehicles -> date range + daily value)
 *   -> onFormSubmit trigger
 *   -> validation, overlap detection, financial-year handling
 *   -> Grid sheet (one row per customer / vehicle / month, one column per day)
 *   -> Log sheet (every submission, with status and message)
 *   -> Summary sheet (live totals by customer and financial year)
 *
 * Entry points (also available from the "Period Logger" menu in the spreadsheet):
 *   runFullDemoSetup()    one-click install: sheets, demo customers, form, trigger, demo submissions
 *   setupSheets()         create the Config, Customers, Grid, Log and Summary sheets
 *   addDemoCustomers()    fill the Customers sheet with sample data
 *   buildForm()           create or rebuild the Google Form from the Customers sheet
 *   installTrigger()      connect the form to onFormSubmit
 *   addDemoSubmissions()  push sample periods through the same code path as a real submission
 *   resetDemo()           clear the Grid and Log sheets, keep customers and the form
 *
 * Runtime: Apps Script V8. No libraries. Tested offline with test/run.js (Node).
 */

var SHEETS = {
  CONFIG: 'Config',
  CUSTOMERS: 'Customers',
  GRID: 'Grid',
  LOG: 'Log',
  SUMMARY: 'Summary'
};

// Grid sheet layout (1-based column numbers).
var GRID = {
  COL_FY: 1,          // A  financial year label, e.g. 2026-27
  COL_CUSTOMER: 2,    // B
  COL_VEHICLE: 3,     // C
  COL_MONTH: 4,       // D  month label, e.g. Apr 2026
  COL_MONTH_KEY: 5,   // E  sortable key yyyymm, hidden
  COL_DAY_START: 6,   // F  day 1 ... AJ day 31
  DAYS: 31,
  COL_TOTAL: 37       // AK SUM of the day cells (formula)
};

var GRID_HEADERS = ['FY', 'Customer', 'Vehicle', 'Month', 'MonthKey'];
for (var _d = 1; _d <= GRID.DAYS; _d++) GRID_HEADERS.push(_d);
GRID_HEADERS.push('Total');

var LOG_HEADERS = ['Timestamp', 'Response ID', 'Customer', 'Vehicle', 'Start', 'End',
  'Daily value', 'Days written', 'Days skipped', 'Status', 'Message'];

var CUSTOMER_HEADERS = ['Customer', 'Vehicle', 'Active'];

var CONFIG_DEFAULTS = {
  FY_START_MONTH: 4,                 // 1 = calendar year, 4 = April to March, 7 = July to June
  FORM_TITLE: 'Vehicle period submission',
  FORM_ID: '',                       // filled in by buildForm()
  FORM_URL: '',
  OVERWRITE_EXISTING: false,         // FALSE = never overwrite a day that already has a value
  MAX_DAYS_PER_SUBMISSION: 92,
  NOTIFY_EMAIL: ''                   // optional: one address that gets a line per submission
};

var CONFIG_NOTES = {
  FY_START_MONTH: 'Month the financial year starts (1 to 12). 4 = April to March.',
  FORM_TITLE: 'Title of the Google Form. Rebuild the form after changing it.',
  FORM_ID: 'Set automatically when the form is built. Leave blank to create a new form.',
  FORM_URL: 'Share this link with the people who submit periods.',
  OVERWRITE_EXISTING: 'TRUE to let a new submission replace values already in the grid.',
  MAX_DAYS_PER_SUBMISSION: 'Submissions longer than this are rejected and logged.',
  NOTIFY_EMAIL: 'Optional. One email per submission goes here. Blank = off.'
};

// ---------------------------------------------------------------------------
// Menu and one-click setup
// ---------------------------------------------------------------------------

function onOpen() {
  var ui = getUiOrNull();
  if (!ui) return;
  ui.createMenu('Period Logger')
    .addItem('Run full demo setup', 'runFullDemoSetup')
    .addSeparator()
    .addItem('1. Set up sheets', 'setupSheets')
    .addItem('2. Add demo customers', 'addDemoCustomers')
    .addItem('3. Build or rebuild form', 'buildForm')
    .addItem('4. Install submit trigger', 'installTrigger')
    .addItem('5. Add demo submissions', 'addDemoSubmissions')
    .addSeparator()
    .addItem('Reset demo data', 'resetDemo')
    .addToUi();
}

function runFullDemoSetup() {
  setupSheets();
  addDemoCustomers();
  buildForm();
  installTrigger();
  var results = addDemoSubmissions();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig(ss);
  var summary = results.map(function (r) { return r.status; }).join(', ');
  say(ss, 'Demo ready. Form: ' + cfg.FORM_URL + '\nDemo submissions: ' + summary);
  return results;
}

// ---------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var config = getOrCreateSheet(ss, SHEETS.CONFIG, ['Key', 'Value', 'What it does']);
  if (config.getLastRow() < 2) {
    var rows = Object.keys(CONFIG_DEFAULTS).map(function (k) {
      return [k, CONFIG_DEFAULTS[k], CONFIG_NOTES[k] || ''];
    });
    config.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  config.setColumnWidths(1, 1, 220);
  config.setColumnWidths(2, 1, 320);
  config.setColumnWidths(3, 1, 420);

  getOrCreateSheet(ss, SHEETS.CUSTOMERS, CUSTOMER_HEADERS);

  var grid = getOrCreateSheet(ss, SHEETS.GRID, GRID_HEADERS);
  grid.setFrozenColumns(4);
  grid.hideColumns(GRID.COL_MONTH_KEY);
  grid.setColumnWidths(GRID.COL_DAY_START, GRID.DAYS, 36);

  getOrCreateSheet(ss, SHEETS.LOG, LOG_HEADERS);

  var summary = getOrCreateSheet(ss, SHEETS.SUMMARY, null);
  summary.getRange(1, 1).setValue('Totals by customer and financial year').setFontWeight('bold');
  summary.getRange(2, 1).setFormula(
    '=IFERROR(QUERY(' + SHEETS.GRID + '!A2:AK, "select B, A, sum(AK) where B is not null ' +
    'group by B, A label B \'Customer\', A \'FY\', sum(AK) \'Total\'", 0), "No data yet")');
  summary.getRange(1, 5).setValue('Totals by vehicle').setFontWeight('bold');
  summary.getRange(2, 5).setFormula(
    '=IFERROR(QUERY(' + SHEETS.GRID + '!A2:AK, "select B, C, sum(AK) where C is not null ' +
    'group by B, C label B \'Customer\', C \'Vehicle\', sum(AK) \'Total\'", 0), "No data yet")');

  say(ss, 'Sheets ready: Config, Customers, Grid, Log, Summary.');
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (headers && sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

function getConfig(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.CONFIG, ['Key', 'Value', 'What it does']);
  var cfg = {};
  Object.keys(CONFIG_DEFAULTS).forEach(function (k) { cfg[k] = CONFIG_DEFAULTS[k]; });
  var last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      var key = String(r[0]).trim();
      if (key) cfg[key] = r[1];
    });
  }
  cfg.FY_START_MONTH = clampInt(cfg.FY_START_MONTH, 1, 12, CONFIG_DEFAULTS.FY_START_MONTH);
  cfg.MAX_DAYS_PER_SUBMISSION = clampInt(cfg.MAX_DAYS_PER_SUBMISSION, 1, 366, CONFIG_DEFAULTS.MAX_DAYS_PER_SUBMISSION);
  cfg.OVERWRITE_EXISTING = toBool(cfg.OVERWRITE_EXISTING);
  cfg.FORM_ID = String(cfg.FORM_ID || '').trim();
  cfg.FORM_URL = String(cfg.FORM_URL || '').trim();
  cfg.FORM_TITLE = String(cfg.FORM_TITLE || CONFIG_DEFAULTS.FORM_TITLE).trim();
  cfg.NOTIFY_EMAIL = String(cfg.NOTIFY_EMAIL || '').trim();
  return cfg;
}

function setConfigValue(ss, key, value) {
  var sheet = getOrCreateSheet(ss, SHEETS.CONFIG, ['Key', 'Value', 'What it does']);
  var last = sheet.getLastRow();
  if (last >= 2) {
    var keys = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) {
        sheet.getRange(i + 2, 2).setValue(value);
        return;
      }
    }
  }
  sheet.appendRow([key, value, CONFIG_NOTES[key] || '']);
}

function clampInt(v, min, max, fallback) {
  var n = parseInt(v, 10);
  if (isNaN(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

function toBool(v) {
  if (typeof v === 'boolean') return v;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === '1' || s === 'y';
}

// ---------------------------------------------------------------------------
// Customers
// ---------------------------------------------------------------------------

/**
 * Reads the Customers sheet.
 * Returns { customers: [{name, vehicles}] (active vehicles only, in sheet order),
 *           vehiclesByCustomer: { name: [all vehicles, active or not] } }
 */
function readCustomers(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.CUSTOMERS, CUSTOMER_HEADERS);
  var last = sheet.getLastRow();
  var customers = [];
  var byName = {};
  var allByName = {};
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) {
      var name = String(r[0]).trim();
      var vehicle = String(r[1]).trim();
      if (!name || !vehicle) return;
      var active = r[2] === '' ? true : toBool(r[2]);
      if (!allByName[name]) allByName[name] = [];
      if (allByName[name].indexOf(vehicle) === -1) allByName[name].push(vehicle);
      if (!active) return;
      if (!byName[name]) {
        byName[name] = { name: name, vehicles: [] };
        customers.push(byName[name]);
      }
      if (byName[name].vehicles.indexOf(vehicle) === -1) byName[name].vehicles.push(vehicle);
    });
  }
  return { customers: customers, vehiclesByCustomer: allByName };
}

function addDemoCustomers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEETS.CUSTOMERS, CUSTOMER_HEADERS);
  if (sheet.getLastRow() >= 2) {
    say(ss, 'Customers sheet already has data. Nothing added.');
    return;
  }
  var rows = [
    ['Acme Logistics', 'VAN-101', true],
    ['Acme Logistics', 'VAN-102', true],
    ['Acme Logistics', 'TRK-200', true],
    ['Birch Farms', 'TRC-07', true],
    ['Birch Farms', 'PKP-12', true],
    ['Cedar Builders', 'EXC-3', true],
    ['Cedar Builders', 'DMP-9', false]
  ];
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  say(ss, 'Added ' + rows.length + ' demo customer rows.');
}

// ---------------------------------------------------------------------------
// Demo submissions and reset
// ---------------------------------------------------------------------------

/**
 * Pushes sample periods through processSubmission(), exactly as the form trigger would.
 * The set is chosen to exercise the interesting paths: a period that crosses a month and a
 * financial-year boundary, an overlap, reversed dates, and a single day.
 */
function addDemoSubmissions() {
  var stamp = new Date();
  var demo = [
    { responseId: 'demo-1', customer: 'Acme Logistics', vehicle: 'VAN-101', startDate: '2026-03-28', endDate: '2026-04-03', value: 120, note: 'Crosses March into April, so two FY labels' },
    { responseId: 'demo-2', customer: 'Acme Logistics', vehicle: 'VAN-102', startDate: '2026-04-01', endDate: '2026-04-14', value: 95, note: '' },
    { responseId: 'demo-3', customer: 'Birch Farms', vehicle: 'TRC-07', startDate: '2026-04-10', endDate: '2026-04-12', value: 300, note: '' },
    { responseId: 'demo-4', customer: 'Birch Farms', vehicle: 'TRC-07', startDate: '2026-04-12', endDate: '2026-04-15', value: 300, note: 'Overlaps demo-3 on the 12th' },
    { responseId: 'demo-5', customer: 'Cedar Builders', vehicle: 'EXC-3', startDate: '2026-05-20', endDate: '2026-05-10', value: 450, note: 'Dates entered backwards' },
    { responseId: 'demo-6', customer: 'Acme Logistics', vehicle: 'VAN-101', startDate: '2026-04-20', endDate: '2026-04-20', value: 120, note: 'Single day' }
  ];
  return demo.map(function (d) {
    d.timestamp = stamp;
    return processSubmission(d);
  });
}

function resetDemo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEETS.GRID, SHEETS.LOG].forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var last = sheet.getLastRow();
    if (last >= 2) sheet.deleteRows(2, last - 1);
  });
  say(ss, 'Grid and Log cleared. Customers and the form were kept.');
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function getUiOrNull() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null; // no UI when running from a trigger or the editor
  }
}

function say(ss, message) {
  try {
    ss.toast(message, 'Period Logger', 8);
  } catch (e) {
    // toast is not available from every context; the message is also returned to callers
  }
  console.log(message);
}
