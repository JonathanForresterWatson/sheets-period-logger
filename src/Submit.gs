/**
 * Form submission handling.
 *
 * onFormSubmit(e) is the installed trigger. Everything below it works on a plain submission
 * object, so the same code path serves real submissions, the demo data and the offline tests.
 *
 * Submission object:
 *   { responseId, timestamp, customer, vehicle, startDate, endDate, value, note }
 */

var LOG_STATUS = {
  OK: 'OK',               // every day written
  PARTIAL: 'PARTIAL',     // some days skipped because they already had a value
  SKIPPED: 'SKIPPED',     // every day already had a value, nothing written
  DUPLICATE: 'DUPLICATE', // this form response was processed before
  ERROR: 'ERROR'          // rejected, nothing written
};

/** Installed trigger: Google Form "On form submit". */
function onFormSubmit(e) {
  var sub;
  try {
    sub = submissionFromEvent(e);
    return processSubmission(sub);
  } catch (err) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var log = getOrCreateSheet(ss, SHEETS.LOG, LOG_HEADERS);
      writeLog(log, sub || { timestamp: new Date() }, 0, 0, LOG_STATUS.ERROR, 'Unexpected: ' + err.message);
    } catch (ignored) {
      // logging failed too; the rethrow below still reaches the trigger failure report
    }
    throw err;
  }
}

/** Turns the FormResponse in the trigger event into a submission object. */
function submissionFromEvent(e) {
  if (!e || !e.response) throw new Error('No form response in the event. Install the trigger with installTrigger().');
  var response = e.response;
  var answers = {};
  response.getItemResponses().forEach(function (ir) {
    answers[ir.getItem().getTitle()] = ir.getResponse();
  });

  // Each customer has its own "Vehicle (name)" question; only the answered one is present.
  var vehicle = '';
  Object.keys(answers).forEach(function (title) {
    if (title.indexOf('Vehicle') === 0 && answers[title]) vehicle = String(answers[title]);
  });

  return {
    responseId: response.getId(),
    timestamp: response.getTimestamp(),
    customer: String(answers['Customer'] || ''),
    vehicle: vehicle,
    startDate: answers['Start date'],
    endDate: answers['End date'],
    value: answers['Daily value'],
    note: String(answers['Note'] || '')
  };
}

/**
 * Validates one submission, writes it into the Grid and logs the outcome.
 * Returns { status, written, skipped, message }.
 */
function processSubmission(sub) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = getConfig(ss);
    var log = getOrCreateSheet(ss, SHEETS.LOG, LOG_HEADERS);

    if (sub.responseId && logHasResponse(log, sub.responseId)) {
      return writeLog(log, sub, 0, 0, LOG_STATUS.DUPLICATE, 'This response was already processed');
    }

    var checked = validateSubmission(ss, cfg, sub);
    if (checked.error) {
      return writeLog(log, sub, 0, 0, LOG_STATUS.ERROR, checked.error);
    }

    var result = writeToGrid(ss, cfg, checked.sub, checked.days);
    var status = result.skipped === 0 ? LOG_STATUS.OK
      : (result.written === 0 ? LOG_STATUS.SKIPPED : LOG_STATUS.PARTIAL);

    var parts = checked.notes.slice();
    if (result.overwritten > 0) parts.push('Overwrote ' + result.overwritten + ' day(s) that already had a value');
    if (result.skipped > 0) parts.push('Skipped ' + result.skipped + ' day(s) that already had a value: ' + result.skippedDays.join(', '));
    if (parts.length === 0) parts.push('Wrote ' + result.written + ' day(s)');

    var outcome = writeLog(log, checked.sub, result.written, result.skipped, status, parts.join('. '));
    notify(cfg, checked.sub, outcome);
    return outcome;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns { error } or { sub, days, notes }.
 * `sub` comes back normalized: trimmed names, parsed dates, numeric value, ISO strings.
 */
function validateSubmission(ss, cfg, sub) {
  var notes = [];
  var customer = String(sub.customer || '').trim();
  var vehicle = String(sub.vehicle || '').trim();
  if (!customer) return { error: 'Missing customer' };
  if (!vehicle) return { error: 'Missing vehicle' };

  var vehicles = readCustomers(ss).vehiclesByCustomer[customer];
  if (!vehicles) return { error: 'Unknown customer: ' + customer };
  if (vehicles.indexOf(vehicle) === -1) return { error: 'Vehicle ' + vehicle + ' is not listed for ' + customer };

  var start = parseIsoDate(sub.startDate);
  var end = parseIsoDate(sub.endDate);
  if (!start) return { error: 'Start date is not a valid date: ' + sub.startDate };
  if (!end) return { error: 'End date is not a valid date: ' + sub.endDate };
  if (dayNumber(start) > dayNumber(end)) {
    var swap = start; start = end; end = swap;
    notes.push('Dates were reversed and have been swapped');
  }

  var days = enumerateDays(start, end);
  if (days.length > cfg.MAX_DAYS_PER_SUBMISSION) {
    return { error: 'Period is ' + days.length + ' days; the limit is ' + cfg.MAX_DAYS_PER_SUBMISSION };
  }

  var value = parseNumber(sub.value);
  if (value === null || value < 0) return { error: 'Daily value must be a number of 0 or more, got: ' + sub.value };

  var clean = {};
  Object.keys(sub).forEach(function (k) { clean[k] = sub[k]; });
  clean.customer = customer;
  clean.vehicle = vehicle;
  clean.start = start;
  clean.end = end;
  clean.startIso = isoString(start);
  clean.endIso = isoString(end);
  clean.value = value;
  return { sub: clean, days: days, notes: notes };
}

/**
 * Writes the daily value into the Grid for every day in the period.
 * Returns { written, skipped, overwritten, skippedDays }.
 */
function writeToGrid(ss, cfg, sub, days) {
  var grid = getOrCreateSheet(ss, SHEETS.GRID, GRID_HEADERS);

  var byMonth = {};
  var months = [];
  days.forEach(function (p) {
    var key = monthKey(p.y, p.m);
    if (!byMonth[key]) {
      byMonth[key] = { y: p.y, m: p.m, monthKey: key, days: [] };
      months.push(byMonth[key]);
    }
    byMonth[key].days.push(p.d);
  });

  var rows = ensureGridRows(grid, sub.customer, sub.vehicle, months, cfg.FY_START_MONTH);

  var result = { written: 0, skipped: 0, overwritten: 0, skippedDays: [] };
  months.forEach(function (mo) {
    var row = rows[mo.monthKey];
    var range = grid.getRange(row, GRID.COL_DAY_START, 1, GRID.DAYS);
    var cells = range.getValues()[0];
    mo.days.forEach(function (d) {
      var current = cells[d - 1];
      var empty = current === '' || current === null || current === undefined;
      if (empty) {
        cells[d - 1] = sub.value;
        result.written++;
      } else if (cfg.OVERWRITE_EXISTING) {
        cells[d - 1] = sub.value;
        result.written++;
        result.overwritten++;
      } else {
        result.skipped++;
        result.skippedDays.push(isoString({ y: mo.y, m: mo.m, d: d }));
      }
    });
    range.setValues([cells]);
  });
  return result;
}

// ---------------------------------------------------------------------------
// Log and notifications
// ---------------------------------------------------------------------------

function writeLog(log, sub, written, skipped, status, message) {
  var ts = isDate(sub.timestamp) ? sub.timestamp : new Date();
  log.appendRow([
    Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    sub.responseId || '',
    sub.customer || '',
    sub.vehicle || '',
    sub.startIso || sub.startDate || '',
    sub.endIso || sub.endDate || '',
    sub.value === undefined ? '' : sub.value,
    written,
    skipped,
    status,
    message
  ]);
  return { status: status, written: written, skipped: skipped, message: message };
}

function logHasResponse(log, responseId) {
  var last = log.getLastRow();
  if (last < 2) return false;
  var ids = log.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(responseId)) return true;
  }
  return false;
}

function notify(cfg, sub, outcome) {
  if (!cfg.NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail(
      cfg.NOTIFY_EMAIL,
      'Period Logger: ' + outcome.status + ' for ' + sub.customer + ' / ' + sub.vehicle,
      sub.startIso + ' to ' + sub.endIso + ', daily value ' + sub.value + '\n' +
      'Days written: ' + outcome.written + ', skipped: ' + outcome.skipped + '\n' + outcome.message
    );
  } catch (e) {
    console.log('Notification failed: ' + e.message);
  }
}
