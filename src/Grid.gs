/**
 * Date, financial-year and grid-row helpers.
 *
 * Dates are handled as plain {y, m, d} objects and UTC day numbers, so the result does not
 * depend on the script time zone or on daylight-saving changes.
 */

var MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** True for a usable Date object (checked by shape, so it also works across sandboxes). */
function isDate(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

/** 'yyyy-MM-dd' or a Date -> {y, m, d}, or null when the input is not a real date. */
function parseIsoDate(input) {
  if (isDate(input)) {
    return { y: input.getFullYear(), m: input.getMonth() + 1, d: input.getDate() };
  }
  var match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(input || '').trim());
  if (!match) return null;
  var y = Number(match[1]), m = Number(match[2]), d = Number(match[3]);
  if (m < 1 || m > 12 || d < 1 || d > daysInMonth(y, m)) return null;
  return { y: y, m: m, d: d };
}

function daysInMonth(y, m) {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

/** Whole days since 1970-01-01, so two dates can be compared with plain numbers. */
function dayNumber(p) {
  return Math.round(Date.UTC(p.y, p.m - 1, p.d) / 86400000);
}

/** Every day from start to end inclusive, as {y, m, d}. */
function enumerateDays(start, end) {
  var out = [];
  var t = Date.UTC(start.y, start.m - 1, start.d);
  var endT = Date.UTC(end.y, end.m - 1, end.d);
  while (t <= endT) {
    var dt = new Date(t);
    out.push({ y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() });
    t += 86400000;
  }
  return out;
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

function isoString(p) {
  return p.y + '-' + pad2(p.m) + '-' + pad2(p.d);
}

/**
 * Financial-year label for a month.
 *   fyStart 1  -> '2026'
 *   fyStart 4  -> April 2026 to March 2027 is '2026-27'
 */
function fyLabel(y, m, fyStart) {
  if (fyStart === 1) return String(y);
  var startYear = m >= fyStart ? y : y - 1;
  return startYear + '-' + pad2((startYear + 1) % 100);
}

function monthKey(y, m) {
  return y * 100 + m;
}

function monthLabel(y, m) {
  return MONTH_NAMES[m - 1] + ' ' + y;
}

/** Accepts 1200, '1200', '1,200.50', ' 95 '. Returns a number or null. */
function parseNumber(v) {
  if (typeof v === 'number') return isNaN(v) ? null : v;
  var s = String(v === undefined || v === null ? '' : v).replace(/[,\s]/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  return Number(s);
}

// ---------------------------------------------------------------------------
// Grid rows
// ---------------------------------------------------------------------------

/** Sort order of the Grid sheet: customer, then vehicle, then month. */
function compareGridKey(a, b) {
  if (a.customer !== b.customer) return a.customer < b.customer ? -1 : 1;
  if (a.vehicle !== b.vehicle) return a.vehicle < b.vehicle ? -1 : 1;
  return a.monthKey - b.monthKey;
}

/**
 * Makes sure the Grid has one row for each month in `months` for this customer and vehicle,
 * inserting rows in sorted position when they are missing.
 *
 * @param {Sheet} sheet   the Grid sheet
 * @param {string} customer
 * @param {string} vehicle
 * @param {Array<{y:number, m:number, monthKey:number}>} months
 * @param {number} fyStart
 * @return {Object} monthKey -> row number
 */
function ensureGridRows(sheet, customer, vehicle, months, fyStart) {
  var existing = [];
  var last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, GRID.COL_MONTH_KEY).getValues().forEach(function (r, i) {
      existing.push({
        row: i + 2,
        customer: String(r[GRID.COL_CUSTOMER - 1]),
        vehicle: String(r[GRID.COL_VEHICLE - 1]),
        monthKey: Number(r[GRID.COL_MONTH_KEY - 1])
      });
    });
  }

  var rows = {};
  months.slice().sort(function (a, b) { return a.monthKey - b.monthKey; }).forEach(function (mo) {
    var target = { customer: customer, vehicle: vehicle, monthKey: mo.monthKey };
    var found = null;
    var insertAt = -1;
    for (var i = 0; i < existing.length; i++) {
      var c = compareGridKey(existing[i], target);
      if (c === 0) { found = existing[i]; break; }
      if (c > 0) { insertAt = i; break; }
    }
    if (found) {
      rows[mo.monthKey] = found.row;
      return;
    }

    // A leading apostrophe keeps the labels as text. Without it Sheets turns 'Aug 2026' into a
    // date and copies the format of a neighboring row, so labels drift between 'Aug' and 'August'.
    var values = ["'" + fyLabel(mo.y, mo.m, fyStart), customer, vehicle, "'" + monthLabel(mo.y, mo.m), mo.monthKey];
    var rowNum;
    if (insertAt === -1) {
      sheet.appendRow(values);
      rowNum = sheet.getLastRow();
      existing.push({ row: rowNum, customer: customer, vehicle: vehicle, monthKey: mo.monthKey });
    } else {
      rowNum = existing[insertAt].row;
      sheet.insertRowBefore(rowNum);
      sheet.getRange(rowNum, 1, 1, values.length).setValues([values]);
      existing.splice(insertAt, 0, { row: rowNum, customer: customer, vehicle: vehicle, monthKey: mo.monthKey });
      for (var j = insertAt + 1; j < existing.length; j++) existing[j].row += 1;
    }
    sheet.getRange(rowNum, GRID.COL_TOTAL).setFormulaR1C1('=SUM(RC[-' + GRID.DAYS + ']:RC[-1])');
    rows[mo.monthKey] = rowNum;
  });
  return rows;
}
