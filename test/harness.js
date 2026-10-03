'use strict';
/**
 * Offline harness for the Apps Script project.
 *
 * Loads every src/*.gs file into a Node vm context whose globals are small in-memory stand-ins
 * for SpreadsheetApp, FormApp, ScriptApp, LockService, PropertiesService, Utilities, Session
 * and MailApp. Only the methods this project uses are implemented, with the same 1-based
 * row/column semantics as the real services, so the logic can be exercised without a Google
 * account. Nothing here talks to Google.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---------------------------------------------------------------------------
// Spreadsheet mocks
// ---------------------------------------------------------------------------

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const line = [];
      for (let c = 0; c < this.numCols; c++) line.push(this.sheet._get(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map(r => r.map(v => String(v))); }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(values) {
    if (values.length !== this.numRows || values.some(r => r.length !== this.numCols)) {
      throw new Error(`setValues: data is ${values.length}x${values[0] && values[0].length} but range is ${this.numRows}x${this.numCols}`);
    }
    values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
    return this;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  setFormula(f) { this.sheet._set(this.row, this.col, f); return this; }
  setFormulaR1C1(f) { this.sheet._set(this.row, this.col, f); return this; }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, '');
    return this;
  }
  setFontWeight() { return this; }
  setNumberFormat() { return this; }
  setBackground() { return this; }
  setHorizontalAlignment() { return this; }
  setWrap() { return this; }
}

class MockSheet {
  constructor(ss, name) {
    this.ss = ss;
    this.name = name;
    this.data = [];          // array of rows; each row an array of cell values ('' = empty)
    this.frozenRows = 0;
    this.frozenColumns = 0;
    this.hiddenColumns = [];
    this.columnWidths = {};
  }
  getName() { return this.name; }
  getParent() { return this.ss; }
  _get(r, c) {
    const row = this.data[r - 1];
    if (!row) return '';
    const v = row[c - 1];
    return v === undefined || v === null ? '' : v;
  }
  _set(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    const row = this.data[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  _isEmptyRow(row) { return !row || row.every(v => v === '' || v === null || v === undefined); }
  getLastRow() {
    let last = this.data.length;
    while (last > 0 && this._isEmptyRow(this.data[last - 1])) last--;
    return last;
  }
  getLastColumn() {
    let max = 0;
    this.data.forEach(row => {
      let c = row.length;
      while (c > 0 && (row[c - 1] === '' || row[c - 1] === null || row[c - 1] === undefined)) c--;
      if (c > max) max = c;
    });
    return max;
  }
  getMaxRows() { return Math.max(this.data.length, 1); }
  getMaxColumns() { return Math.max(this.getLastColumn(), 1); }
  getRange(row, col, numRows, numCols) {
    if (typeof row === 'string') throw new Error('A1 notation is not supported by the mock');
    return new MockRange(this, row, col, numRows === undefined ? 1 : numRows, numCols === undefined ? 1 : numCols);
  }
  getDataRange() { return new MockRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  appendRow(values) {
    const insertAt = this.getLastRow();            // append after the last row with content
    this.data.splice(insertAt, 0, values.slice());
    return this;
  }
  insertRowBefore(rowIndex) {
    while (this.data.length < rowIndex - 1) this.data.push([]);
    this.data.splice(rowIndex - 1, 0, []);
    return this;
  }
  insertRowAfter(rowIndex) { return this.insertRowBefore(rowIndex + 1); }
  deleteRows(rowPosition, howMany) { this.data.splice(rowPosition - 1, howMany); return this; }
  clearContents() { this.data = []; return this; }
  clear() { return this.clearContents(); }
  setFrozenRows(n) { this.frozenRows = n; return this; }
  setFrozenColumns(n) { this.frozenColumns = n; return this; }
  hideColumns(col, n) { for (let i = 0; i < (n || 1); i++) this.hiddenColumns.push(col + i); return this; }
  setColumnWidths(col, n, width) { for (let i = 0; i < n; i++) this.columnWidths[col + i] = width; return this; }
  setColumnWidth(col, width) { this.columnWidths[col] = width; return this; }
  autoResizeColumns() { return this; }
}

class MockSpreadsheet {
  constructor(id) {
    this.id = id || 'SPREADSHEET-ID';
    this.sheets = [];
    this.toasts = [];
  }
  getId() { return this.id; }
  getName() { return 'Mock spreadsheet'; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(name) { return this.sheets.find(s => s.name === name) || null; }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`A sheet named ${name} already exists`);
    const s = new MockSheet(this, name);
    this.sheets.push(s);
    return s;
  }
  deleteSheet(sheet) { this.sheets = this.sheets.filter(s => s !== sheet); }
  toast(message, title) { this.toasts.push({ message, title }); }
}

class MockUi {
  constructor() { this.menus = []; this.alerts = []; }
  createMenu(name) {
    const menu = { name, items: [] };
    const builder = {
      addItem: (label, fn) => { menu.items.push({ label, fn }); return builder; },
      addSeparator: () => builder,
      addSubMenu: () => builder,
      addToUi: () => { this.menus.push(menu); }
    };
    return builder;
  }
  alert(msg) { this.alerts.push(msg); }
}

// ---------------------------------------------------------------------------
// Form mocks
// ---------------------------------------------------------------------------

class MockItem {
  constructor(form, type) {
    this.form = form;
    this.type = type;
    this.title = '';
    this.helpText = '';
    this.required = false;
    this.choices = null;        // multiple choice: [{value, navigation}]
    this.choiceValues = null;   // list item: [value]
    this.goToPage = null;       // page break: item to jump to after the previous page
    this.validation = null;
  }
  setTitle(t) { this.title = t; return this; }
  getTitle() { return this.title; }
  setHelpText(t) { this.helpText = t; return this; }
  setRequired(b) { this.required = b; return this; }
  getType() { return this.type; }
  getIndex() { return this.form.items.indexOf(this); }
  createChoice(value, navigation) { return { value, navigation: navigation || null }; }
  setChoices(choices) { this.choices = choices; return this; }
  setChoiceValues(values) { this.choiceValues = values.slice(); return this; }
  setGoToPage(page) { this.goToPage = page; return this; }
  setValidation(v) { this.validation = v; return this; }
}

class MockForm {
  constructor(id, title) {
    this.id = id;
    this.title = title;
    this.description = '';
    this.items = [];
    this.destination = null;
  }
  getId() { return this.id; }
  setTitle(t) { this.title = t; return this; }
  getTitle() { return this.title; }
  setDescription(d) { this.description = d; return this; }
  setDestination(type, id) { this.destination = { type, id }; return this; }
  getItems() { return this.items.slice(); }
  deleteItem(itemOrIndex) {
    const idx = typeof itemOrIndex === 'number' ? itemOrIndex : this.items.indexOf(itemOrIndex);
    if (idx < 0) throw new Error('Item not found');
    this.items.splice(idx, 1);
  }
  getPublishedUrl() { return `https://docs.google.com/forms/d/e/${this.id}/viewform`; }
  getEditUrl() { return `https://docs.google.com/forms/d/${this.id}/edit`; }
  _add(type) { const it = new MockItem(this, type); this.items.push(it); return it; }
  addMultipleChoiceItem() { return this._add('MULTIPLE_CHOICE'); }
  addListItem() { return this._add('LIST'); }
  addPageBreakItem() { return this._add('PAGE_BREAK'); }
  addDateItem() { return this._add('DATE'); }
  addTextItem() { return this._add('TEXT'); }
  addParagraphTextItem() { return this._add('PARAGRAPH_TEXT'); }
}

/** Builds a fake FormResponse like the one the trigger receives. answers: {title: value}. */
function makeFormResponse(id, timestamp, answers) {
  return {
    getId: () => id,
    getTimestamp: () => timestamp,
    getItemResponses: () => Object.keys(answers).map(title => ({
      getItem: () => ({ getTitle: () => title }),
      getResponse: () => answers[title]
    }))
  };
}

// ---------------------------------------------------------------------------
// Service mocks and loader
// ---------------------------------------------------------------------------

function pad(n) { return (n < 10 ? '0' : '') + n; }

function createEnvironment() {
  const ss = new MockSpreadsheet();
  const ui = new MockUi();
  const forms = {};
  let formCounter = 0;
  const triggers = [];
  const properties = {};
  const mail = [];
  const logs = [];

  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ss,
    getActive: () => ss,
    getUi: () => ui,
    flush: () => {}
  };

  const FormApp = {
    DestinationType: { SPREADSHEET: 'SPREADSHEET' },
    create: title => {
      const form = new MockForm(`FORM-${++formCounter}`, title);
      forms[form.id] = form;
      return form;
    },
    openById: id => {
      if (!forms[id]) throw new Error(`No form with id ${id}`);
      return forms[id];
    },
    createTextValidation: () => {
      const rule = {};
      const builder = {
        requireNumberGreaterThanOrEqualTo: n => { rule.min = n; return builder; },
        requireNumber: () => { rule.number = true; return builder; },
        build: () => rule
      };
      return builder;
    }
  };

  const ScriptApp = {
    newTrigger: fn => {
      const spec = { handler: fn };
      const builder = {
        forForm: f => { spec.formId = typeof f === 'string' ? f : f.getId(); return builder; },
        forSpreadsheet: s => { spec.spreadsheetId = typeof s === 'string' ? s : s.getId(); return builder; },
        onFormSubmit: () => { spec.event = 'ON_FORM_SUBMIT'; return builder; },
        create: () => {
          const trigger = { getHandlerFunction: () => spec.handler, spec };
          triggers.push(trigger);
          return trigger;
        }
      };
      return builder;
    },
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: t => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); }
  };

  const LockService = {
    getScriptLock: () => ({ waitLock: () => {}, tryLock: () => true, releaseLock: () => {}, hasLock: () => true })
  };

  const PropertiesService = {
    getScriptProperties: () => ({
      getProperty: k => (k in properties ? properties[k] : null),
      setProperty: (k, v) => { properties[k] = String(v); },
      deleteProperty: k => { delete properties[k]; }
    })
  };

  const Utilities = {
    formatDate: (date, tz, format) => {
      const d = new Date(date);
      return format
        .replace('yyyy', d.getUTCFullYear())
        .replace('MM', pad(d.getUTCMonth() + 1))
        .replace('dd', pad(d.getUTCDate()))
        .replace('HH', pad(d.getUTCHours()))
        .replace('mm', pad(d.getUTCMinutes()))
        .replace('ss', pad(d.getUTCSeconds()));
    },
    sleep: () => {}
  };

  const Session = { getScriptTimeZone: () => 'Etc/UTC' };
  const MailApp = { sendEmail: (to, subject, body) => { mail.push({ to, subject, body }); } };
  const Logger = { log: msg => logs.push(String(msg)) };
  const quietConsole = { log: msg => logs.push(String(msg)), error: msg => logs.push('ERROR ' + String(msg)), warn: msg => logs.push('WARN ' + String(msg)) };

  const context = vm.createContext({
    SpreadsheetApp, FormApp, ScriptApp, LockService, PropertiesService, Utilities, Session, MailApp, Logger,
    console: quietConsole
  });

  const srcDir = path.join(__dirname, '..', 'src');
  fs.readdirSync(srcDir).filter(f => f.endsWith('.gs')).sort().forEach(f => {
    vm.runInContext(fs.readFileSync(path.join(srcDir, f), 'utf8'), context, { filename: f });
  });

  return {
    script: context,          // call Apps Script functions as script.setupSheets(), etc.
    ss, ui, forms, triggers, properties, mail, logs,
    makeFormResponse,
    sheet: name => ss.getSheetByName(name),
    rows: name => { const s = ss.getSheetByName(name); return s ? s.data.slice(0, s.getLastRow()) : []; }
  };
}

module.exports = { createEnvironment, makeFormResponse, MockSheet, MockSpreadsheet, MockForm };
