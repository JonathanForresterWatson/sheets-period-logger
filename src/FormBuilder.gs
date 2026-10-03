/**
 * Builds the Google Form from the Customers sheet and installs the submit trigger.
 *
 * Form structure:
 *   Page 1   Customer (multiple choice). Each choice jumps to that customer's page.
 *   Page 2+  One page per active customer with a dropdown of that customer's vehicles.
 *            After it, the form jumps straight to the Period page, skipping the other customers.
 *   Last     Period: start date, end date, daily value, note.
 *
 * Rebuilding keeps the same form (and its URL) and replaces the questions in place.
 */

var FORM = {
  CUSTOMER_TITLE: 'Customer',
  VEHICLE_PREFIX: 'Vehicle (',
  PERIOD_TITLE: 'Period',
  START_TITLE: 'Start date',
  END_TITLE: 'End date',
  VALUE_TITLE: 'Daily value',
  NOTE_TITLE: 'Note'
};

function buildForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig(ss);
  var data = readCustomers(ss);
  if (data.customers.length === 0) {
    throw new Error('No active customers found. Fill the Customers sheet first (or run addDemoCustomers).');
  }

  var form = cfg.FORM_ID ? openFormOrNull(cfg.FORM_ID) : null;
  if (form) {
    clearFormItems(form);
  } else {
    form = FormApp.create(cfg.FORM_TITLE);
    setConfigValue(ss, 'FORM_ID', form.getId());
  }

  form.setTitle(cfg.FORM_TITLE)
    .setDescription('Pick the customer, then the vehicle, then the period. The daily value is applied to every day in the period.');
  try {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  } catch (e) {
    // already linked to this spreadsheet
  }

  // Page 1: customer choice. Choices get their navigation after the pages exist.
  var customerItem = form.addMultipleChoiceItem()
    .setTitle(FORM.CUSTOMER_TITLE)
    .setRequired(true);

  // One page per customer.
  var pages = data.customers.map(function (c) {
    var page = form.addPageBreakItem().setTitle(c.name);
    form.addListItem()
      .setTitle(FORM.VEHICLE_PREFIX + c.name + ')')
      .setChoiceValues(c.vehicles)
      .setRequired(true);
    return page;
  });

  // Shared period page.
  var periodPage = form.addPageBreakItem().setTitle(FORM.PERIOD_TITLE);
  form.addDateItem().setTitle(FORM.START_TITLE).setRequired(true);
  form.addDateItem().setTitle(FORM.END_TITLE).setRequired(true);
  form.addTextItem()
    .setTitle(FORM.VALUE_TITLE)
    .setHelpText('A number. It is written into every day of the period.')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build());
  form.addParagraphTextItem().setTitle(FORM.NOTE_TITLE);

  // Navigation. A choice sends the respondent to its customer page. A page break's
  // setGoToPage() controls where the respondent goes after finishing the page BEFORE it,
  // so every customer page except the last gets sent to the period page by the page break
  // that follows it. The last customer page reaches the period page in normal order.
  customerItem.setChoices(data.customers.map(function (c, i) {
    return customerItem.createChoice(c.name, pages[i]);
  }));
  for (var i = 1; i < pages.length; i++) {
    pages[i].setGoToPage(periodPage);
  }

  setConfigValue(ss, 'FORM_URL', form.getPublishedUrl());
  say(ss, 'Form ready with ' + data.customers.length + ' customer page(s): ' + form.getPublishedUrl());
  return form;
}

function openFormOrNull(id) {
  try {
    return FormApp.openById(id);
  } catch (e) {
    return null; // deleted or not accessible; a new form will be created
  }
}

function clearFormItems(form) {
  var items = form.getItems();
  for (var i = items.length - 1; i >= 0; i--) {
    form.deleteItem(items[i]);
  }
}

/** Replaces any existing onFormSubmit trigger with one bound to the current form. */
function installTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig(ss);
  if (!cfg.FORM_ID) throw new Error('No form yet. Run buildForm() first.');

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(cfg.FORM_ID)
    .onFormSubmit()
    .create();
  say(ss, 'Submit trigger installed.');
}
