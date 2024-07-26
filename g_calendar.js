function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📌 Menu')
    .addItem('Update Calendar', 'managePaymentsCalendar')
    .addToUi();
}

const SHEET_NAME = "Regular Payments";
const PAYEE_COL = 0;
const AMOUNT_COL = 2;
const DATE_COL = 6;
const TYPE_COL = 1;
const ACC_COL = 4;
const PUR_COL = 3;
const ROWHEAD = 4;

function managePaymentsCalendar() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  let calendar = getOrCreateCalendar();
  processPaymentEntries(data, calendar);
  alertUser();
}

function getOrCreateCalendar() {
  let calendarId = PropertiesService.getUserProperties().getProperty('calendarId');
  let calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : null;

  if (!calendar) {
    calendar = CalendarApp.createCalendar("Regular Payments");
    PropertiesService.getUserProperties().setProperty('calendarId', calendar.getId());
  }

  return calendar;
}

function processPaymentEntries(data, calendar) {
  for (let i = ROWHEAD; i < data.length; i++) {
    const [payee, amount, dueDate, type, accountNumber, purpose] = [
      data[i][PAYEE_COL],
      data[i][AMOUNT_COL],
      new Date(data[i][DATE_COL]),
      data[i][TYPE_COL],
      data[i][ACC_COL],  
      data[i][PUR_COL]  
    ];
    
    if (!isNaN(dueDate.getTime())) {
      createOrUpdateEvent(calendar, payee, amount, dueDate, type, accountNumber, purpose);
    }
  }
}

function createOrUpdateEvent(calendar, payee, amount, dueDate, type, accountNumber, purpose) {
  const eventTitle = `Payment to ${payee}: $${amount}`;
  const eventDescription = `Type: ${type}, Account: ${accountNumber}, Purpose: ${purpose}`;
  const events = calendar.getEventsForDay(dueDate, {search: eventTitle});

  if (events.length === 0) {
    const recurrence = CalendarApp.newRecurrence().addMonthlyRule().until(new Date(dueDate.getFullYear() + 10, dueDate.getMonth(), dueDate.getDate()));
    const event = calendar.createAllDayEventSeries(eventTitle, dueDate, recurrence, {description: eventDescription});
    event.setColor("3"); 
  }
}

function alertUser() {
  SpreadsheetApp.getUi().alert('Your Calendar Has Been Updated with Monthly Payments.');
}

