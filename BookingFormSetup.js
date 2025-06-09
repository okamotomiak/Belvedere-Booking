// BookingFormSetup.js
// Google Apps Script to create a booking form and handle submissions

/**
 * Creates a Google Form for booking rooms and sets up a trigger to handle
 * submissions. The form collects basic booking details and writes them to the
 * "Bookings" sheet if the requested time slot is available.
 */
function setupBookingForm() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.create('Room Booking Form');

  // Fetch room list from the Rooms sheet for the dropdown item
  var roomsSheet = spreadsheet.getSheetByName('Rooms');
  var roomsData = roomsSheet.getRange(2, 1, roomsSheet.getLastRow() - 1, 3).getValues();
  var roomChoices = roomsData.map(function(row) {
    return row[0] + ' - ' + row[2]; // ROOM_ID - Room Name
  });

  form.addListItem()
      .setTitle('Room')
      .setChoiceValues(roomChoices)
      .setRequired(true);

  form.addTextItem().setTitle('Customer Name').setRequired(true);
  form.addTextItem().setTitle('Customer Email').setRequired(true);
  form.addTextItem().setTitle('Customer Phone');
  form.addDateItem().setTitle('Start Date').setRequired(true);
  form.addDateItem().setTitle('End Date').setRequired(true);
  form.addTimeItem().setTitle('Start Time').setRequired(true);
  form.addTimeItem().setTitle('End Time').setRequired(true);
  form.addTextItem().setTitle('Guest Count');
  form.addParagraphTextItem().setTitle('Purpose');
  form.addParagraphTextItem().setTitle('Special Requests');

  // Create form submission trigger
  ScriptApp.newTrigger('handleBookingFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  Logger.log('Booking form created: ' + form.getEditUrl());
}

/**
 * Trigger function that runs when the booking form is submitted. It validates
 * the requested time slot and records the booking if available.
 */
function handleBookingFormSubmit(e) {
  var response = e.response;
  var items = response.getItemResponses();
  var data = {};

  items.forEach(function(item) {
    data[item.getItem().getTitle()] = item.getResponse();
  });

  // Extract room id from "ROOM_ID - Room Name"
  var roomInfo = data['Room'].split(' - ');
  var roomId = roomInfo[0];

  var startDate = Utilities.formatDate(new Date(data['Start Date']), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var endDate = Utilities.formatDate(new Date(data['End Date']), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var startTime = Utilities.formatDate(new Date(data['Start Time']), Session.getScriptTimeZone(), 'HH:mm');
  var endTime = Utilities.formatDate(new Date(data['End Time']), Session.getScriptTimeZone(), 'HH:mm');

  // Check availability
  if (!isTimeSlotAvailable(roomId, 'room', startDate, endDate, startTime, endTime)) {
    // Send an email notification about unavailability
    MailApp.sendEmail({
      to: data['Customer Email'],
      subject: 'Room Not Available',
      htmlBody: 'The room is not available for the selected time slot. Please try another time.'
    });
    return;
  }

  // Append booking to the Bookings sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  var bookingId = generateId('BK');

  var bookingRow = [
    bookingId,
    'room',
    roomId,
    data['Customer Name'],
    data['Customer Email'],
    data['Customer Phone'],
    new Date(startDate),
    new Date(endDate),
    startTime,
    endTime,
    data['Guest Count'],
    data['Purpose'],
    data['Special Requests'],
    '', // Total Cost
    'Pending', // Payment Status
    'Submitted', // Booking Status
    new Date(),
    new Date(),
    '' // Notes
  ];

  sheet.appendRow(bookingRow);

  MailApp.sendEmail({
    to: data['Customer Email'],
    subject: 'Booking Received',
    htmlBody: 'Your booking request has been recorded. We will confirm shortly.'
  });
}

