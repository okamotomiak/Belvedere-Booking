/**
 * Booking System - Base Sheets Setup
 * Creates all necessary sheets for a hierarchical booking system
 * Property > Buildings > Rooms with custom rules
 */

function setupBookingSystem() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing sheets except the default one
  const sheets = spreadsheet.getSheets();
  if (sheets.length > 1) {
    for (let i = 1; i < sheets.length; i++) {
      spreadsheet.deleteSheet(sheets[i]);
    }
  }
  
  // Rename the first sheet or create sheets
  try {
    spreadsheet.getSheetByName('Properties') || 
    spreadsheet.insertSheet('Properties');
  } catch (e) {
    sheets[0].setName('Properties');
  }
  
  // Create all necessary sheets
  createPropertiesSheet(spreadsheet);
  createBuildingsSheet(spreadsheet);
  createRoomsSheet(spreadsheet);
  createBookingsSheet(spreadsheet);
  createRulesSheet(spreadsheet);
  createConfigSheet(spreadsheet);
  
  Logger.log('Booking system setup complete!');
  SpreadsheetApp.getUi().alert('Booking system sheets created successfully!');
}

function createPropertiesSheet(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Properties') || 
                spreadsheet.insertSheet('Properties');
  
  // Clear existing content
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Property ID',
    'Property Name',
    'Description',
    'Address',
    'Contact Email',
    'Contact Phone',
    'Default Check-in Time',
    'Default Check-out Time',
    'Time Zone',
    'Status',
    'Created Date',
    'Modified Date'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Add sample data
  const sampleData = [
    ['PROP001', 'Mountain Retreat Center', 'A peaceful retreat center in the mountains', '123 Mountain View Rd', 'admin@mountainretreat.com', '555-0123', '15:00', '11:00', 'America/New_York', 'Active', new Date(), new Date()]
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

function createBuildingsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Buildings');
  
  const headers = [
    'Building ID',
    'Property ID',
    'Building Name',
    'Description',
    'Building Type',
    'Capacity',
    'Floor Count',
    'Amenities',
    'Booking Type', // 'whole_building' or 'individual_rooms'
    'Status',
    'Created Date',
    'Modified Date'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Add sample data
  const sampleData = [
    ['BLDG001', 'PROP001', 'The Barn', 'Rustic barn for events and gatherings', 'Event Hall', 100, 1, 'Kitchen, Sound System, Projector', 'whole_building', 'Active', new Date(), new Date()],
    ['BLDG002', 'PROP001', 'Community Center', 'Multi-purpose building with individual rooms', 'Multi-Purpose', 200, 2, 'Kitchen, WiFi, Parking', 'individual_rooms', 'Active', new Date(), new Date()],
    ['BLDG003', 'PROP001', 'Guest Lodge', 'Accommodation building', 'Lodging', 50, 2, 'WiFi, Laundry, Common Room', 'individual_rooms', 'Active', new Date(), new Date()]
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function createRoomsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Rooms');
  
  const headers = [
    'Room ID',
    'Building ID',
    'Room Name',
    'Room Number',
    'Description',
    'Room Type',
    'Capacity',
    'Floor',
    'Square Footage',
    'Amenities',
    'Hourly Rate',
    'Daily Rate',
    'Status',
    'Created Date',
    'Modified Date'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#fbbc04');
  headerRange.setFontColor('black');
  headerRange.setFontWeight('bold');
  
  // Add sample data
  const sampleData = [
    ['ROOM001', 'BLDG002', 'Room A', 'A', 'Small meeting room', 'Meeting Room', 8, 1, 200, 'Whiteboard, TV', 25, 200, 'Active', new Date(), new Date()],
    ['ROOM002', 'BLDG002', 'Room B', 'B', 'Medium conference room', 'Conference Room', 15, 1, 350, 'Projector, Conference Phone', 40, 320, 'Active', new Date(), new Date()],
    ['ROOM003', 'BLDG002', 'Room G', 'G', 'Large event space', 'Event Space', 50, 2, 800, 'Sound System, Stage, Kitchen Access', 75, 600, 'Active', new Date(), new Date()],
    ['ROOM004', 'BLDG003', 'Suite 1', '101', 'Two-bedroom guest suite', 'Guest Suite', 4, 1, 600, 'Kitchenette, Private Bath', 0, 150, 'Active', new Date(), new Date()],
    ['ROOM005', 'BLDG003', 'Suite 2', '102', 'One-bedroom guest suite', 'Guest Suite', 2, 1, 400, 'Kitchenette, Private Bath', 0, 120, 'Active', new Date(), new Date()]
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function createBookingsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Bookings');
  
  const headers = [
    'Booking ID',
    'Booking Type', // 'property', 'building', 'room'
    'Resource ID', // Property ID, Building ID, or Room ID
    'Customer Name',
    'Customer Email',
    'Customer Phone',
    'Start Date',
    'End Date',
    'Start Time',
    'End Time',
    'Guest Count',
    'Purpose',
    'Special Requests',
    'Total Cost',
    'Payment Status',
    'Booking Status',
    'Created Date',
    'Modified Date',
    'Notes'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Add sample booking
  const sampleData = [
    ['BK001', 'room', 'ROOM003', 'John Smith', 'john@example.com', '555-0456', new Date('2025-07-15'), new Date('2025-07-15'), '09:00', '17:00', 30, 'Corporate Workshop', 'Need tables and chairs setup', 600, 'Pending', 'Confirmed', new Date(), new Date(), 'First time customer']
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function createRulesSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Rules');
  
  const headers = [
    'Rule ID',
    'Resource Type', // 'property', 'building', 'room'
    'Resource ID',
    'Rule Type', // 'operating_hours', 'minimum_booking', 'maximum_booking', 'blackout_dates', 'pricing_rules'
    'Rule Name',
    'Rule Value', // JSON string or simple value
    'Days of Week', // 'Mon,Tue,Wed,Thu,Fri,Sat,Sun' or specific days
    'Start Date',
    'End Date',
    'Status',
    'Created Date',
    'Modified Date'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Add sample rules
  const sampleData = [
    ['RULE001', 'room', 'ROOM001', 'operating_hours', 'Business Hours Only', '{"start": "08:00", "end": "18:00"}', 'Mon,Tue,Wed,Thu,Fri', null, null, 'Active', new Date(), new Date()],
    ['RULE002', 'room', 'ROOM003', 'operating_hours', 'Extended Hours', '{"start": "06:00", "end": "22:00"}', 'Mon,Tue,Wed,Thu,Fri,Sat,Sun', null, null, 'Active', new Date(), new Date()],
    ['RULE003', 'building', 'BLDG001', 'minimum_booking', 'Minimum 4 Hours', '4', 'Mon,Tue,Wed,Thu,Fri,Sat,Sun', null, null, 'Active', new Date(), new Date()],
    ['RULE004', 'room', 'ROOM004', 'minimum_booking', 'Minimum 1 Night', '1', 'Mon,Tue,Wed,Thu,Fri,Sat,Sun', null, null, 'Active', new Date(), new Date()],
    ['RULE005', 'property', 'PROP001', 'blackout_dates', 'Annual Maintenance', '{"dates": ["2025-12-24", "2025-12-25", "2025-12-31", "2026-01-01"]}', null, new Date('2025-12-24'), new Date('2026-01-01'), 'Active', new Date(), new Date()]
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function createConfigSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Config');
  
  const headers = [
    'Setting Name',
    'Setting Value',
    'Description',
    'Data Type',
    'Modified Date'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#607d8b');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Add configuration settings
  const configData = [
    ['default_timezone', 'America/New_York', 'Default timezone for the booking system', 'string', new Date()],
    ['booking_lead_time_hours', '24', 'Minimum hours in advance for bookings', 'number', new Date()],
    ['max_booking_duration_days', '30', 'Maximum booking duration in days', 'number', new Date()],
    ['admin_email', 'admin@example.com', 'Administrator email for notifications', 'string', new Date()],
    ['currency', 'USD', 'Currency for pricing', 'string', new Date()],
    ['auto_confirm_bookings', 'false', 'Automatically confirm bookings without admin approval', 'boolean', new Date()],
    ['send_confirmation_emails', 'true', 'Send confirmation emails to customers', 'boolean', new Date()],
    ['booking_id_prefix', 'BK', 'Prefix for booking ID generation', 'string', new Date()]
  ];
  
  sheet.getRange(2, 1, configData.length, configData[0].length).setValues(configData);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

// Helper function to generate unique IDs
function generateId(prefix = 'ID') {
  const timestamp = new Date().getTime().toString(36);
  const random = Math.random().toString(36).substr(2, 5);
  return `${prefix}${timestamp}${random}`.toUpperCase();
}

// Helper function to get all bookings for a specific resource
function getBookingsForResource(resourceId, resourceType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const bookings = [];
  
  const resourceTypeIndex = headers.indexOf('Booking Type');
  const resourceIdIndex = headers.indexOf('Resource ID');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][resourceTypeIndex] === resourceType && data[i][resourceIdIndex] === resourceId) {
      bookings.push(data[i]);
    }
  }
  
  return bookings;
}

// Helper function to check if a time slot is available
function isTimeSlotAvailable(resourceId, resourceType, startDate, endDate, startTime, endTime) {
  const existingBookings = getBookingsForResource(resourceId, resourceType);
  
  // Convert dates and times for comparison
  const newStart = new Date(`${startDate} ${startTime}`);
  const newEnd = new Date(`${endDate} ${endTime}`);
  
  for (const booking of existingBookings) {
    const bookingStart = new Date(`${booking[6]} ${booking[8]}`); // Start Date + Start Time
    const bookingEnd = new Date(`${booking[7]} ${booking[9]}`);   // End Date + End Time
    
    // Check for overlap
    if (newStart < bookingEnd && newEnd > bookingStart) {
      return false; // Overlap found
    }
  }
  
  return true; // No overlap
}
