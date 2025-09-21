/**
 * ElevateGRE Email Collection - FIXED VERSION
 * 
 * This version includes better error handling and debugging
 */

function doPost(e) {
  try {
    // Log the incoming request for debugging
    console.log('Received request:', e);
    
    // Get the email from the form data
    const email = e.parameter.email;
    const timestamp = e.parameter.timestamp;
    
    console.log('Email:', email);
    console.log('Timestamp:', timestamp);
    
    // Validate email
    if (!email || !email.includes('@')) {
      console.log('Invalid email provided');
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: 'Invalid email'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get or create the spreadsheet
    const spreadsheet = getOrCreateSpreadsheet();
    console.log('Spreadsheet ID:', spreadsheet.getId());
    
    const sheet = spreadsheet.getActiveSheet();
    console.log('Sheet name:', sheet.getName());
    
    // Add headers if this is the first row
    if (sheet.getLastRow() === 0) {
      console.log('Adding headers to new sheet');
      sheet.getRange(1, 1, 1, 4).setValues([['Email', 'Timestamp', 'Status', 'Date Added']]);
      // Format header row
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      sheet.getRange(1, 1, 1, 4).setBackground('#5170ff');
      sheet.getRange(1, 1, 1, 4).setFontColor('white');
    }
    
    // Check if email already exists
    const data = sheet.getDataRange().getValues();
    console.log('Current data rows:', data.length);
    
    const emailExists = data.some(row => row[0] === email);
    console.log('Email exists:', emailExists);
    
    if (emailExists) {
      console.log('Email already exists, not adding');
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: 'Email already exists'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Add the new email to the spreadsheet
    const currentDate = new Date();
    console.log('Adding new email to sheet');
    sheet.appendRow([email, timestamp, 'Subscribed', currentDate]);
    
    console.log('Email added successfully');
    
    return ContentService
      .createTextOutput(JSON.stringify({success: true, message: 'Email added successfully'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error in doPost:', error);
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSpreadsheet() {
  const spreadsheetName = 'ElevateGRE Email Subscriptions';
  
  try {
    // Try to find existing spreadsheet
    const files = DriveApp.getFilesByName(spreadsheetName);
    if (files.hasNext()) {
      const file = files.next();
      console.log('Found existing spreadsheet:', file.getName());
      return SpreadsheetApp.open(file);
    }
    
    // Create new spreadsheet if it doesn't exist
    console.log('Creating new spreadsheet');
    const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
    console.log('New spreadsheet created with ID:', newSpreadsheet.getId());
    return newSpreadsheet;
    
  } catch (error) {
    console.error('Error in getOrCreateSpreadsheet:', error);
    throw error;
  }
}

// Test function - run this to test if the script works
function testScript() {
  console.log('Testing script...');
  
  const testData = {
    parameter: {
      email: 'test@example.com',
      timestamp: new Date().toISOString()
    }
  };
  
  try {
    const result = doPost(testData);
    console.log('Test result:', result.getContent());
    return result;
  } catch (error) {
    console.error('Test failed:', error);
    return null;
  }
}

// Function to check if the script has proper permissions
function checkPermissions() {
  try {
    // Try to access Drive
    const files = DriveApp.getFiles();
    console.log('Drive access: OK');
    
    // Try to create a test spreadsheet
    const testSheet = SpreadsheetApp.create('Test Permission Check');
    console.log('Spreadsheet creation: OK');
    
    // Clean up test file
    DriveApp.getFileById(testSheet.getId()).setTrashed(true);
    console.log('Test file cleaned up');
    
    return true;
  } catch (error) {
    console.error('Permission check failed:', error);
    return false;
  }
}

// Function to manually add an email (for testing)
function addTestEmail() {
  const email = 'manual-test@example.com';
  const timestamp = new Date().toISOString();
  
  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  
  // Add headers if needed
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['Email', 'Timestamp', 'Status', 'Date Added']]);
  }
  
  // Add the email
  sheet.appendRow([email, timestamp, 'Manual Test', new Date()]);
  console.log('Manual test email added');
}
