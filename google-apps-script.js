/**
 * Google Apps Script for ElevateGRE Email Collection
 * 
 * Instructions:
 * 1. Go to https://script.google.com/
 * 2. Create a new project
 * 3. Replace the default code with this script
 * 4. Save the project
 * 5. Deploy as web app with execute permissions for "Anyone"
 * 6. Copy the web app URL and replace 'YOUR_SCRIPT_ID' in index.html
 */

function doPost(e) {
  try {
    // Get the email from the form data
    const email = e.parameter.email;
    const timestamp = e.parameter.timestamp;
    
    // Validate email
    if (!email || !email.includes('@')) {
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: 'Invalid email'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get or create the spreadsheet
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    
    // Add headers if this is the first row
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 3).setValues([['Email', 'Timestamp', 'Status']]);
    }
    
    // Check if email already exists
    const data = sheet.getDataRange().getValues();
    const emailExists = data.some(row => row[0] === email);
    
    if (emailExists) {
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: 'Email already exists'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Add the new email to the spreadsheet
    sheet.appendRow([email, timestamp, 'Subscribed']);
    
    // Optional: Send confirmation email
    // sendConfirmationEmail(email);
    
    return ContentService
      .createTextOutput(JSON.stringify({success: true, message: 'Email added successfully'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error:', error);
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSpreadsheet() {
  const spreadsheetName = 'ElevateGRE Email Subscriptions';
  
  // Try to find existing spreadsheet
  const files = DriveApp.getFilesByName(spreadsheetName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  
  // Create new spreadsheet if it doesn't exist
  const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
  
  // Move to a specific folder (optional)
  // const folder = DriveApp.getFolderById('YOUR_FOLDER_ID');
  // DriveApp.getFileById(newSpreadsheet.getId()).moveTo(folder);
  
  return newSpreadsheet;
}

function sendConfirmationEmail(email) {
  try {
    const subject = 'Welcome to ElevateGRE - Early Access Confirmed!';
    const body = `
      Hi there!
      
      Thank you for subscribing to ElevateGRE early access!
      
      We're working hard to bring you the best GRE preparation platform. You'll be among the first to know when we launch.
      
      Stay tuned for updates!
      
      Best regards,
      The ElevateGRE Team
    `;
    
    GmailApp.sendEmail(email, subject, body);
  } catch (error) {
    console.error('Error sending confirmation email:', error);
  }
}

// Optional: Function to get all subscribers
function getAllSubscribers() {
  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Skip header row
  return data.slice(1).map(row => ({
    email: row[0],
    timestamp: row[1],
    status: row[2]
  }));
}

// Optional: Function to export subscribers to CSV
function exportSubscribersToCSV() {
  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Convert to CSV format
  const csvContent = data.map(row => row.join(',')).join('\n');
  
  // Create CSV file
  const blob = Utilities.newBlob(csvContent, 'text/csv', 'elevategre_subscribers.csv');
  
  return blob;
}
