/**
 * ElevateGRE Email Collection - Google Apps Script
 * 
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com/
 * 2. Click "New Project"
 * 3. Delete all existing code
 * 4. Copy and paste this entire script
 * 5. Save the project (Ctrl+S or Cmd+S)
 * 6. Click "Deploy" > "New Deployment"
 * 7. Choose "Web app" as the type
 * 8. Set "Execute as" to "Me"
 * 9. Set "Who has access" to "Anyone"
 * 10. Click "Deploy"
 * 11. Copy the web app URL and replace 'YOUR_SCRIPT_ID' in your HTML file
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
      sheet.getRange(1, 1, 1, 4).setValues([['Email', 'Timestamp', 'Status', 'Date Added']]);
      // Format header row
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      sheet.getRange(1, 1, 1, 4).setBackground('#5170ff');
      sheet.getRange(1, 1, 1, 4).setFontColor('white');
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
    const currentDate = new Date();
    sheet.appendRow([email, timestamp, 'Subscribed', currentDate]);
    
    // Optional: Send confirmation email (uncomment if you want this)
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
  
  // Optional: Move to a specific folder
  // const folder = DriveApp.getFolderById('YOUR_FOLDER_ID');
  // DriveApp.getFileById(newSpreadsheet.getId()).moveTo(folder);
  
  return newSpreadsheet;
}

function sendConfirmationEmail(email) {
  try {
    const subject = 'Welcome to ElevateGRE - Early Access Confirmed! 🎉';
    const body = `
Hi there!

Thank you for subscribing to ElevateGRE early access!

We're working hard to bring you the best GRE preparation platform. You'll be among the first to know when we launch.

What to expect:
✅ Personalized study plans
✅ Expert-led video lessons
✅ Practice tests and analytics
✅ 24/7 support

Stay tuned for updates!

Best regards,
The ElevateGRE Team

---
Follow us:
LinkedIn: https://linkedin.com/company/elevate-gre/
Twitter: https://twitter.com/elevategre
Instagram: https://instagram.com/elevate_gre
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
    status: row[2],
    dateAdded: row[3]
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

// Optional: Function to send bulk updates to all subscribers
function sendBulkUpdate(subject, message) {
  const subscribers = getAllSubscribers();
  
  subscribers.forEach(subscriber => {
    try {
      GmailApp.sendEmail(subscriber.email, subject, message);
      Utilities.sleep(1000); // Wait 1 second between emails to avoid rate limits
    } catch (error) {
      console.error(`Error sending to ${subscriber.email}:`, error);
    }
  });
}

// Optional: Function to get subscriber count
function getSubscriberCount() {
  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  return Math.max(0, sheet.getLastRow() - 1); // Subtract 1 for header row
}

// Test function - you can run this to test the script
function testScript() {
  const testData = {
    parameter: {
      email: 'test@example.com',
      timestamp: new Date().toISOString()
    }
  };
  
  const result = doPost(testData);
  console.log('Test result:', result.getContent());
}
