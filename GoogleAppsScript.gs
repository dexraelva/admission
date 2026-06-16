// AIMTECH CHRISTIAN COLLEGE - ADMISSION FORM GOOGLE APPS SCRIPT
// This script handles form submissions, stores data in Google Sheets, and manages file uploads

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; // Replace with your Google Sheet ID
const SHEET_NAME = 'Applications'; // Name of the sheet to store data
const UPLOAD_FOLDER_ID = 'YOUR_FOLDER_ID'; // Replace with Google Drive folder ID for file uploads

// Main POST handler - receives form data
function doPost(e) {
    try {
        const formData = e.parameter;
        
        // Log the data (for debugging)
        console.log('Form data received:', formData);
        
        // Store form data in Google Sheets
        const result = storeFormData(formData);
        
        // Handle file uploads if any
        if (e.contents) {
            handleFileUploads(e, formData.email);
        }
        
        return ContentService.createTextOutput(JSON.stringify({
            status: 'success',
            message: 'Application submitted successfully',
            id: result
        })).setMimeType(ContentService.MimeType.JSON);
        
    } catch (error) {
        console.error('Error processing form:', error);
        return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

// Store form data in Google Sheets
function storeFormData(formData) {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        
        // Get headers from first row
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        // Prepare data row
        const row = [];
        const timestamp = new Date();
        
        // Add timestamp and basic info
        row.push(timestamp);
        row.push(formData.fullname || '');
        row.push(formData.email || '');
        row.push(formData.mobile || '');
        row.push(formData.program || '');
        row.push(formData.churchName || '');
        row.push(formData.pastorName || '');
        row.push(formData.dob || '');
        row.push(formData.gender || '');
        row.push(formData.presentAddress || '');
        row.push(formData.provincialAddress || '');
        row.push(formData.citizenship || '');
        row.push(formData.civilStatus || '');
        row.push(formData.spouseName || '');
        row.push(formData.dateAccepted || '');
        row.push(formData.dateBaptized || '');
        row.push(formData.baptizedBy || '');
        row.push(formData.denomination || '');
        row.push(formData.testimonyJesus || '');
        row.push(formData.callingDescription || '');
        row.push(formData.membershipYears || '');
        row.push(formData.emergencyName || '');
        row.push(formData.emergencyContact || '');
        
        // Add application date
        row.push(formData.applicationDate || '');
        
        // Add status
        row.push('Pending');
        
        // Append row to sheet
        sheet.appendRow(row);
        
        // Send notification email to admin (optional)
        sendAdminNotification(formData);
        
        return timestamp.getTime();
        
    } catch (error) {
        console.error('Error storing form data:', error);
        throw error;
    }
}

// Send notification email to admin
function sendAdminNotification(formData) {
    try {
        const adminEmail = 'admin@aimtechseminary.org'; // Replace with admin email
        const subject = `New Admission Application - ${formData.fullname}`;
        
        const body = `
New Admission Application Received

Applicant Name: ${formData.fullname}
Email: ${formData.email}
Phone: ${formData.mobile}
Program Applied: ${formData.program}
Church: ${formData.churchName}
Submission Date: ${new Date().toString()}

Please review the application in your Google Sheet.
        `;
        
        MailApp.sendEmail(adminEmail, subject, body);
        
    } catch (error) {
        console.error('Error sending admin notification:', error);
    }
}

// Handle file uploads
function handleFileUploads(e, applicantEmail) {
    try {
        const uploadFolder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
        const applicantFolder = uploadFolder.createFolder(`${applicantEmail}_${new Date().getTime()}`);
        
        // Get all uploaded files from the multipart form data
        const blobs = e.parameter.files;
        
        if (blobs && Array.isArray(blobs)) {
            blobs.forEach((blob, index) => {
                applicantFolder.createFile(blob);
            });
        }
        
        return applicantFolder.getUrl();
        
    } catch (error) {
        console.error('Error uploading files:', error);
        // Don't throw - file upload failure shouldn't block form submission
        return null;
    }
}

// Setup function - create initial sheet structure (run once)
function setupSheet() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        
        const headers = [
            'Timestamp',
            'Full Name',
            'Email',
            'Mobile',
            'Program',
            'Church Name',
            'Pastor Name',
            'Date of Birth',
            'Gender',
            'Present Address',
            'Provincial Address',
            'Citizenship',
            'Civil Status',
            'Spouse Name',
            'Date Accepted Christ',
            'Date Baptized',
            'Baptized By',
            'Denomination',
            'Testimony of Jesus',
            'Calling Description',
            'Church Membership Years',
            'Emergency Contact Name',
            'Emergency Contact',
            'Application Date',
            'Status'
        ];
        
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        sheet.getRange(1, 1, 1, headers.length).setBackground('#0A4DB3');
        sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
        
        console.log('Sheet setup complete');
        
    } catch (error) {
        console.error('Error setting up sheet:', error);
    }
}

// Test function to verify setup
function test() {
    console.log('Google Apps Script is deployed and working!');
    console.log('Spreadsheet ID:', SPREADSHEET_ID);
    console.log('Upload Folder ID:', UPLOAD_FOLDER_ID);
}
