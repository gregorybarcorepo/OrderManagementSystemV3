/**
 * ============================================
 * FORM SUBMISSION HANDLER - FINAL VERSION
 * Handles web form submissions via doPost()
 * 
 * CRITICAL RULE: Backend NEVER generates confirmation numbers
 * Frontend generates and sends confirmation number in formData
 * ============================================
 */

// Sheet name mapping
const SHEET_NAMES = {
  "target-order": "TargetOrders",
  "amazon-order": "AmazonOrders",
  "document-submission": "DocumentSubmissions",
  "campus-club-space-feedback": "CampusClubSpaceFeedback"
};

/**
 * Main entry point for web form submissions
 * @param {Object} e - Event object from POST request
 * @returns {ContentService.TextOutput} JSON response
 */
function doPost(e) {
  const startTime = Date.now();
  console.log('🚀 Form submission received');
  
  try {
    // Parse form data
    const formData = JSON.parse(e.postData.contents);
    console.log('📝 Form type:', formData.form_type);
    
    // CRITICAL: Verify frontend sent confirmation number
    if (!formData.Confirmation_Number || formData.Confirmation_Number.trim() === '') {
      throw new Error('Frontend must send confirmation number');
    }
    
    const confirmationNumber = formData.Confirmation_Number;
    console.log('✅ Using frontend confirmation number:', confirmationNumber);
    
    // Detect form type if missing
    if (!formData.form_type || !SHEET_NAMES[formData.form_type]) {
      formData.form_type = detectFormTypeFromData(formData);
    }
    
    // Process files if present
    if (formData.fileData) {
      console.log('📁 Processing files...');
      processFilesAndAddLinksToFormData(formData);
    }
    
    // Write to sheet (Confirmation_Number already in formData)
    console.log('💾 Writing to sheet...');
    const writeResult = writeToSheet(formData);
    
    if (!writeResult.success) {
      throw new Error('Failed to write: ' + writeResult.error);
    }
    
    // Assign IDs
    const sheet = getCachedSheet(writeResult.sheetName);
    const idResult = assignIDsToRow(sheet, writeResult.row);
    
    // Send email using frontend's confirmation number
    console.log('📧 Sending email with confirmation:', confirmationNumber);
    const emailResult = sendConfirmationEmail(
      formData, 
      formData.form_type, 
      confirmationNumber
    );
    
    const executionTime = Date.now() - startTime;
    console.log(`✅ Completed in ${executionTime}ms`);
    
    // Return same confirmation number frontend sent
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        confirmationNumber: confirmationNumber,
        emailSent: emailResult.sent,
        historicalId: idResult.assignedIds?.Historical_Submission_Id,
        timestamp: new Date().toLocaleString('en-US', { timeZone: 'America/New_York' })
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('❌ Error:', error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: error.toString(),
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========================================
// FORM TYPE DETECTION
// ========================================

function detectFormTypeFromData(formData) {
  const keys = Object.keys(formData);
  
  if (keys.includes('feedback_type') || keys.includes('space_rating')) {
    return "campus-club-space-feedback";
  }
  
  if (keys.includes('First_Name') && keys.includes('Last_Name') && 
      (keys.includes('nonPO') || keys.includes('signIn'))) {
    return "document-submission";
  }
  
  if (keys.includes('Wishlist_Link') || keys.includes('wishlist_link')) {
    return "amazon-order";
  }
  
  if (keys.includes('First_Item_Url') || keys.some(key => key.includes('Item_Url'))) {
    return "target-order";
  }
  
  console.log('⚠️ Could not detect form type, using target-order as default');
  return "target-order";
}

// ========================================
// SHEET WRITING
// ========================================

function writeToSheet(formData) {
  try {
    const formType = formData.form_type;
    const sheetName = SHEET_NAMES[formType];
    
    if (!sheetName) {
      throw new Error(`Unknown form type: ${formType}`);
    }
    
    const sheet = getCachedSheet(sheetName);
    
    // Get existing headers
    const lastColumn = sheet.getLastColumn();
    const headers = lastColumn > 0 ? 
      sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : 
      [];
    
    const newHeaders = headers.length > 0 ? [...headers] : ['Timestamp'];
    let columnsAdded = 0;
    
    // Ensure Confirmation_Number column exists
    if (!newHeaders.includes('Confirmation_Number')) {
      newHeaders.push('Confirmation_Number');
      columnsAdded++;
    }
    
    // Ensure ID columns exist
    ['Historical_Submission_Id', 'Semester_Id', 'Submission_Id'].forEach(col => {
      if (!newHeaders.includes(col)) {
        newHeaders.push(col);
        columnsAdded++;
      }
    });
    
    // Add file columns
    const expectedFileColumns = getExpectedFileColumnsForFormType(formType);
    expectedFileColumns.forEach(columnName => {
      if (!newHeaders.includes(columnName) && formData[columnName]) {
        newHeaders.push(columnName);
        columnsAdded++;
      }
    });
    
    // Add form field columns
    Object.keys(formData).forEach(key => {
      if (!newHeaders.includes(key) && 
          key !== 'fileData' && 
          key !== 'form_type' && 
          key !== 'sheet_target' &&
          key !== 'timestamp') {
        
        const value = formData[key];
        
        // Always include Event_Date
        if (key === 'Event_Date' || key === 'event_date' || key === 'Event Date') {
          newHeaders.push(key);
          columnsAdded++;
          return;
        }
        
        // Handle HTML date format
        if (isHTMLDateFormat(value)) {
          newHeaders.push(key);
          columnsAdded++;
          return;
        }
        
        // Standard fields
        if (value !== null && value !== undefined && value !== "" && typeof value !== 'object') {
          newHeaders.push(key);
          columnsAdded++;
        }
      }
    });
    
    // Update headers if needed
    if (columnsAdded > 0 || headers.length === 0) {
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    }
    
    // Build row data
    const rowData = newHeaders.map(header => {
      if (header === 'Timestamp') {
        return Utilities.formatDate(new Date(), 'America/New_York', 'MM/dd/yyyy hh:mm:ss a');
      }
      
      if (header === 'Historical_Submission_Id' || 
          header === 'Semester_Id' || 
          header === 'Submission_Id') {
        return ''; // Will be filled by assignIDsToRow
      }
      
      let value = formData[header];
      
      if (isHTMLDateFormat(value)) {
        return formatHTMLDate(value);
      }
      
      if (value === null || value === undefined) {
        return "";
      }
      
      if (typeof value === 'object' && !(value instanceof Date)) {
        return "";
      }
      
      return value;
    });
    
    // Write row
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    console.log(`✅ Written to row ${nextRow} in ${sheetName}`);
    
    return {
      success: true,
      row: nextRow,
      sheetName: sheetName
    };
    
  } catch (error) {
    console.error('❌ Sheet writing failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}
