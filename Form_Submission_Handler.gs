/**
 * ============================================
 * FORM SUBMISSION HANDLER
 * Handles web form submissions via doPost()
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
    console.log('📝 Form data keys:', Object.keys(formData));
    console.log('📝 Form type received:', formData.form_type);
    
    // CRITICAL: Detect form type if missing or incorrect
    if (!formData.form_type || !SHEET_NAMES[formData.form_type]) {
      formData.form_type = detectFormTypeFromData(formData);
      console.log('🔍 Auto-detected form type:', formData.form_type);
    }
    
    // Process files BEFORE writing to sheet
    if (formData.fileData) {
      console.log('📁 Processing files...');
      processFilesAndAddLinksToFormData(formData);
    }
    
    // Write to sheet
    console.log('💾 Writing to sheet...');
    const writeResult = writeToSheet(formData);
    
    if (!writeResult.success) {
      throw new Error('Failed to write to sheet: ' + writeResult.error);
    }
    
    // ✨ CRITICAL: Assign IDs immediately after writing
    console.log('🎯 Assigning IDs...');
    const sheet = getCachedSheet(writeResult.sheetName);
    const idResult = assignIDsToRow(sheet, writeResult.row);
    
    if (!idResult.success) {
      console.error('⚠️ Warning: ID assignment failed:', idResult.error);
    }
    
    // Generate confirmation number
    const clubName = formData.Student_Organization || 
                     formData.Organization_Name || 
                     'UNKNOWN';
    const confirmationNumber = generateConfirmationNumber(clubName);
    
    // Add confirmation number to sheet
    addConfirmationNumberToSheet(sheet, writeResult.row, confirmationNumber);
    
    // Send confirmation email
    console.log('📧 Sending confirmation email...');
    const emailResult = sendConfirmationEmail(
      formData, 
      formData.form_type, 
      confirmationNumber
    );
    
    const executionTime = Date.now() - startTime;
    console.log(`✅ Submission completed in ${executionTime}ms`);
    
    // Return success response
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
    console.error('❌ Submission failed:', error.toString());
    
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

/**
 * Auto-detect form type from form data fields
 * @param {Object} formData - Form data object
 * @returns {string} Detected form type
 */
function detectFormTypeFromData(formData) {
  const keys = Object.keys(formData);
  
  // Check for feedback form fields
  if (keys.includes('feedback_type') || 
      keys.includes('space_rating') || 
      keys.includes('overall_satisfaction') ||
      keys.includes('Campus_Club_Space_Feedback')) {
    return "campus-club-space-feedback";
  }
  
  // Check for document submission fields
  if (keys.includes('First_Name') && keys.includes('Last_Name') && 
      (keys.includes('nonPO') || keys.includes('signIn') || keys.includes('invoice'))) {
    return "document-submission";
  }
  
  // Check for Amazon order fields
  if (keys.includes('Wishlist_Link') || keys.includes('wishlist_link')) {
    return "amazon-order";
  }
  
  // Check for target order fields (multiple item URLs)
  if (keys.includes('First_Item_Url') || 
      keys.includes('Second_Item_Url') ||
      keys.some(key => key.includes('Item_Url'))) {
    return "target-order";
  }
  
  // Final fallback
  console.log('⚠️ Could not detect form type from fields:', keys);
  return "target-order"; // Default fallback
}

// ========================================
// SHEET WRITING
// ========================================

/**
 * Write form data to appropriate sheet
 * @param {Object} formData - Form submission data
 * @returns {Object} Write result with success status and row number
 */
function writeToSheet(formData) {
  try {
    const formType = formData.form_type;
    const sheetName = SHEET_NAMES[formType];
    
    if (!sheetName) {
      throw new Error(`Unknown form type: ${formType}. Available: ${Object.keys(SHEET_NAMES).join(', ')}`);
    }
    
    const sheet = getCachedSheet(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    console.log(`📊 Writing to sheet: ${sheetName} (form type: ${formType})`);
    
    // Get existing headers
    const lastColumn = sheet.getLastColumn();
    const headers = lastColumn > 0 ? 
      sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : 
      [];
    
    console.log(`📊 Existing headers (${headers.length}):`, headers.join(', '));
    
    // Build new headers array
    const newHeaders = headers.length > 0 ? [...headers] : ['Timestamp'];
    let columnsAdded = 0;
    
    // Ensure Confirmation_Number column exists
    if (!newHeaders.includes('Confirmation_Number')) {
      newHeaders.push('Confirmation_Number');
      columnsAdded++;
      console.log('➕ Adding Confirmation_Number column');
    }
    
    // Ensure ID columns exist
    const idColumns = ['Historical_Submission_Id', 'Semester_Id', 'Submission_Id'];
    idColumns.forEach(col => {
      if (!newHeaders.includes(col)) {
        newHeaders.push(col);
        columnsAdded++;
        console.log(`➕ Adding ${col} column`);
      }
    });
    
    // Add file link columns for this form type
    const expectedFileColumns = getExpectedFileColumnsForFormType(formType);
    expectedFileColumns.forEach(columnName => {
      if (!newHeaders.includes(columnName) && formData[columnName]) {
        newHeaders.push(columnName);
        columnsAdded++;
        console.log(`➕ Adding file column: ${columnName}`);
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
        
        // Always include Event_Date columns
        if (key === 'Event_Date' || key === 'event_date' || key === 'Event Date') {
          newHeaders.push(key);
          columnsAdded++;
          console.log(`✅ Added Event Date column: ${key}`);
          return;
        }
        
        // Handle HTML date format
        if (isHTMLDateFormat(value)) {
          newHeaders.push(key);
          columnsAdded++;
          console.log(`✅ Added HTML date field: ${key}`);
          return;
        }
        
        // Standard field inclusion
        if (value !== null && 
            value !== undefined && 
            value !== "" && 
            typeof value !== 'object') {
          newHeaders.push(key);
          columnsAdded++;
          console.log(`➕ Adding field column: ${key}`);
        }
      }
    });
    
    // Update headers in sheet if new columns were added
    if (columnsAdded > 0 || headers.length === 0) {
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      console.log(`📊 Headers updated: ${newHeaders.length} total columns`);
    }
    
    // Build row data
    const rowData = newHeaders.map(header => {
      // Timestamp
      if (header === 'Timestamp') {
        const now = new Date();
        return Utilities.formatDate(
          now,
          'America/New_York',
          'MM/dd/yyyy hh:mm:ss a'
        );
      }
      
      // ID columns (will be filled by assignIDsToRow)
      if (header === 'Historical_Submission_Id' || 
          header === 'Semester_Id' || 
          header === 'Submission_Id') {
        return ''; // Leave blank for ID assignment
      }
      
      // Confirmation Number (will be filled after)
      if (header === 'Confirmation_Number') {
        return ''; // Leave blank for now
      }
      
      let value = formData[header];
      
      // Handle HTML date format
      if (isHTMLDateFormat(value)) {
        console.log(`📅 Processing HTML date: ${header} = "${value}"`);
        const formattedDate = formatHTMLDate(value);
        console.log(`📅 Formatted date: "${value}" → "${formattedDate}"`);
        return formattedDate;
      }
      
      // Standard value processing
      if (value === null || value === undefined) {
        value = "";
      } else if (typeof value === 'object' && !(value instanceof Date)) {
        value = "";
      } else if (typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
        value = String(value);
      }
      
      return value;
    });
    
    // Write the row
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    console.log(`✅ Data written to row ${nextRow} in sheet ${sheetName}`);
    
    return {
      success: true,
      row: nextRow,
      sheetName: sheetName,
      columnsAdded: columnsAdded
    };
    
  } catch (error) {
    console.error('❌ Sheet writing failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ========================================
// CONFIRMATION NUMBER
// ========================================

/**
 * Generate confirmation number (CD-MMDDYYYY-CLUBNAME)
 * @param {string} clubName - Organization name
 * @param {Date} submissionDate - Optional submission date
 * @returns {string} Confirmation number
 */
function generateConfirmationNumber(clubName, submissionDate) {
  const date = submissionDate || new Date();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  
  const sanitizedClubName = sanitizeFileName(clubName, 15).toUpperCase();
  
  return `CD-${month}${day}${year}-${sanitizedClubName}`;
}

/**
 * Add confirmation number to sheet row
 * @param {Sheet} sheet - Google Sheets sheet
 * @param {number} row - Row number
 * @param {string} confirmationNumber - Confirmation number to add
 */
function addConfirmationNumberToSheet(sheet, row, confirmationNumber) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const confirmationCol = findColumnIndex(headers, 'Confirmation_Number');
    
    if (confirmationCol !== -1) {
      sheet.getRange(row, confirmationCol + 1).setValue(confirmationNumber);
      console.log(`✅ Added confirmation number to row ${row}: ${confirmationNumber}`);
    }
  } catch (error) {
    console.error('❌ Failed to add confirmation number:', error);
  }
}
