/**
 * ============================================
 * FILE PROCESSING SYSTEM
 * Unified file upload and Drive management
 * ============================================
 */

// Configuration
const DRIVE_CONFIG = {
  FOLDER_ID: "1Zkk4aysbiblbpLHcfzjVC5jjDH-Q2Ct7"
};

let _cachedFolder = null;

// ========================================
// FOLDER ACCESS
// ========================================

/**
 * Get cached Drive folder
 * @returns {Folder} Google Drive folder
 */
function getCachedFolder() {
  if (!_cachedFolder) {
    console.log('📁 Caching Google Drive folder...');
    try {
      _cachedFolder = DriveApp.getFolderById(DRIVE_CONFIG.FOLDER_ID);
      console.log('✅ Drive folder cached');
    } catch (error) {
      console.error('❌ Drive folder failed:', error);
      throw new Error('Failed to access Drive folder: ' + error.toString());
    }
  }
  return _cachedFolder;
}

// ========================================
// FILE PROCESSING - WEB FORM SUBMISSIONS
// ========================================

/**
 * Process files from web form submission and add links to formData
 * @param {Object} formData - Form data with fileData property
 */
function processFilesAndAddLinksToFormData(formData) {
  try {
    const folder = getCachedFolder();
    
    // Check if it's multiple files (document submission) or single file (target/amazon form)
    const isMultipleFiles = typeof formData.fileData === 'object' && !formData.fileData.fileName;
    
    if (isMultipleFiles) {
      console.log('📄 Processing multiple files for document submission');
      processMultipleFilesAndAddLinks(formData, folder);
    } else {
      console.log('📄 Processing single file');
      processSingleFileAndAddLink(formData, folder);
    }
    
  } catch (error) {
    console.error('❌ File processing failed:', error);
    // Don't throw - allow form submission without files
  }
}

/**
 * Process multiple files (document submission)
 * @param {Object} formData - Form data
 * @param {Folder} folder - Drive folder
 */
function processMultipleFilesAndAddLinks(formData, folder) {
  // Extract naming components
  const orgName = sanitizeFileName(formData.Student_Organization || formData.Organization_Name || 'UnknownOrg');
  const firstName = sanitizeFileName(formData.First_Name || 'Unknown');
  const lastName = sanitizeFileName(formData.Last_Name || 'User');
  const dateString = getDateString();
  
  console.log(`📄 File naming: ${orgName}_${firstName}_${lastName}_[type]_${dateString}`);
  
  let filesProcessed = 0;
  
  // Process each file in fileData object
  Object.entries(formData.fileData).forEach(([inputName, fileObj]) => {
    try {
      console.log(`📄 Processing ${inputName}:`, fileObj.fileName);
      
      // Create standardized filename
      const fileExtension = getFileExtension(fileObj.fileName);
      const docType = getDocumentTypeLabel(inputName);
      const fileName = `${orgName}_${firstName}_${lastName}_${docType}_${dateString}.${fileExtension}`;
      
      // Create file blob from base64 data
      const fileBlob = Utilities.newBlob(
        Utilities.base64Decode(fileObj.data),
        fileObj.mimeType,
        fileName
      );
      
      // Upload to Drive
      const driveFile = folder.createFile(fileBlob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileUrl = driveFile.getUrl();
      
      // Add file link and timestamp to formData
      const columnName = getColumnNameForFileType(inputName);
      formData[columnName] = fileUrl;
      formData[`${columnName}_Timestamp`] = getFormattedTimestamp();
      
      console.log(`✅ ${inputName} → ${columnName} = ${fileUrl}`);
      filesProcessed++;
      
    } catch (error) {
      console.error(`❌ Failed to process ${inputName}:`, error);
    }
  });
  
  console.log(`📁 Document files processed: ${filesProcessed}`);
}

/**
 * Process single file (target/amazon form)
 * @param {Object} formData - Form data
 * @param {Folder} folder - Drive folder
 */
function processSingleFileAndAddLink(formData, folder) {
  try {
    const fileObj = formData.fileData;
    console.log(`📄 Processing single file:`, fileObj.fileName);
    
    // Create standardized filename
    const orgName = sanitizeFileName(formData.Student_Organization || 'UnknownOrg');
    const dateString = getDateString();
    const fileExtension = getFileExtension(fileObj.fileName);
    const fileName = `${orgName}_Upload_${dateString}.${fileExtension}`;
    
    // Create file blob from base64 data
    const fileBlob = Utilities.newBlob(
      Utilities.base64Decode(fileObj.data),
      fileObj.mimeType,
      fileName
    );
    
    // Upload to Drive
    const driveFile = folder.createFile(fileBlob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = driveFile.getUrl();
    
    // Add file link and timestamp to formData
    formData['FileLink'] = fileUrl;
    formData['FileLink_Timestamp'] = getFormattedTimestamp();
    formData['FileName'] = fileName;
    
    console.log(`✅ Single file → FileLink = ${fileUrl}`);
    
  } catch (error) {
    console.error('❌ Single file processing failed:', error);
  }
}

// ========================================
// FILE PROCESSING - BACKEND UI UPLOADS
// ========================================

/**
 * Upload file manually from backend Order Management UI
 * @param {Object} fileData - File data with base64 encoding
 * @param {string} orderId - Order ID
 * @param {string} platform - Platform (amazon/target/document)
 * @param {string} fileType - File type identifier
 * @returns {Object} Upload result
 */
function uploadManualFileToOrder(fileData, orderId, platform, fileType) {
  try {
    console.log(`📤 Manual file upload for ${platform} order ${orderId}`);
    
    // Get the Drive folder
    const folder = getCachedFolder();
    
    // Create filename
    const timestamp = new Date();
    const dateString = Utilities.formatDate(timestamp, 'America/New_York', 'MM-dd-yyyy_HHmmss');
    const fileExtension = getFileExtension(fileData.fileName);
    const fileName = `${platform}_${orderId}_${fileType}_${dateString}.${fileExtension}`;
    
    // Create file blob
    const fileBlob = Utilities.newBlob(
      Utilities.base64Decode(fileData.data),
      fileData.mimeType,
      fileName
    );
    
    // Upload to Drive
    const driveFile = folder.createFile(fileBlob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = driveFile.getUrl();
    
    // Determine sheet name
    const sheetMap = {
      'amazon': 'AmazonOrders',
      'target': 'TargetOrders',
      'document': 'DocumentSubmissions'
    };
    const sheetName = sheetMap[platform];
    
    // Update sheet with file link and timestamp
    const updates = {};
    updates[`${fileType}_FileLink`] = fileUrl;
    updates[`${fileType}_Upload_Timestamp`] = getFormattedTimestamp(timestamp);
    updates[`${fileType}_FileName`] = fileName;
    
    // Find and update the order row
    updateOrderFileLinks(orderId, sheetName, updates);
    
    console.log(`✅ Manual file uploaded: ${fileName}`);
    
    return {
      success: true,
      fileUrl: fileUrl,
      fileName: fileName,
      timestamp: updates[`${fileType}_Upload_Timestamp`]
    };
    
  } catch (error) {
    console.error('❌ Manual file upload failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Update order with file links
 * @param {string} orderId - Order ID
 * @param {string} sheetName - Sheet name
 * @param {Object} updates - Updates to apply
 */
function updateOrderFileLinks(orderId, sheetName, updates) {
  try {
    const sheet = getCachedSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find row by Historical_Submission_Id or Submission_Id
    const historicalIdCol = findColumnIndex(headers, 'Historical_Submission_Id');
    const submissionIdCol = findColumnIndex(headers, 'Submission_Id');
    
    let targetRow = -1;
    
    // Search for matching row
    for (let i = 1; i < data.length; i++) {
      const historicalId = historicalIdCol !== -1 ? data[i][historicalIdCol] : null;
      const submissionId = submissionIdCol !== -1 ? data[i][submissionIdCol] : null;
      
      if (historicalId == orderId || submissionId == orderId) {
        targetRow = i + 1; // Convert to 1-based
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`Order ${orderId} not found in ${sheetName}`);
    }
    
    // Apply updates
    Object.keys(updates).forEach(key => {
      const columnIndex = findColumnIndex(headers, key);
      if (columnIndex !== -1) {
        sheet.getRange(targetRow, columnIndex + 1).setValue(updates[key]);
        console.log(`✅ Updated ${key} in row ${targetRow}`);
      } else {
        // Add new column if it doesn't exist
        const newColIndex = headers.length + 1;
        sheet.getRange(1, newColIndex).setValue(key);
        sheet.getRange(targetRow, newColIndex).setValue(updates[key]);
        console.log(`➕ Added new column ${key} at position ${newColIndex}`);
      }
    });
    
    console.log(`✅ File links updated for order ${orderId}`);
    
  } catch (error) {
    console.error('❌ Error updating file links:', error);
    throw error;
  }
}

// ========================================
// HELPER FUNCTIONS
// ========================================

/**
 * Get column name for file type
 * @param {string} inputName - Input field name from form
 * @returns {string} Column name for sheet
 */
function getColumnNameForFileType(inputName) {
  const columnMap = {
    'nonPO': 'nonPO_FileLink',
    'signIn': 'signIn_FileLink',
    'invoice': 'invoice_FileLink',
    'eventFlyer': 'eventFlyer_FileLink',
    'officialReceipt': 'officialReceipt_FileLink',
    'theFile': 'FileLink'
  };
  
  return columnMap[inputName] || `${inputName}_FileLink`;
}

/**
 * Get expected file columns for form type
 * @param {string} formType - Form type identifier
 * @returns {Array} Array of expected column names
 */
function getExpectedFileColumnsForFormType(formType) {
  const columnsByFormType = {
    'document-submission': [
      'nonPO_FileLink',
      'nonPO_FileLink_Timestamp',
      'signIn_FileLink',
      'signIn_FileLink_Timestamp',
      'invoice_FileLink',
      'invoice_FileLink_Timestamp',
      'eventFlyer_FileLink',
      'eventFlyer_FileLink_Timestamp',
      'officialReceipt_FileLink',
      'officialReceipt_FileLink_Timestamp'
    ],
    'target-order': ['FileLink', 'FileLink_Timestamp', 'FileName'],
    'amazon-order': [], // Amazon doesn't support file uploads
    'campus-club-space-feedback': [] // Feedback doesn't support file uploads
  };
  
  return columnsByFormType[formType] || [];
}

/**
 * Get document type label
 * @param {string} inputName - Input field name
 * @returns {string} Human-readable document type
 */
function getDocumentTypeLabel(inputName) {
  const typeMap = {
    'nonPO': 'NonTaxPO',
    'signIn': 'SignInSheet',
    'invoice': 'Invoice',
    'eventFlyer': 'EventFlyer',
    'officialReceipt': 'OfficialReceipt',
    'theFile': 'Upload'
  };
  
  return typeMap[inputName] || 'Document';
}

/**
 * Validate file before upload
 * @param {Object} fileData - File data object
 * @returns {Object} Validation result
 */
function validateFile(fileData) {
  const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
  const ALLOWED_TYPES = [
    'application/pdf',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'image/jpeg',
    'image/png',
    'image/gif'
  ];
  
  // Check if file data exists
  if (!fileData || !fileData.data) {
    return {
      valid: false,
      error: 'No file data provided'
    };
  }
  
  // Check MIME type
  if (!ALLOWED_TYPES.includes(fileData.mimeType)) {
    return {
      valid: false,
      error: `File type ${fileData.mimeType} not allowed`
    };
  }
  
  // Estimate file size from base64 data
  const estimatedSize = (fileData.data.length * 3) / 4;
  if (estimatedSize > MAX_FILE_SIZE) {
    return {
      valid: false,
      error: `File size exceeds 10MB limit`
    };
  }
  
  return {
    valid: true
  };
}
