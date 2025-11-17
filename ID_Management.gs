/**
 * ============================================
 * ID MANAGEMENT SYSTEM
 * Consolidated ID generation and assignment
 * ============================================
 */

// ========================================
// ID GENERATION
// ========================================

/**
 * Get next Historical Submission ID for a sheet
 * @param {Sheet} sheet - Google Sheets sheet object
 * @param {number} columnIndex - Column index for Historical_Submission_Id
 * @returns {number} Next ID to use
 */
function getNextHistoricalId(sheet, columnIndex) {
  try {
    const sheetName = sheet.getName();
    console.log(`🔍 Getting next Historical ID for ${sheetName}`);
    
    const lastRow = sheet.getLastRow();
    let maxHistoricalId = 0;
    
    // Scan existing IDs in column
    if (lastRow > 1 && columnIndex !== -1) {
      const historicalIds = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues();
      maxHistoricalId = getMaxIdFromColumn(historicalIds);
      console.log(`📊 ${sheetName} - Current max Historical ID: ${maxHistoricalId}`);
    }
    
    const nextId = maxHistoricalId + 1;
    console.log(`🎯 ${sheetName} - Next Historical ID: ${nextId}`);
    
    // Update sheet-specific counter
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`${sheetName}_historical_counter`, maxHistoricalId.toString());
    
    return nextId;
    
  } catch (error) {
    console.error('❌ Error getting next Historical ID:', error);
    // Fallback to stored counter + 1
    return getNextSheetSpecificId(sheet.getName(), 'historical');
  }
}

/**
 * Get next Submission ID for a sheet
 * @param {Sheet} sheet - Google Sheets sheet object
 * @param {number} columnIndex - Column index for Submission_Id
 * @returns {number} Next ID to use
 */
function getNextSubmissionId(sheet, columnIndex) {
  try {
    const sheetName = sheet.getName();
    console.log(`🔍 Getting next Submission ID for ${sheetName}`);
    
    const lastRow = sheet.getLastRow();
    let maxSubmissionId = 0;
    
    // Scan existing IDs in column
    if (lastRow > 1 && columnIndex !== -1) {
      const submissionIds = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues();
      maxSubmissionId = getMaxIdFromColumn(submissionIds);
      console.log(`📊 ${sheetName} - Current max Submission ID: ${maxSubmissionId}`);
    }
    
    const nextId = maxSubmissionId + 1;
    console.log(`🎯 ${sheetName} - Next Submission ID: ${nextId}`);
    
    // Update sheet-specific counter
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`${sheetName}_submission_counter`, maxSubmissionId.toString());
    
    return nextId;
    
  } catch (error) {
    console.error('❌ Error getting next Submission ID:', error);
    // Fallback to stored counter + 1
    return getNextSheetSpecificId(sheet.getName(), 'submission');
  }
}

/**
 * Fallback: Get next sheet-specific ID from stored counter
 * @param {string} sheetName - Sheet name
 * @param {string} idType - 'historical' or 'submission'
 * @returns {number} Next ID
 */
function getNextSheetSpecificId(sheetName, idType) {
  const props = PropertiesService.getScriptProperties();
  const counterKey = `${sheetName}_${idType}_counter`;
  const current = parseInt(props.getProperty(counterKey) || '0');
  const next = current + 1;
  props.setProperty(counterKey, next.toString());
  console.log(`🔄 ${sheetName} - Fallback ${idType} ID: ${next}`);
  return next;
}

// ========================================
// ID ASSIGNMENT
// ========================================

/**
 * Assign all IDs to a specific row immediately after writing
 * @param {Sheet} sheet - Google Sheets sheet object
 * @param {number} targetRow - Row number to assign IDs to
 * @returns {Object} Success status and assigned IDs
 */
function assignIDsToRow(sheet, targetRow) {
  try {
    console.log(`🎯 Assigning IDs to row ${targetRow} in ${sheet.getName()}`);
    
    // Get headers
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find ID columns
    const historicalIdCol = findColumnIndex(headers, 'Historical_Submission_Id');
    const semesterIdCol = findColumnIndex(headers, 'Semester_Id');
    const submissionIdCol = findColumnIndex(headers, 'Submission_Id');
    
    console.log('Column positions:', {
      Historical: historicalIdCol,
      Semester: semesterIdCol,
      Submission: submissionIdCol
    });
    
    // Generate IDs
    const assignedIds = {};
    
    // Historical ID
    if (historicalIdCol !== -1) {
      const currentValue = sheet.getRange(targetRow, historicalIdCol + 1).getValue();
      if (!currentValue || currentValue === '') {
        const nextId = getNextHistoricalId(sheet, historicalIdCol);
        sheet.getRange(targetRow, historicalIdCol + 1).setValue(nextId);
        assignedIds.Historical_Submission_Id = nextId;
        console.log(`✅ Set Historical_Submission_Id: ${nextId}`);
      } else {
        assignedIds.Historical_Submission_Id = currentValue;
        console.log(`ℹ️ Historical_Submission_Id already set: ${currentValue}`);
      }
    }
    
    // Semester ID
    if (semesterIdCol !== -1) {
      const currentValue = sheet.getRange(targetRow, semesterIdCol + 1).getValue();
      if (!currentValue || currentValue === '') {
        const semester = getSemester();
        sheet.getRange(targetRow, semesterIdCol + 1).setValue(semester);
        assignedIds.Semester_Id = semester;
        console.log(`✅ Set Semester_Id: ${semester}`);
      } else {
        assignedIds.Semester_Id = currentValue;
        console.log(`ℹ️ Semester_Id already set: ${currentValue}`);
      }
    }
    
    // Submission ID
    if (submissionIdCol !== -1) {
      const currentValue = sheet.getRange(targetRow, submissionIdCol + 1).getValue();
      if (!currentValue || currentValue === '') {
        const nextId = getNextSubmissionId(sheet, submissionIdCol);
        sheet.getRange(targetRow, submissionIdCol + 1).setValue(nextId);
        assignedIds.Submission_Id = nextId;
        console.log(`✅ Set Submission_Id: ${nextId}`);
      } else {
        assignedIds.Submission_Id = currentValue;
        console.log(`ℹ️ Submission_Id already set: ${currentValue}`);
      }
    }
    
    console.log(`🎉 IDs assigned successfully to row ${targetRow}`);
    
    return {
      success: true,
      row: targetRow,
      assignedIds: assignedIds
    };
    
  } catch (error) {
    console.error('❌ Error assigning IDs:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ========================================
// INITIALIZATION
// ========================================

/**
 * Initialize counters by scanning existing data in all sheets
 */
function initializeCountersFromExistingData() {
  try {
    console.log('🔍 Scanning existing data to initialize sheet-specific counters...');
    
    const spreadsheet = getCachedSpreadsheet();
    const targetSheets = ['TargetOrders', 'AmazonOrders', 'DocumentSubmissions', 'CampusClubSpaceFeedback'];
    
    const props = PropertiesService.getScriptProperties();
    
    targetSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        console.log(`⚠️ Sheet ${sheetName} not found, skipping...`);
        return;
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        console.log(`📄 Sheet ${sheetName} is empty, setting counters to 0...`);
        props.setProperty(`${sheetName}_historical_counter`, '0');
        props.setProperty(`${sheetName}_submission_counter`, '0');
        return;
      }
      
      console.log(`📊 Scanning ${sheetName} with ${lastRow} rows...`);
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const historicalCol = findColumnIndex(headers, 'Historical_Submission_Id');
      const submissionCol = findColumnIndex(headers, 'Submission_Id');
      
      let maxHistoricalId = 0;
      let maxSubmissionId = 0;
      
      if (historicalCol !== -1) {
        const historicalIds = sheet.getRange(2, historicalCol + 1, lastRow - 1, 1).getValues();
        maxHistoricalId = getMaxIdFromColumn(historicalIds);
        console.log(`📈 ${sheetName} - Max Historical ID: ${maxHistoricalId}`);
      }
      
      if (submissionCol !== -1) {
        const submissionIds = sheet.getRange(2, submissionCol + 1, lastRow - 1, 1).getValues();
        maxSubmissionId = getMaxIdFromColumn(submissionIds);
        console.log(`📈 ${sheetName} - Max Submission ID: ${maxSubmissionId}`);
      }
      
      // Store sheet-specific counters
      props.setProperty(`${sheetName}_historical_counter`, maxHistoricalId.toString());
      props.setProperty(`${sheetName}_submission_counter`, maxSubmissionId.toString());
      
      console.log(`🎯 ${sheetName} - Initialized counters - Historical: ${maxHistoricalId}, Submission: ${maxSubmissionId}`);
    });
    
    console.log('✅ Counter initialization complete');
    
    return {
      success: true,
      message: 'Counters initialized successfully'
    };
    
  } catch (error) {
    console.error('❌ Error initializing counters:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ========================================
// BACKFILL MISSING IDs
// ========================================

/**
 * Fill missing IDs in existing rows retroactively
 * @param {string} sheetName - Optional: specific sheet to backfill
 * @returns {Object} Results of backfill operation
 */
function backfillMissingIds(sheetName = null) {
  try {
    console.log('🔄 Starting retroactive ID assignment...');
    
    const spreadsheet = getCachedSpreadsheet();
    const targetSheets = sheetName ? 
      [sheetName] : 
      ['TargetOrders', 'AmazonOrders', 'DocumentSubmissions', 'CampusClubSpaceFeedback'];
    
    let totalUpdates = 0;
    const results = {};
    
    targetSheets.forEach(name => {
      console.log(`\n📊 Processing sheet: ${name}`);
      
      const sheet = spreadsheet.getSheetByName(name);
      if (!sheet) {
        console.log(`❌ Sheet ${name} not found, skipping...`);
        results[name] = { success: false, message: 'Sheet not found' };
        return;
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        console.log(`📄 Sheet ${name} is empty, skipping...`);
        results[name] = { success: true, updates: 0, message: 'No data to process' };
        return;
      }
      
      // Get headers and find column positions
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const historicalCol = findColumnIndex(headers, 'Historical_Submission_Id');
      const semesterCol = findColumnIndex(headers, 'Semester_Id');
      const submissionCol = findColumnIndex(headers, 'Submission_Id');
      
      console.log(`Column positions - Historical: ${historicalCol}, Semester: ${semesterCol}, Submission: ${submissionCol}`);
      
      let sheetUpdates = 0;
      
      // Process each row
      for (let row = 2; row <= lastRow; row++) {
        let rowUpdated = false;
        
        // Check and fill Historical_Submission_Id
        if (historicalCol !== -1) {
          const currentValue = sheet.getRange(row, historicalCol + 1).getValue();
          if (!currentValue || currentValue === '') {
            const nextHistoricalId = getNextHistoricalId(sheet, historicalCol);
            sheet.getRange(row, historicalCol + 1).setValue(nextHistoricalId);
            console.log(`✅ Row ${row}: Set Historical_Submission_Id to ${nextHistoricalId}`);
            rowUpdated = true;
          }
        }
        
        // Check and fill Semester_Id
        if (semesterCol !== -1) {
          const currentValue = sheet.getRange(row, semesterCol + 1).getValue();
          if (!currentValue || currentValue === '') {
            const currentSemester = getSemester();
            sheet.getRange(row, semesterCol + 1).setValue(currentSemester);
            console.log(`✅ Row ${row}: Set Semester_Id to ${currentSemester}`);
            rowUpdated = true;
          }
        }
        
        // Check and fill Submission_Id
        if (submissionCol !== -1) {
          const currentValue = sheet.getRange(row, submissionCol + 1).getValue();
          if (!currentValue || currentValue === '') {
            const nextSubmissionId = getNextSubmissionId(sheet, submissionCol);
            sheet.getRange(row, submissionCol + 1).setValue(nextSubmissionId);
            console.log(`✅ Row ${row}: Set Submission_Id to ${nextSubmissionId}`);
            rowUpdated = true;
          }
        }
        
        if (rowUpdated) {
          sheetUpdates++;
        }
        
        // Small delay to avoid hitting rate limits
        if (row % 10 === 0) {
          Utilities.sleep(100);
        }
      }
      
      results[name] = {
        success: true,
        updates: sheetUpdates,
        totalRows: lastRow - 1,
        message: `Updated ${sheetUpdates} rows`
      };
      
      totalUpdates += sheetUpdates;
      console.log(`📋 ${name} completed: ${sheetUpdates} updates`);
    });
    
    console.log(`\n🎉 Retroactive ID assignment completed!`);
    console.log(`📊 Total rows updated: ${totalUpdates}`);
    
    return {
      success: true,
      totalUpdates: totalUpdates,
      sheetResults: results,
      message: `Successfully updated ${totalUpdates} rows with missing IDs`
    };
    
  } catch (error) {
    console.error('❌ Error in retroactive ID assignment:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ========================================
// SETUP FUNCTION
// ========================================

/**
 * Initial setup - run once to configure the system
 */
function setupIDManagement() {
  try {
    console.log('🔧 Setting up ID Management system...');
    
    // Initialize spreadsheet cache
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    PropertiesService.getScriptProperties().setProperty('key', spreadsheet.getId());
    
    console.log('✅ Spreadsheet key stored');
    
    // Initialize counters from existing data
    const initResult = initializeCountersFromExistingData();
    
    if (initResult.success) {
      console.log('✅ ID Management setup complete');
      return {
        success: true,
        message: 'ID Management system configured successfully'
      };
    } else {
      throw new Error(initResult.error);
    }
    
  } catch (error) {
    console.error('❌ Setup failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}
