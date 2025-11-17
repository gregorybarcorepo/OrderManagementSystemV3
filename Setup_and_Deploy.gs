/**
 * ============================================
 * SETUP AND DEPLOYMENT
 * Initial setup and deployment functions
 * ============================================
 */

/**
 * MASTER SETUP FUNCTION - Run this FIRST after creating new Google Sheets
 * This sets up everything needed for the system to work
 */
function masterSetup() {
  console.log('🚀 Starting Master Setup...');
  console.log('═══════════════════════════════════════════════════════');
  
  const results = {
    steps: [],
    success: true,
    errors: []
  };
  
  try {
    // Step 1: Setup spreadsheet connection
    console.log('\n📊 Step 1: Setting up spreadsheet connection...');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    PropertiesService.getScriptProperties().setProperty('key', spreadsheetId);
    results.steps.push({
      step: 'Spreadsheet Connection',
      status: '✅ Success',
      details: `Connected to: ${spreadsheet.getName()}`
    });
    console.log('✅ Spreadsheet connection established');
    
    // Step 2: Create required sheets
    console.log('\n📋 Step 2: Creating required sheets...');
    const requiredSheets = ['TargetOrders', 'AmazonOrders', 'DocumentSubmissions', 'CampusClubSpaceFeedback', 'Names'];
    const createdSheets = [];
    
    requiredSheets.forEach(sheetName => {
      let sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        createdSheets.push(sheetName);
        console.log(`  ➕ Created sheet: ${sheetName}`);
      } else {
        console.log(`  ✓ Sheet exists: ${sheetName}`);
      }
    });
    
    results.steps.push({
      step: 'Sheet Creation',
      status: '✅ Success',
      details: `Created ${createdSheets.length} new sheets: ${createdSheets.join(', ') || 'None (all existed)'}`
    });
    
    // Step 3: Initialize sheet headers
    console.log('\n📝 Step 3: Initializing sheet headers...');
    initializeSheetHeaders();
    results.steps.push({
      step: 'Sheet Headers',
      status: '✅ Success',
      details: 'Basic headers initialized for all sheets'
    });
    
    // Step 4: Initialize ID management system
    console.log('\n🎯 Step 4: Initializing ID management system...');
    const idSetup = setupIDManagement();
    if (idSetup.success) {
      results.steps.push({
        step: 'ID Management',
        status: '✅ Success',
        details: 'ID counters initialized'
      });
    } else {
      throw new Error('ID Management setup failed: ' + idSetup.error);
    }
    
    // Step 5: Verify Drive folder access
    console.log('\n📁 Step 5: Verifying Drive folder access...');
    try {
      const folder = DriveApp.getFolderById(SYSTEM_CONFIG.DRIVE_FOLDER_ID);
      results.steps.push({
        step: 'Drive Folder',
        status: '✅ Success',
        details: `Connected to folder: ${folder.getName()}`
      });
      console.log(`✅ Drive folder accessible: ${folder.getName()}`);
    } catch (error) {
      results.steps.push({
        step: 'Drive Folder',
        status: '⚠️ Warning',
        details: 'Could not access Drive folder. Update DRIVE_FOLDER_ID in Configuration.gs'
      });
      console.log('⚠️ Drive folder not accessible. You need to update DRIVE_FOLDER_ID in Configuration.gs');
    }
    
    // Step 6: Deploy as web app instructions
    console.log('\n🌐 Step 6: Web App Deployment Status...');
    results.steps.push({
      step: 'Web App Deployment',
      status: 'ℹ️ Manual Step Required',
      details: 'After setup completes, deploy as web app (see instructions below)'
    });
    
    console.log('\n═══════════════════════════════════════════════════════');
    console.log('🎉 MASTER SETUP COMPLETED SUCCESSFULLY!');
    console.log('═══════════════════════════════════════════════════════\n');
    
    // Print summary
    console.log('📊 SETUP SUMMARY:');
    results.steps.forEach((step, index) => {
      console.log(`${index + 1}. ${step.step}: ${step.status}`);
      console.log(`   ${step.details}`);
    });
    
    console.log('\n📝 NEXT STEPS:');
    console.log('1. Deploy this script as a web app:');
    console.log('   - Click "Deploy" > "New deployment"');
    console.log('   - Choose "Web app" type');
    console.log('   - Execute as: "Me"');
    console.log('   - Who has access: "Anyone"');
    console.log('   - Click "Deploy"');
    console.log('   - Copy the web app URL');
    console.log('\n2. Update your HTML forms with the new web app URL');
    console.log('\n3. Test a form submission');
    console.log('\n4. Run testFullSystem() to verify everything works');
    
    return {
      success: true,
      message: 'Setup completed successfully',
      results: results
    };
    
  } catch (error) {
    console.error('\n❌ SETUP FAILED:', error);
    results.success = false;
    results.errors.push(error.toString());
    
    return {
      success: false,
      message: 'Setup failed: ' + error.toString(),
      results: results
    };
  }
}

/**
 * Initialize basic headers for all sheets
 */
function initializeSheetHeaders() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // TargetOrders headers
  const targetSheet = spreadsheet.getSheetByName('TargetOrders');
  if (targetSheet.getLastRow() === 0) {
    const targetHeaders = [
      'Timestamp', 'Historical_Submission_Id', 'Semester_Id', 'Submission_Id', 
      'Confirmation_Number', 'Submission_Email_Address', 'Student_Organization',
      'Event_Name', 'Event_Date', 'Pickup_Person_Name', 'Pickup_Person_Email',
      'Pickup_Person_Phone', 'First_Item_Url', 'First_Item_Quantity',
      'Cart_Total', 'Order_Status', 'Processed_By', 'Comments'
    ];
    targetSheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    targetSheet.setFrozenRows(1);
    console.log('  ✓ TargetOrders headers initialized');
  }
  
  // AmazonOrders headers
  const amazonSheet = spreadsheet.getSheetByName('AmazonOrders');
  if (amazonSheet.getLastRow() === 0) {
    const amazonHeaders = [
      'Timestamp', 'Historical_Submission_Id', 'Semester_Id', 'Submission_Id',
      'Confirmation_Number', 'Submission_Email_Address', 'Student_Organization',
      'Event_Name', 'Event_Date', 'Pickup_Person_Name', 'Pickup_Person_Email',
      'Pickup_Person_Phone', 'Wishlist_Link', 'Associated Order Numbers',
      'Total_1', 'Total_2', 'Total_Order', 'Order_Status', 'Processed_By', 'Comments'
    ];
    amazonSheet.getRange(1, 1, 1, amazonHeaders.length).setValues([amazonHeaders]);
    amazonSheet.setFrozenRows(1);
    console.log('  ✓ AmazonOrders headers initialized');
  }
  
  // DocumentSubmissions headers
  const docSheet = spreadsheet.getSheetByName('DocumentSubmissions');
  if (docSheet.getLastRow() === 0) {
    const docHeaders = [
      'Timestamp', 'Historical_Submission_Id', 'Semester_Id', 'Submission_Id',
      'Confirmation_Number', 'Email_Address', 'First_Name', 'Last_Name',
      'Student_Organization', 'Event_Name', 'Event_Date',
      'nonPO_FileLink', 'signIn_FileLink', 'invoice_FileLink',
      'eventFlyer_FileLink', 'officialReceipt_FileLink',
      'Processed_By', 'Comments'
    ];
    docSheet.getRange(1, 1, 1, docHeaders.length).setValues([docHeaders]);
    docSheet.setFrozenRows(1);
    console.log('  ✓ DocumentSubmissions headers initialized');
  }
  
  // CampusClubSpaceFeedback headers
  const feedbackSheet = spreadsheet.getSheetByName('CampusClubSpaceFeedback');
  if (feedbackSheet.getLastRow() === 0) {
    const feedbackHeaders = [
      'Timestamp', 'Historical_Submission_Id', 'Semester_Id', 'Submission_Id',
      'Confirmation_Number', 'Organization_Name', 'Feedback_Type',
      'Space_Rating', 'Overall_Satisfaction', 'Comments'
    ];
    feedbackSheet.getRange(1, 1, 1, feedbackHeaders.length).setValues([feedbackHeaders]);
    feedbackSheet.setFrozenRows(1);
    console.log('  ✓ CampusClubSpaceFeedback headers initialized');
  }
  
  // Names sheet
  const namesSheet = spreadsheet.getSheetByName('Names');
  if (namesSheet.getLastRow() === 0) {
    const namesHeaders = ['Club Names', 'Employee Names'];
    namesSheet.getRange(1, 1, 1, namesHeaders.length).setValues([namesHeaders]);
    namesSheet.setFrozenRows(1);
    
    // Add some sample data
    const sampleData = [
      ['Sample Club 1', 'John Doe'],
      ['Sample Club 2', 'Jane Smith'],
      ['Sample Club 3', 'Mike Johnson']
    ];
    namesSheet.getRange(2, 1, sampleData.length, 2).setValues(sampleData);
    console.log('  ✓ Names sheet initialized with sample data');
  }
}

/**
 * Test the complete system end-to-end
 */
function testFullSystem() {
  console.log('🧪 Starting Full System Test...');
  console.log('═══════════════════════════════════════════════════════');
  
  const testResults = {
    tests: [],
    passed: 0,
    failed: 0
  };
  
  // Test 1: Spreadsheet Connection
  console.log('\n📊 Test 1: Spreadsheet Connection');
  try {
    const ss = getCachedSpreadsheet();
    console.log('✅ PASS: Connected to spreadsheet:', ss.getName());
    testResults.tests.push({ name: 'Spreadsheet Connection', status: 'PASS' });
    testResults.passed++;
  } catch (error) {
    console.log('❌ FAIL:', error.toString());
    testResults.tests.push({ name: 'Spreadsheet Connection', status: 'FAIL', error: error.toString() });
    testResults.failed++;
  }
  
  // Test 2: Sheet Access
  console.log('\n📋 Test 2: Sheet Access');
  try {
    const requiredSheets = ['TargetOrders', 'AmazonOrders', 'DocumentSubmissions', 'Names'];
    requiredSheets.forEach(sheetName => {
      const sheet = getCachedSheet(sheetName);
      if (!sheet) throw new Error(`Sheet ${sheetName} not found`);
    });
    console.log('✅ PASS: All required sheets accessible');
    testResults.tests.push({ name: 'Sheet Access', status: 'PASS' });
    testResults.passed++;
  } catch (error) {
    console.log('❌ FAIL:', error.toString());
    testResults.tests.push({ name: 'Sheet Access', status: 'FAIL', error: error.toString() });
    testResults.failed++;
  }
  
  // Test 3: ID Generation
  console.log('\n🎯 Test 3: ID Generation');
  try {
    const sheet = getCachedSheet('TargetOrders');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const historicalCol = findColumnIndex(headers, 'Historical_Submission_Id');
    const nextId = getNextHistoricalId(sheet, historicalCol);
    
    if (typeof nextId === 'number' && nextId > 0) {
      console.log('✅ PASS: Generated ID:', nextId);
      testResults.tests.push({ name: 'ID Generation', status: 'PASS' });
      testResults.passed++;
    } else {
      throw new Error('Invalid ID generated: ' + nextId);
    }
  } catch (error) {
    console.log('❌ FAIL:', error.toString());
    testResults.tests.push({ name: 'ID Generation', status: 'FAIL', error: error.toString() });
    testResults.failed++;
  }
  
  // Test 4: Names Data Retrieval
  console.log('\n👥 Test 4: Names Data Retrieval');
  try {
    const namesData = getNamesData();
    if (namesData.success && Array.isArray(namesData.clubs) && Array.isArray(namesData.employees)) {
      console.log('✅ PASS: Retrieved', namesData.clubs.length, 'clubs and', namesData.employees.length, 'employees');
      testResults.tests.push({ name: 'Names Data Retrieval', status: 'PASS' });
      testResults.passed++;
    } else {
      throw new Error('Invalid names data structure');
    }
  } catch (error) {
    console.log('❌ FAIL:', error.toString());
    testResults.tests.push({ name: 'Names Data Retrieval', status: 'FAIL', error: error.toString() });
    testResults.failed++;
  }
  
  // Test 5: Email Template Generation
  console.log('\n📧 Test 5: Email Template Generation');
  try {
    const testFormData = {
      Student_Organization: 'Test Club',
      Event_Name: 'Test Event',
      Event_Date: '12/31/2025'
    };
    const emailContent = generateEmailContent(testFormData, 'amazon-order', 'CD-12312025-TESTCLUB', getFormattedTimestamp());
    
    if (emailContent.subject && emailContent.htmlBody && emailContent.htmlBody.includes('CD-12312025-TESTCLUB')) {
      console.log('✅ PASS: Email template generated successfully');
      testResults.tests.push({ name: 'Email Template Generation', status: 'PASS' });
      testResults.passed++;
    } else {
      throw new Error('Invalid email template');
    }
  } catch (error) {
    console.log('❌ FAIL:', error.toString());
    testResults.tests.push({ name: 'Email Template Generation', status: 'FAIL', error: error.toString() });
    testResults.failed++;
  }
  
  // Test 6: Utility Functions
  console.log('\n🔧 Test 6: Utility Functions');
  try {
    const semester = getSemester();
    const dateStr = getDateString();
    const sanitized = sanitizeFileName('Test Club Name!@#');
    
    if (semester && dateStr && sanitized === 'TestClubName') {
      console.log('✅ PASS: Utility functions working');
      console.log('  Semester:', semester);
      console.log('  Date String:', dateStr);
      console.log('  Sanitized:', sanitized);
      testResults.tests.push({ name: 'Utility Functions', status: 'PASS' });
      testResults.passed++;
    } else {
      throw new Error('Utility function returned unexpected values');
    }
  } catch (error) {
    console.log('❌ FAIL:', error.toString());
    testResults.tests.push({ name: 'Utility Functions', status: 'FAIL', error: error.toString() });
    testResults.failed++;
  }
  
  // Print summary
  console.log('\n═══════════════════════════════════════════════════════');
  console.log('🧪 TEST SUMMARY');
  console.log('═══════════════════════════════════════════════════════');
  console.log(`✅ Passed: ${testResults.passed}`);
  console.log(`❌ Failed: ${testResults.failed}`);
  console.log(`📊 Total: ${testResults.tests.length}`);
  console.log(`📈 Success Rate: ${((testResults.passed / testResults.tests.length) * 100).toFixed(1)}%`);
  
  console.log('\n📋 Detailed Results:');
  testResults.tests.forEach((test, index) => {
    const icon = test.status === 'PASS' ? '✅' : '❌';
    console.log(`${index + 1}. ${icon} ${test.name}: ${test.status}`);
    if (test.error) {
      console.log(`   Error: ${test.error}`);
    }
  });
  
  if (testResults.failed === 0) {
    console.log('\n🎉 ALL TESTS PASSED! System is ready to use.');
  } else {
    console.log('\n⚠️ Some tests failed. Please review and fix issues before using the system.');
  }
  
  return testResults;
}

/**
 * Create a test submission (simulates form submission)
 */
function createTestSubmission() {
  console.log('🧪 Creating test submission...');
  
  const testFormData = {
    form_type: 'target-order',
    Submission_Email_Address: 'test@brooklyn.cuny.edu',
    Student_Organization: 'Test Club for System',
    Event_Name: 'Test Event',
    Event_Date: '2025-12-31',
    Pickup_Person_Name: 'Test Person',
    Pickup_Person_Email: 'pickup@brooklyn.cuny.edu',
    Pickup_Person_Phone: '555-1234',
    First_Item_Url: 'https://www.target.com/test-item',
    First_Item_Quantity: '2',
    Cart_Total: '25.99',
    Order_Status: 'New Order', // Updated
    'Form Submitter Notes': 'This is a test submission created by createTestSubmission()'
  };
  
  try {
    // Write to sheet
    const writeResult = writeToSheet(testFormData);
    
    if (!writeResult.success) {
      throw new Error('Failed to write test submission: ' + writeResult.error);
    }
    
    console.log('✅ Test data written to row', writeResult.row);
    
    // Assign IDs
    const sheet = getCachedSheet(writeResult.sheetName);
    const idResult = assignIDsToRow(sheet, writeResult.row);
    
    if (idResult.success) {
      console.log('✅ IDs assigned:', idResult.assignedIds);
    } else {
      console.log('⚠️ ID assignment warning:', idResult.error);
    }
    
    // Generate confirmation number
    const confirmationNumber = generateConfirmationNumber(testFormData.Student_Organization);
    addConfirmationNumberToSheet(sheet, writeResult.row, confirmationNumber);
    console.log('✅ Confirmation number:', confirmationNumber);
    
    console.log('\n🎉 Test submission created successfully!');
    console.log('📊 Check the TargetOrders sheet to see the new entry');
    console.log(`📍 Row: ${writeResult.row}`);
    console.log(`🎫 Confirmation: ${confirmationNumber}`);
    
    return {
      success: true,
      row: writeResult.row,
      confirmationNumber: confirmationNumber,
      ids: idResult.assignedIds
    };
    
  } catch (error) {
    console.error('❌ Test submission failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get deployment information
 */
function getDeploymentInfo() {
  const info = {
    scriptId: ScriptApp.getScriptId(),
    webAppUrl: ScriptApp.getService().getUrl(),
    spreadsheetId: PropertiesService.getScriptProperties().getProperty('key'),
    timezone: SYSTEM_CONFIG.TIMEZONE,
    driveFolder: SYSTEM_CONFIG.DRIVE_FOLDER_ID
  };
  
  console.log('📋 DEPLOYMENT INFORMATION');
console.log('📋 DEPLOYMENT INFORMATION');
  console.log('═══════════════════════════════════════════════════════');
  console.log('Script ID:', info.scriptId);
  console.log('Web App URL:', info.webAppUrl || 'Not deployed yet');
  console.log('Spreadsheet ID:', info.spreadsheetId);
  console.log('Timezone:', info.timezone);
  console.log('Drive Folder ID:', info.driveFolder);
  console.log('═══════════════════════════════════════════════════════');
  
  return info;
}

/**
 * Quick diagnostic check
 */
function quickDiagnostic() {
  console.log('🔍 Running Quick Diagnostic...\n');
  
  const checks = [];
  
  // Check 1: Spreadsheet
  try {
    const ss = getCachedSpreadsheet();
    checks.push({ check: 'Spreadsheet Access', status: '✅', details: ss.getName() });
  } catch (error) {
    checks.push({ check: 'Spreadsheet Access', status: '❌', details: error.message });
  }
  
  // Check 2: Sheets
  try {
    const sheetNames = ['TargetOrders', 'AmazonOrders', 'DocumentSubmissions', 'Names'];
    const missing = [];
    sheetNames.forEach(name => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
      if (!sheet) missing.push(name);
    });
    
    if (missing.length === 0) {
      checks.push({ check: 'Required Sheets', status: '✅', details: 'All present' });
    } else {
      checks.push({ check: 'Required Sheets', status: '❌', details: `Missing: ${missing.join(', ')}` });
    }
  } catch (error) {
    checks.push({ check: 'Required Sheets', status: '❌', details: error.message });
  }
  
  // Check 3: ID Management
  try {
    const props = PropertiesService.getScriptProperties();
    const hasKey = !!props.getProperty('key');
    checks.push({ check: 'ID Management Setup', status: hasKey ? '✅' : '⚠️', details: hasKey ? 'Configured' : 'Run masterSetup()' });
  } catch (error) {
    checks.push({ check: 'ID Management Setup', status: '❌', details: error.message });
  }
  
  // Check 4: Drive Folder
  try {
    const folder = DriveApp.getFolderById(SYSTEM_CONFIG.DRIVE_FOLDER_ID);
    checks.push({ check: 'Drive Folder Access', status: '✅', details: folder.getName() });
  } catch (error) {
    checks.push({ check: 'Drive Folder Access', status: '❌', details: 'Update DRIVE_FOLDER_ID in Configuration.gs' });
  }
  
  // Check 5: Web App Deployment
  try {
    const url = ScriptApp.getService().getUrl();
    checks.push({ check: 'Web App Deployment', status: url ? '✅' : '⚠️', details: url || 'Not deployed' });
  } catch (error) {
    checks.push({ check: 'Web App Deployment', status: '❌', details: error.message });
  }
  
  // Print results
  checks.forEach((check, index) => {
    console.log(`${index + 1}. ${check.status} ${check.check}`);
    console.log(`   ${check.details}`);
  });
  
  const allGood = checks.every(c => c.status === '✅');
  console.log('\n' + (allGood ? '🎉 All systems operational!' : '⚠️ Some issues detected. Review above.'));
  
  return checks;
}

/**
 * Reset system (use with caution - clears counters)
 */
function resetSystem() {
  const confirmation = Browser.msgBox(
    'Reset System',
    'This will reset all ID counters and cached data. Are you sure?',
    Browser.Buttons.YES_NO
  );
  
  if (confirmation !== 'yes') {
    console.log('❌ Reset cancelled by user');
    return { success: false, message: 'Reset cancelled' };
  }
  
  console.log('🔄 Resetting system...');
  
  try {
    // Clear script properties
    const props = PropertiesService.getScriptProperties();
    props.deleteAllProperties();
    console.log('✅ Cleared script properties');
    
    // Reinitialize
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    props.setProperty('key', spreadsheet.getId());
    console.log('✅ Reinitialized spreadsheet key');
    
    // Reinitialize counters
    initializeCountersFromExistingData();
    console.log('✅ Reinitialized ID counters from existing data');
    
    console.log('🎉 System reset complete');
    
    return {
      success: true,
      message: 'System reset successfully'
    };
    
  } catch (error) {
    console.error('❌ Reset failed:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Backup current configuration
 */
function backupConfiguration() {
  const config = {
    timestamp: new Date().toISOString(),
    spreadsheetId: PropertiesService.getScriptProperties().getProperty('key'),
    systemConfig: SYSTEM_CONFIG,
    scriptProperties: PropertiesService.getScriptProperties().getProperties()
  };
  
  console.log('💾 Configuration Backup:');
  console.log(JSON.stringify(config, null, 2));
  
  return config;
}

/**
 * Migrate old statuses to new status system
 * Run this ONCE after deploying the new code
 * 
 * IMPORTANT: This will update ALL existing orders in your sheets
 * Make sure you have a backup before running!
 */
function migrateStatusValues() {
  console.log('🔄 Starting status migration...');
  console.log('═══════════════════════════════════════════════════════');
  
  const spreadsheet = getCachedSpreadsheet();
  const sheetsToMigrate = ['TargetOrders', 'AmazonOrders', 'DocumentSubmissions'];
  
  const statusMapping = {
    // Completed variations → Completed
    'completed': 'Completed',
    'complete': 'Completed',
    'done': 'Completed',
    'fulfilled': 'Completed',
    
    // Delivered → Delivered to Mailroom
    'delivered': 'Delivered to Mailroom',
    'received': 'Delivered to Mailroom',
    
    // Processing/Ordered → Ordered
    'processing': 'Ordered',
    'ordered': 'Ordered',
    
    // Waiting/On Hold → Awaiting Club Response
    'waiting': 'Awaiting Club Response',
    'on hold': 'Awaiting Club Response',
    'awaiting': 'Awaiting Club Response',
    
    // Pending/New → New Order
    'pending': 'New Order',
    'new': 'New Order',
    
    // Cancelled variations → Cancelled
    'cancelled': 'Cancelled',
    'canceled': 'Cancelled',
    'void': 'Cancelled'
  };
  
  let totalUpdates = 0;
  const migrationReport = [];
  
  sheetsToMigrate.forEach(sheetName => {
    console.log(`\n📋 Migrating ${sheetName}...`);
    
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      console.log(`⚠️ Sheet ${sheetName} not found, skipping...`);
      migrationReport.push({
        sheet: sheetName,
        status: 'NOT FOUND',
        updates: 0
      });
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log(`📄 ${sheetName} is empty, skipping...`);
      migrationReport.push({
        sheet: sheetName,
        status: 'EMPTY',
        updates: 0
      });
      return;
    }
    
    const headers = data[0];
    const statusCol = findColumnIndex(headers, 'Order_Status');
    
    if (statusCol === -1) {
      console.log(`⚠️ Order_Status column not found in ${sheetName}`);
      migrationReport.push({
        sheet: sheetName,
        status: 'NO STATUS COLUMN',
        updates: 0
      });
      return;
    }
    
    let sheetUpdates = 0;
    const statusCounts = {};
    
    for (let row = 2; row <= data.length; row++) {
      const currentStatus = String(sheet.getRange(row, statusCol + 1).getValue()).trim();
      const normalizedCurrent = currentStatus.toLowerCase();
      
      if (statusMapping[normalizedCurrent]) {
        const newStatus = statusMapping[normalizedCurrent];
        sheet.getRange(row, statusCol + 1).setValue(newStatus);
        
        // Track what we're changing
        if (!statusCounts[currentStatus]) {
          statusCounts[currentStatus] = { count: 0, newStatus: newStatus };
        }
        statusCounts[currentStatus].count++;
        
        console.log(`  Row ${row}: "${currentStatus}" → "${newStatus}"`);
        sheetUpdates++;
      } else if (currentStatus && currentStatus !== '') {
        console.log(`  ⚠️ Row ${row}: Unknown status "${currentStatus}" - skipped`);
      }
    }
    
    console.log(`\n📊 ${sheetName} Summary:`);
    Object.keys(statusCounts).forEach(oldStatus => {
      console.log(`  • "${oldStatus}" (${statusCounts[oldStatus].count} rows) → "${statusCounts[oldStatus].newStatus}"`);
    });
    console.log(`✅ Total updated: ${sheetUpdates} rows\n`);
    
    migrationReport.push({
      sheet: sheetName,
      status: 'SUCCESS',
      updates: sheetUpdates,
      details: statusCounts
    });
    
    totalUpdates += sheetUpdates;
  });
  
  console.log('\n═══════════════════════════════════════════════════════');
  console.log('🎉 MIGRATION COMPLETE!');
  console.log('═══════════════════════════════════════════════════════');
  console.log(`📊 Total rows updated: ${totalUpdates}`);
  console.log('\n📋 Summary by Sheet:');
  
  migrationReport.forEach(report => {
    console.log(`\n${report.sheet}:`);
    console.log(`  Status: ${report.status}`);
    console.log(`  Updates: ${report.updates}`);
    if (report.details) {
      console.log(`  Changes:`);
      Object.keys(report.details).forEach(oldStatus => {
        console.log(`    "${oldStatus}" → "${report.details[oldStatus].newStatus}" (${report.details[oldStatus].count} rows)`);
      });
    }
  });
  
  console.log('\n✅ Migration complete! You can now use the new status system.');
  console.log('💡 Tip: Refresh your Order Management UI to see the changes.');
  
  return {
    success: true,
    totalUpdates: totalUpdates,
    report: migrationReport,
    message: `Successfully migrated ${totalUpdates} status values across ${migrationReport.length} sheets`
  };
}
