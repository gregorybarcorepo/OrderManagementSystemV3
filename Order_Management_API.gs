/**
 * ============================================
 * ORDER MANAGEMENT API - V3 with Logging
 * Backend API for web-based Order Management UI
 * ============================================
 */

/**
 * Serve the Order Management web interface
 * @returns {HtmlOutput} Web page
 */
function doGet() {
  const htmlOutput = HtmlService.createTemplateFromFile('Display_UI_Webpage');
  return htmlOutput.evaluate()
    .setTitle('Order Management System v3')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include external files in HTML
 * @param {string} filename - File to include
 * @returns {string} File content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========================================
// DATA RETRIEVAL FUNCTIONS
// ========================================

/**
 * Get all orders with optimized processing and logging
 * @returns {Object} Orders data with metadata
 */
function getAllOrders() {
  const startTime = Date.now();
  
  try {
    console.log('📊 Getting all orders...');
    
    const spreadsheet = getCachedSpreadsheet();
    
    // Parallel sheet access
    const amazonSheet = spreadsheet.getSheetByName('AmazonOrders');
    const targetSheet = spreadsheet.getSheetByName('TargetOrders');
    
    // Get raw data
    const amazonData = amazonSheet ? amazonSheet.getDataRange().getValues() : [[]];
    const targetData = targetSheet ? targetSheet.getDataRange().getValues() : [[]];
    
    // Process on backend
    const amazonOrders = processAmazonOrdersOptimized(amazonData);
    const targetOrders = processTargetOrdersOptimized(targetData);
    
    const elapsed = Date.now() - startTime;
    const result = {
      success: true,
      amazon: amazonOrders,
      target: targetOrders,
      total: amazonOrders.length + targetOrders.length,
      message: 'Orders loaded successfully',
      loadTimeMs: elapsed
    };
    
    // Log success
    logAction('ORDERS_LOADED', { 
      amazonCount: amazonOrders.length,
      targetCount: targetOrders.length,
      totalCount: result.total,
      loadTimeMs: elapsed 
    });
    
    console.log(`✅ Retrieved ${result.total} orders in ${elapsed}ms`);
    return result;
    
  } catch (error) {
    console.error('❌ Error getting orders:', error);
    
    // Log error
    logAction('ORDERS_LOAD_FAILED', { 
      error: error.message,
      stack: error.stack 
    }, 'ERROR');
    
    return {
      success: false,
      amazon: [],
      target: [],
      total: 0,
      message: `Failed to get orders: ${error.message}`
    };
  }
}
/**
 * Process Amazon orders with frontend-ready mappings
 * @param {Array} data - Raw sheet data
 * @returns {Array} Processed orders
 */
function processAmazonOrdersOptimized(data) {
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const orders = [];
  
  // Pre-compute column indices once
  const colIndices = {
    timestamp: headers.indexOf('Timestamp'),
    submissionId: headers.indexOf('Submission_Id'),
    historicalId: headers.indexOf('Historical_Submission_Id'),
    organization: headers.indexOf('Student_Organization'),
    eventName: headers.indexOf('Event_Name'),
    eventDate: headers.indexOf('Event_Date'),
    pickupPerson: headers.indexOf('Pickup_Person_Name'),
    pickupEmail: headers.indexOf('Pickup_Person_Email'),
    pickupPhone: headers.indexOf('Pickup_Person_Phone'),
    submissionEmail: headers.indexOf('Submission_Email_Address'),
    wishlistLink: headers.indexOf('Wishlist_Link'),
    orderNumbers: headers.indexOf('Associated Order Numbers'),
    total1: headers.indexOf('Total_1'),
    total2: headers.indexOf('Total_2'),
    totalOrder: headers.indexOf('Total_Order'),
    backupItems: headers.indexOf('Backup_Items_And_Quantity'),
    formNotes: headers.indexOf('Form Submitter Notes'),
    processedBy: headers.indexOf('Processed_By'),
    timeProcessedBy: headers.indexOf('Time_Processed_By'),
    comments: headers.indexOf('Comments'),
    status: headers.indexOf('Order_Status'),
    pickedUp: headers.indexOf('Picked_Up_Status'),
    nonPO: headers.indexOf('Non_PO_Submitted'),
    confirmationNumber: headers.indexOf('Confirmation_Number')
  };
  
  // ✅ Find all item columns dynamically
  const itemColumns = [];
  const quantityColumns = [];
  
  headers.forEach((header, index) => {
    const headerStr = String(header).toLowerCase();
    
    // Match Item_1, Item_2, etc.
    if (headerStr.match(/^item[_\s]*\d+$/)) {
      const itemNum = parseInt(headerStr.match(/\d+/)[0]);
      itemColumns.push({ index: index, number: itemNum });
    }
    
    // Match Item_1_Quantity, Quantity_1, etc.
    if (headerStr.match(/quantity[_\s]*\d+/) || headerStr.match(/item[_\s]*\d+[_\s]*quantity/)) {
      const qtyNum = parseInt(headerStr.match(/\d+/)[0]);
      quantityColumns.push({ index: index, number: qtyNum });
    }
  });
  
  // Sort by item number
  itemColumns.sort((a, b) => a.number - b.number);
  quantityColumns.sort((a, b) => a.number - b.number);
  
  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[colIndices.organization] && !row[colIndices.eventName] && !row[colIndices.submissionId]) {
      continue;
    }
    
    // ✅ Extract items
    const items = [];
    itemColumns.forEach((itemCol, idx) => {
      const itemName = row[itemCol.index];
      const qtyCol = quantityColumns[idx];
      const quantity = qtyCol ? row[qtyCol.index] : '';
      
      if (itemName && String(itemName).trim() !== '') {
        items.push({
          name: String(itemName).trim(),
          quantity: quantity ? String(quantity).trim() : '1'
        });
      }
    });
    
    // Build order with frontend-ready property names
    const order = {
      // Platform
      platform: 'amazon',
      
      // IDs - guaranteed to have at least one
      id: row[colIndices.submissionId] || row[colIndices.historicalId] || `temp_amazon_${i}`,
      submissionId: row[colIndices.submissionId] || '',
      historicalId: row[colIndices.historicalId] || '',
      confirmationNumber: row[colIndices.confirmationNumber] || '',
      
      // Core info
      organization: formatValue(row[colIndices.organization]),
      eventName: formatValue(row[colIndices.eventName]),
      eventDate: formatDate(row[colIndices.eventDate]),
      timestamp: formatDate(row[colIndices.timestamp]),
      
      // Contact info
      pickupPerson: formatValue(row[colIndices.pickupPerson]),
      pickupEmail: formatValue(row[colIndices.pickupEmail]),
      pickupPhone: formatValue(row[colIndices.pickupPhone]),
      submissionEmail: formatValue(row[colIndices.submissionEmail]),
      
      // Amazon-specific
      wishlistLink: formatValue(row[colIndices.wishlistLink]),
      orderNumbers: formatValue(row[colIndices.orderNumbers]),
      Total_1: parseFloat(row[colIndices.total1]) || 0,
      Total_2: parseFloat(row[colIndices.total2]) || 0,
      totalOrder: parseFloat(row[colIndices.totalOrder]) || 0,
      backupItems: formatValue(row[colIndices.backupItems]),
      formNotes: formatValue(row[colIndices.formNotes]),
      
      // ✅ Items array
      items: items,
      
      // Staff fields
      processedBy: formatValue(row[colIndices.processedBy]),
      timeProcessedBy: formatDate(row[colIndices.timeProcessedBy]),
      comments: formatValue(row[colIndices.comments]),
      pickedUpStatus: formatValue(row[colIndices.pickedUp]),
      nonPOSubmitted: formatValue(row[colIndices.nonPO]),
      
      // Status - normalized to kebab-case
      status: normalizeStatusForFrontend(row[colIndices.status])
    };
    
    orders.push(order);
  }
  
  return orders;
}

/**
 * Process Target orders with frontend-ready mappings
 * @param {Array} data - Raw sheet data
 * @returns {Array} Processed orders
 */
function processTargetOrdersOptimized(data) {
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const orders = [];
  
  // Pre-compute column indices
  const colIndices = {
    timestamp: headers.indexOf('Timestamp'),
    submissionId: headers.indexOf('Submission_Id'),
    historicalId: headers.indexOf('Historical_Submission_Id'),
    organization: headers.indexOf('Student_Organization'),
    eventName: headers.indexOf('Event_Name'),
    eventDate: headers.indexOf('Event_Date'),
    pickupPerson: headers.indexOf('Pickup_Person_Name'),
    pickupEmail: headers.indexOf('Pickup_Person_Email'),
    pickupPhone: headers.indexOf('Pickup_Person_Phone'),
    submissionEmail: headers.indexOf('Submission_Email_Address'),
    cartTotal: headers.indexOf('Cart_Total'),
    orderConfirmation: headers.indexOf('Order_Confirmation_Number'),
    backupItems: headers.indexOf('Backup_Items_And_Quantity'),
    formNotes: headers.indexOf('Form Submitter Notes'),
    processedBy: headers.indexOf('Processed_By'),
    timeProcessedBy: headers.indexOf('Time_Processed_By'),
    comments: headers.indexOf('Comments'),
    status: headers.indexOf('Order_Status'),
    pickedUp: headers.indexOf('Picked_Up_Status'),
    confirmationNumber: headers.indexOf('Confirmation_Number')
  };
  
  // ✅ Find all item columns dynamically (Item_1, Item_2, Item_3, etc.)
  const itemColumns = [];
  const quantityColumns = [];
  
  headers.forEach((header, index) => {
    const headerStr = String(header).toLowerCase();
    
    // Match Item_1, Item_2, etc.
    if (headerStr.match(/^item[_\s]*\d+$/)) {
      const itemNum = parseInt(headerStr.match(/\d+/)[0]);
      itemColumns.push({ index: index, number: itemNum });
    }
    
    // Match Item_1_Quantity, Quantity_1, etc.
    if (headerStr.match(/quantity[_\s]*\d+/) || headerStr.match(/item[_\s]*\d+[_\s]*quantity/)) {
      const qtyNum = parseInt(headerStr.match(/\d+/)[0]);
      quantityColumns.push({ index: index, number: qtyNum });
    }
  });
  
  // Sort by item number
  itemColumns.sort((a, b) => a.number - b.number);
  quantityColumns.sort((a, b) => a.number - b.number);
  
  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[colIndices.organization] && !row[colIndices.eventName] && !row[colIndices.submissionId]) {
      continue;
    }
    
    // ✅ Extract items
    const items = [];
    itemColumns.forEach((itemCol, idx) => {
      const itemName = row[itemCol.index];
      const qtyCol = quantityColumns[idx];
      const quantity = qtyCol ? row[qtyCol.index] : '';
      
      if (itemName && String(itemName).trim() !== '') {
        items.push({
          name: String(itemName).trim(),
          quantity: quantity ? String(quantity).trim() : '1'
        });
      }
    });
    
    // Build order with frontend-ready property names
    const order = {
      // Platform
      platform: 'target',
      
      // IDs - guaranteed to have at least one
      id: row[colIndices.submissionId] || row[colIndices.historicalId] || `temp_target_${i}`,
      submissionId: row[colIndices.submissionId] || '',
      historicalId: row[colIndices.historicalId] || '',
      confirmationNumber: row[colIndices.confirmationNumber] || '',
      
      // Core info
      organization: formatValue(row[colIndices.organization]),
      eventName: formatValue(row[colIndices.eventName]),
      eventDate: formatDate(row[colIndices.eventDate]),
      timestamp: formatDate(row[colIndices.timestamp]),
      
      // Contact info
      pickupPerson: formatValue(row[colIndices.pickupPerson]),
      pickupEmail: formatValue(row[colIndices.pickupEmail]),
      pickupPhone: formatValue(row[colIndices.pickupPhone]),
      submissionEmail: formatValue(row[colIndices.submissionEmail]),
      
      // Target-specific
      cartTotal: parseFloat(row[colIndices.cartTotal]) || 0,
      orderConfirmation: formatValue(row[colIndices.orderConfirmation]),
      backupItems: formatValue(row[colIndices.backupItems]),
      formNotes: formatValue(row[colIndices.formNotes]),
      
      // ✅ Items array
      items: items,
      
      // Staff fields
      processedBy: formatValue(row[colIndices.processedBy]),
      timeProcessedBy: formatDate(row[colIndices.timeProcessedBy]),
      comments: formatValue(row[colIndices.comments]),
      pickedUpStatus: formatValue(row[colIndices.pickedUp]),
      
      // Status - normalized to kebab-case
      status: normalizeStatusForFrontend(row[colIndices.status])
    };
    
    orders.push(order);
  }
  
  return orders;
}

/**
 * Get document submissions - FIXED for 4 files only
 * @returns {Object} Document submissions data
 */
function getDocumentSubmissions() {
  const startTime = Date.now();
  
  try {
    const spreadsheet = getCachedSpreadsheet();
    const sheet = spreadsheet.getSheetByName('DocumentSubmissions');
    
    if (!sheet) {
      logAction('DOCUMENTS_LOAD_FAILED', { error: 'Sheet not found' }, 'WARNING');
      return {
        success: false,
        documents: [],
        message: 'DocumentSubmissions sheet not found'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        documents: [],
        message: 'No document submissions found'
      };
    }
    
    const headers = data[0];
    const documents = [];
    
    // Pre-compute column indices
    const colIndices = {
      timestamp: headers.indexOf('Timestamp'),
      submissionId: headers.indexOf('Submission_Id'),
      historicalId: headers.indexOf('Historical_Submission_Id'),
      firstName: headers.indexOf('First_Name'),
      lastName: headers.indexOf('Last_Name'),
      email: headers.indexOf('Email_Address') >= 0 ? headers.indexOf('Email_Address') : headers.indexOf('Submission_Email_Address'),
      organization: headers.indexOf('Student_Organization'),
      eventName: headers.indexOf('Event_Name'),
      eventDate: headers.indexOf('Event_Date'),
      nonPOFile: headers.indexOf('nonPO_FileLink'),
      signInFile: headers.indexOf('signIn_FileLink'),
      invoiceFile: headers.indexOf('invoice_FileLink'),
      eventFlyerFile: headers.indexOf('eventFlyer_FileLink'),
      // ✅ REMOVED: officialReceiptFile - only 4 files now
      confirmationNumber: headers.indexOf('Confirmation_Number'),
      notes: headers.indexOf('Form Submitter Notes'),
      comments: headers.indexOf('Comments')
    };
    
    // Process rows
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[colIndices.firstName] && !row[colIndices.lastName] && !row[colIndices.organization]) {
        continue;
      }
      
      documents.push({
        // IDs
        id: row[colIndices.submissionId] || row[colIndices.historicalId] || `doc_${i}`,
        submissionId: row[colIndices.submissionId] || '',
        historicalId: row[colIndices.historicalId] || '',
        
        // Timestamps
        timestamp: formatDate(row[colIndices.timestamp]),
        
        // Person info
        firstName: formatValue(row[colIndices.firstName]),
        lastName: formatValue(row[colIndices.lastName]),
        email: formatValue(row[colIndices.email]),
        
        // Event info
        organization: formatValue(row[colIndices.organization]),
        eventName: formatValue(row[colIndices.eventName]),
        eventDate: formatDate(row[colIndices.eventDate]),
        
        // Files - ONLY 4 FILES
        nonPOFile: formatValue(row[colIndices.nonPOFile]),
        signInFile: formatValue(row[colIndices.signInFile]),
        invoiceFile: formatValue(row[colIndices.invoiceFile]),
        eventFlyerFile: formatValue(row[colIndices.eventFlyerFile]),
        // ✅ NO officialReceiptFile
        
        // Metadata
        confirmationNumber: formatValue(row[colIndices.confirmationNumber]),
        notes: formatValue(row[colIndices.notes]),
        comments: formatValue(row[colIndices.comments])
      });
    }
    
    const elapsed = Date.now() - startTime;
    
    logAction('DOCUMENTS_LOADED', {
      count: documents.length,
      loadTimeMs: elapsed
    });
    
    console.log(`✅ Loaded ${documents.length} document submissions in ${elapsed}ms`);
    
    return {
      success: true,
      documents: documents,
      message: `Found ${documents.length} document submissions`,
      loadTimeMs: elapsed
    };
    
  } catch (error) {
    console.error('❌ Error getting documents:', error);
    
    logAction('DOCUMENTS_LOAD_FAILED', {
      error: error.message
    }, 'ERROR');
    
    return {
      success: false,
      documents: [],
      message: `Error: ${error.toString()}`
    };
  }
}

// ========================================
// HELPER FUNCTIONS
// ========================================

/**
 * Format value - handles nulls, dates, empty strings
 */
function formatValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'America/New_York', 'MM/dd/yyyy');
  }
  return String(value).trim();
}

/**
 * Format date value
 */
function formatDate(value) {
  if (!value || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'America/New_York', 'MM/dd/yyyy HH:mm:ss');
  }
  return String(value).trim();
}

/**
 * Normalize status to kebab-case for frontend
 */
function normalizeStatusForFrontend(status) {
  if (!status) return 'new-order';
  
  const s = String(status).trim().toLowerCase();
  
  // Direct mapping
  const mapping = {
    'new order': 'new-order',
    'pending': 'new-order',
    'ordered': 'ordered',
    'processing': 'ordered',
    'delivered to mailroom': 'delivered-to-mailroom',
    'delivered': 'delivered-to-mailroom',
    'completed': 'completed',
    'complete': 'completed',
    'awaiting club response': 'awaiting-club-response',
    'on hold': 'awaiting-club-response',
    'cancelled': 'cancelled',
    'canceled': 'cancelled'
  };
  
  return mapping[s] || s.replace(/\s+/g, '-');
}
/**
 * TEST FUNCTION - Run this to verify the fix works
 * This will log the first few documents to verify field mapping
 */
function testDocumentSubmissionsDebug() {
  console.log('🧪 Testing getDocumentSubmissions()...');
  console.log('═══════════════════════════════════════');
  
  const result = getDocumentSubmissions();
  
  console.log('\n📊 Result:');
  console.log('Success:', result.success);
  console.log('Message:', result.message);
  console.log('Document count:', result.documents.length);
  
  if (result.documents.length > 0) {
    console.log('\n📄 First document sample:');
    const firstDoc = result.documents[0];
    console.log(JSON.stringify(firstDoc, null, 2));
    
    console.log('\n✅ Field Verification:');
    console.log('- ID:', firstDoc.id ? '✓' : '✗');
    console.log('- Timestamp:', firstDoc.timestamp ? '✓' : '✗');
    console.log('- First Name:', firstDoc.firstName ? '✓' : '✗');
    console.log('- Last Name:', firstDoc.lastName ? '✓' : '✗');
    console.log('- Email:', firstDoc.email ? '✓' : '✗');
    console.log('- Organization:', firstDoc.organization ? '✓' : '✗');
    console.log('- Event Name:', firstDoc.eventName ? '✓' : '✗');
    console.log('- nonPO File:', firstDoc.nonPOFile ? '✓' : '✗');
  }
  
  console.log('\n═══════════════════════════════════════');
  console.log('🎉 Test complete!');
  
  return result;
}

/**
 * Get club names and employee names
 * @returns {Object} Names data
 */
function getNamesData() {
  try {
    const spreadsheet = getCachedSpreadsheet();
    const namesSheet = spreadsheet.getSheetByName('Names');
    
    if (!namesSheet) {
      return {
        success: false,
        clubs: [],
        employees: [],
        message: 'Names sheet not found'
      };
    }
    
    const data = namesSheet.getDataRange().getValues();
    const headers = data[0];
    
    const clubNamesIndex = headers.findIndex(h => 
      h.toString().toLowerCase().includes('club')
    );
    const employeeNamesIndex = headers.findIndex(h => 
      h.toString().toLowerCase().includes('employee')
    );
    
    const clubs = [];
    const employees = [];
    
    for (let i = 1; i < data.length; i++) {
      if (clubNamesIndex !== -1 && data[i][clubNamesIndex]) {
        clubs.push(data[i][clubNamesIndex].toString().trim());
      }
      if (employeeNamesIndex !== -1 && data[i][employeeNamesIndex]) {
        employees.push(data[i][employeeNamesIndex].toString().trim());
      }
    }
    
    const uniqueClubs = [...new Set(clubs)].filter(c => c).sort();
    const uniqueEmployees = [...new Set(employees)].filter(e => e).sort();
    
    return {
      success: true,
      clubs: uniqueClubs,
      employees: uniqueEmployees,
      message: `Loaded ${uniqueClubs.length} clubs and ${uniqueEmployees.length} employees`
    };
  } catch (error) {
    console.error('Error in getNamesData:', error);
    return {
      success: false,
      clubs: [],
      employees: [],
      message: error.toString()
    };
  }
}

/**
 * Get dashboard totals and statistics
 * @returns {Object} Dashboard data
 */
function getOrderTotals() {
  try {
    const spreadsheet = getCachedSpreadsheet();
    
    // Amazon Orders
    const amazonSheet = spreadsheet.getSheetByName('AmazonOrders');
    const amazonData = amazonSheet ? amazonSheet.getDataRange().getValues() : [[]];
    const amazonHeaders = amazonData[0];
    const amazonRows = amazonData.slice(1).filter(r => r.join("").trim() !== "");
    const amazonCounts = countStatuses(amazonRows, amazonHeaders);
    const amazonSpent = calculateTotalSpent(amazonRows, amazonHeaders, 'Total_Order');
    
    // Target Orders
    const targetSheet = spreadsheet.getSheetByName('TargetOrders');
    const targetData = targetSheet ? targetSheet.getDataRange().getValues() : [[]];
    const targetHeaders = targetData[0];
    const targetRows = targetData.slice(1).filter(r => r.join("").trim() !== "");
    const targetCounts = countStatuses(targetRows, targetHeaders);
    const targetSpent = calculateTotalSpent(targetRows, targetHeaders, 'Cart_Total');
    
    // Totals
    const totalOrders = amazonRows.length + targetRows.length;
    const allKeys = new Set([...Object.keys(amazonCounts), ...Object.keys(targetCounts)]);
    const totalCounts = {};
    allKeys.forEach(k => {
      totalCounts[k] = (amazonCounts[k] || 0) + (targetCounts[k] || 0);
    });
    
    return {
      success: true,
      amazonOrders: amazonRows.length,
      targetOrders: targetRows.length,
      totalOrders: totalOrders,
      amazonSpent: amazonSpent,
      targetSpent: targetSpent,
      totalSpent: amazonSpent + targetSpent,
      statuses: {
        amazon: amazonCounts,
        target: targetCounts,
        total: totalCounts
      }
    };
  } catch (error) {
    console.error('Error getting order totals:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Count statuses in order rows
 */
function countStatuses(rows, headers) {
  const statusIndex = headers.indexOf("Order_Status");
  const counts = {};
  
  rows.forEach(row => {
    let raw = (statusIndex !== -1) ? row[statusIndex] : null;
    const status = normalizeStatusForFrontend(raw);
    counts[status] = (counts[status] || 0) + 1;
  });
  
  return counts;
}

/**
 * Calculate total spent
 */
function calculateTotalSpent(rows, headers, columnName) {
  const columnIndex = headers.indexOf(columnName);
  if (columnIndex === -1) return 0;
  
  return rows.reduce((sum, row) => {
    const value = parseMoneyValue(row[columnIndex]);
    return sum + value;
  }, 0);
}

// ========================================
// DATA UPDATE FUNCTIONS
// ========================================

/**
 * Update order status with logging
 */
function updateOrderStatus(orderId, platform, newStatus) {
  try {
    console.log(`Updating status: ID=${orderId}, Platform=${platform}, Status=${newStatus}`);
    
    // Validate status
    const statusMap = {
      'new-order': 'New Order',
      'ordered': 'Ordered',
      'delivered-to-mailroom': 'Delivered to Mailroom',
      'completed': 'Completed',
      'awaiting-club-response': 'Awaiting Club Response',
      'cancelled': 'Cancelled'
    };
    
    const normalizedStatus = newStatus.toLowerCase().replace(/\s+/g, '-');
    
    if (!statusMap[normalizedStatus]) {
      throw new Error(`Invalid status: ${newStatus}`);
    }
    
    const displayStatus = statusMap[normalizedStatus];
    
    const sheetMap = {
      'amazon': 'AmazonOrders',
      'target': 'TargetOrders',
      'document': 'DocumentSubmissions'
    };
    
    const sheetName = sheetMap[platform];
    if (!sheetName) {
      throw new Error(`Invalid platform: ${platform}`);
    }
    
    const sheet = getCachedSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find row
    const historicalIdIndex = findColumnIndex(headers, 'Historical_Submission_Id');
    const submissionIdIndex = findColumnIndex(headers, 'Submission_Id');
    let foundRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      const historicalId = historicalIdIndex !== -1 ? data[i][historicalIdIndex] : null;
      const submissionId = submissionIdIndex !== -1 ? data[i][submissionIdIndex] : null;
      
      if (historicalId == orderId || submissionId == orderId) {
        foundRow = i;
        break;
      }
    }
    
    if (foundRow === -1) {
      throw new Error(`Order with ID ${orderId} not found`);
    }
    
    // Find or create status column
    let statusColumnIndex = findColumnIndex(headers, 'Order_Status');
    if (statusColumnIndex === -1) {
      statusColumnIndex = headers.length;
      sheet.getRange(1, statusColumnIndex + 1).setValue('Order_Status');
    }
    
    // Update status
    sheet.getRange(foundRow + 1, statusColumnIndex + 1).setValue(displayStatus);
    
    // Log success
    logAction('ORDER_STATUS_UPDATED', {
      orderId: orderId,
      platform: platform,
      oldStatus: data[foundRow][statusColumnIndex] || 'Unknown',
      newStatus: displayStatus
    });
    
    console.log(`✅ Updated status for row ${foundRow + 1} to: ${displayStatus}`);
    
    return {
      success: true,
      message: `Order status updated to: ${displayStatus}`
    };
    
  } catch (error) {
    console.error('Error updating order status:', error);
    
    logAction('ORDER_STATUS_UPDATE_FAILED', {
      orderId: orderId,
      platform: platform,
      error: error.message
    }, 'ERROR');
    
    return { success: false, message: error.message };
  }
}

/**
 * Update order details with logging
 */
function updateOrderDetails(orderId, platform, updatedData) {
  try {
    console.log(`Updating ${platform} order ${orderId}:`, Object.keys(updatedData));
    
    const sheetMap = {
      'amazon': 'AmazonOrders',
      'target': 'TargetOrders'
    };
    
    const sheetName = sheetMap[platform];
    const sheet = getCachedSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find row
    const historicalIdIndex = findColumnIndex(headers, 'Historical_Submission_Id');
    const submissionIdIndex = findColumnIndex(headers, 'Submission_Id');
    let foundRow = -1;
    
    if (historicalIdIndex !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][historicalIdIndex] == orderId) {
          foundRow = i;
          break;
        }
      }
    }
    
    if (foundRow === -1 && submissionIdIndex !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][submissionIdIndex] == orderId) {
          foundRow = i;
          break;
        }
      }
    }
    
    if (foundRow === -1) {
      throw new Error(`Order with ID ${orderId} not found`);
    }
    
    // Update each field
    const updatedFields = [];
    Object.keys(updatedData).forEach(key => {
      let columnIndex = findColumnIndex(headers, key);
      
      if (columnIndex === -1) {
        columnIndex = headers.length;
        sheet.getRange(1, columnIndex + 1).setValue(key);
      }
      
      sheet.getRange(foundRow + 1, columnIndex + 1).setValue(updatedData[key]);
      updatedFields.push(key);
    });
    
    // Log success
    logAction('ORDER_DETAILS_UPDATED', {
      orderId: orderId,
      platform: platform,
      fieldsUpdated: updatedFields
    });
    
    console.log(`✅ Updated order ${orderId}`);
    
    return {
      success: true,
      message: 'Order details updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating order details:', error);
    
    logAction('ORDER_UPDATE_FAILED', {
      orderId: orderId,
      platform: platform,
      error: error.message
    }, 'ERROR');
    
    return {
      success: false,
      message: `Failed to update order: ${error.message}`
    };
  }
}

/**
 * Update document submission
 */
function updateDocumentSubmission(docId, updatedData) {
  try {
    const spreadsheet = getCachedSpreadsheet();
    const sheet = spreadsheet.getSheetByName('DocumentSubmissions');
    
    if (!sheet) {
      throw new Error('DocumentSubmissions sheet not found');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idIndex = headers.indexOf('Submission_Id');
    const historicalIdIndex = headers.indexOf('Historical_Submission_Id');
    
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      const submissionId = data[i][idIndex];
      const historicalId = historicalIdIndex !== -1 ? data[i][historicalIdIndex] : null;
      
      if (submissionId == docId || historicalId == docId) {
        foundRow = i;
        break;
      }
    }
    
    if (foundRow === -1) {
      throw new Error(`Document with ID ${docId} not found`);
    }
    
    // Update fields
    const updatedFields = [];
    Object.keys(updatedData).forEach(columnName => {
      const colIndex = headers.indexOf(columnName);
      if (colIndex !== -1) {
        sheet.getRange(foundRow + 1, colIndex + 1).setValue(updatedData[columnName]);
        updatedFields.push(columnName);
      }
    });
    
    logAction('DOCUMENT_UPDATED', {
      docId: docId,
      fieldsUpdated: updatedFields
    });
    
    return {
      success: true,
      message: 'Document submission updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating document:', error);
    
    logAction('DOCUMENT_UPDATE_FAILED', {
      docId: docId,
      error: error.message
    }, 'ERROR');
    
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Test spreadsheet connection
 */
function testSpreadsheetConnection() {
  try {
    const spreadsheet = getCachedSpreadsheet();
    const sheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    return {
      success: true,
      message: 'Connection successful',
      spreadsheetName: spreadsheet.getName(),
      availableSheets: sheets,
      url: spreadsheet.getUrl()
    };
  } catch (error) {
    console.error('Connection test failed:', error);
    return {
      success: false,
      message: `Connection failed: ${error.message}`
    };
  }
}

