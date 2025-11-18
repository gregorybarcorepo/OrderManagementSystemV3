/**
 * ============================================
 * ORDER MANAGEMENT API - V3 with New Features
 * Backend API for web-based Order Management UI
 * 
 * NEW FEATURES:
 * - Target items display fix
 * - Manual order creation
 * - Manual document submission
 * - Comments viewing and editing
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
  
  // Find all item columns dynamically
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
    
    // Extract items
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
      
      // Items array
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
 * Process Target orders with frontend-ready mappings - FIXED for item extraction
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
  
  // ✅ FIX: Find all item URL and quantity columns dynamically
  const itemUrlColumns = [];
  const itemQuantityColumns = [];
  
  headers.forEach((header, index) => {
    const headerStr = String(header).toLowerCase().replace(/[_\s]/g, '');
    
    // Match patterns like: First_Item_Url, Second_Item_Url, Third_Item_Url, etc.
    // OR Item_1_Url, Item_2_Url, etc.
    if (headerStr.includes('item') && headerStr.includes('url')) {
      // Try to extract order number from header
      let itemNum = 0;
      
      // Check for written numbers (First, Second, Third, etc.)
      const writtenNumbers = {
        'first': 1, 'second': 2, 'third': 3, 'fourth': 4, 'fifth': 5,
        'sixth': 6, 'seventh': 7, 'eighth': 8, 'ninth': 9, 'tenth': 10
      };
      
      for (const [word, num] of Object.entries(writtenNumbers)) {
        if (headerStr.includes(word)) {
          itemNum = num;
          break;
        }
      }
      
      // If no written number, try to extract numeric value
      if (itemNum === 0) {
        const numMatch = headerStr.match(/\d+/);
        if (numMatch) {
          itemNum = parseInt(numMatch[0]);
        }
      }
      
      itemUrlColumns.push({ index: index, number: itemNum || itemUrlColumns.length + 1 });
    }
    
    // Match quantity columns: First_Item_Quantity, Item_1_Quantity, etc.
    if (headerStr.includes('item') && headerStr.includes('quantity')) {
      let itemNum = 0;
      
      const writtenNumbers = {
        'first': 1, 'second': 2, 'third': 3, 'fourth': 4, 'fifth': 5,
        'sixth': 6, 'seventh': 7, 'eighth': 8, 'ninth': 9, 'tenth': 10
      };
      
      for (const [word, num] of Object.entries(writtenNumbers)) {
        if (headerStr.includes(word)) {
          itemNum = num;
          break;
        }
      }
      
      if (itemNum === 0) {
        const numMatch = headerStr.match(/\d+/);
        if (numMatch) {
          itemNum = parseInt(numMatch[0]);
        }
      }
      
      itemQuantityColumns.push({ index: index, number: itemNum || itemQuantityColumns.length + 1 });
    }
  });
  
  // Sort by item number
  itemUrlColumns.sort((a, b) => a.number - b.number);
  itemQuantityColumns.sort((a, b) => a.number - b.number);
  
  console.log(`Found ${itemUrlColumns.length} item URL columns and ${itemQuantityColumns.length} quantity columns`);
  
  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[colIndices.organization] && !row[colIndices.eventName] && !row[colIndices.submissionId]) {
      continue;
    }
    
    // ✅ Extract items with URLs and quantities
    const items = [];
    itemUrlColumns.forEach((urlCol, idx) => {
      const itemUrl = row[urlCol.index];
      const qtyCol = itemQuantityColumns.find(q => q.number === urlCol.number);
      const quantity = qtyCol ? row[qtyCol.index] : '';
      
      if (itemUrl && String(itemUrl).trim() !== '') {
        items.push({
          name: String(itemUrl).trim(), // Use URL as name for Target
          quantity: quantity ? String(quantity).trim() : '1',
          url: String(itemUrl).trim()
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
      
      // ✅ Items array with URLs
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
 * Get document submissions
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

// ========================================
// DATA UPDATE FUNCTIONS
// ========================================

/**
 * Update order status with employee tracking
 */
function updateOrderStatus(orderId, platform, newStatus, processedBy) {
  try {
    console.log(`Updating status: ID=${orderId}, Platform=${platform}, Status=${newStatus}, By=${processedBy}`);
    
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
    
    // Validate platform
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
    
    // ✅ NEW: Find or create all needed columns
    let statusColumnIndex = findColumnIndex(headers, 'Order_Status');
    if (statusColumnIndex === -1) {
      statusColumnIndex = headers.length;
      sheet.getRange(1, statusColumnIndex + 1).setValue('Order_Status');
    }
    
    let processedByIndex = findColumnIndex(headers, 'Processed_By');
    if (processedByIndex === -1) {
      processedByIndex = headers.length;
      sheet.getRange(1, processedByIndex + 1).setValue('Processed_By');
    }
    
    let timeProcessedIndex = findColumnIndex(headers, 'Time_Processed_By');
    if (timeProcessedIndex === -1) {
      timeProcessedIndex = headers.length;
      sheet.getRange(1, timeProcessedIndex + 1).setValue('Time_Processed_By');
    }
    
    // ✅ UPDATE: Set all three values
    sheet.getRange(foundRow + 1, statusColumnIndex + 1).setValue(displayStatus);
    sheet.getRange(foundRow + 1, processedByIndex + 1).setValue(processedBy || 'Unknown');
    sheet.getRange(foundRow + 1, timeProcessedIndex + 1).setValue(getFormattedTimestamp());
    
    // Log success
    logAction('ORDER_STATUS_UPDATED', {
      orderId: orderId,
      platform: platform,
      oldStatus: data[foundRow][statusColumnIndex] || 'Unknown',
      newStatus: displayStatus,
      processedBy: processedBy,
      timestamp: getFormattedTimestamp()
    });
    
    console.log(`✅ Updated order ${orderId} - Status: ${displayStatus}, By: ${processedBy}`);
    
    return {
      success: true,
      message: `Order status updated to: ${displayStatus}`,
      processedBy: processedBy,
      processedAt: getFormattedTimestamp()
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
 * ✅ NEW: Update order comments
 * @param {string} orderId - Order ID
 * @param {string} platform - Platform (amazon/target)
 * @param {string} newComments - New comments text
 * @returns {Object} Update result
 */
function updateOrderComments(orderId, platform, newComments) {
  try {
    console.log(`Updating comments: ID=${orderId}, Platform=${platform}`);
    
    const sheetMap = {
      'amazon': 'AmazonOrders',
      'target': 'TargetOrders'
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
    
    // Find or create Comments column
    let commentsColumnIndex = findColumnIndex(headers, 'Comments');
    if (commentsColumnIndex === -1) {
      commentsColumnIndex = headers.length;
      sheet.getRange(1, commentsColumnIndex + 1).setValue('Comments');
    }
    
    // Update comments
    sheet.getRange(foundRow + 1, commentsColumnIndex + 1).setValue(newComments);
    
    // Log success
    logAction('ORDER_COMMENTS_UPDATED', {
      orderId: orderId,
      platform: platform
    });
    
    console.log(`✅ Updated comments for row ${foundRow + 1}`);
    
    return {
      success: true,
      message: 'Comments updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating comments:', error);
    
    logAction('ORDER_COMMENTS_UPDATE_FAILED', {
      orderId: orderId,
      platform: platform,
      error: error.message
    }, 'ERROR');
    
    return { success: false, message: error.message };
  }
}

/**
 * ✅ NEW: Create manual order (Amazon or Target)
 * @param {Object} orderData - Order data from form
 * @returns {Object} Creation result
 */
function createManualOrder(orderData) {
  try {
    console.log('📝 Creating manual order:', orderData.platform);
    
    const platform = orderData.platform; // 'amazon' or 'target'
    const sheetMap = {
      'amazon': 'AmazonOrders',
      'target': 'TargetOrders'
    };
    
    const sheetName = sheetMap[platform];
    if (!sheetName) {
      throw new Error(`Invalid platform: ${platform}`);
    }
    
    // Prepare form data object
    const formData = {
      form_type: platform === 'amazon' ? 'amazon-order' : 'target-order',
      Student_Organization: orderData.organization,
      Event_Name: orderData.eventName,
      Event_Date: orderData.eventDate,
      Pickup_Person_Name: orderData.pickupPerson,
      Pickup_Person_Email: orderData.pickupEmail,
      Pickup_Person_Phone: orderData.pickupPhone,
      Submission_Email_Address: orderData.pickupEmail,
      'Form Submitter Notes': orderData.notes || '',
      Order_Status: 'New Order'
    };
    
    // Platform-specific fields
    if (platform === 'amazon') {
      formData.Wishlist_Link = orderData.wishlistLink;
      formData.Total_Order = orderData.totalOrder || 0;
    } else {
      formData.Cart_Total = orderData.cartTotal || 0;
      // Add item URLs and quantities
      if (orderData.items && orderData.items.length > 0) {
        orderData.items.forEach((item, index) => {
          const itemNum = index + 1;
          const itemNumWords = ['First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth', 'Seventh', 'Eighth', 'Ninth', 'Tenth'];
          const itemWord = itemNumWords[index] || `Item_${itemNum}`;
          formData[`${itemWord}_Item_Url`] = item.url;
          formData[`${itemWord}_Item_Quantity`] = item.quantity || '1';
        });
      }
    }
    
    // Write to sheet
    const writeResult = writeToSheet(formData);
    
    if (!writeResult.success) {
      throw new Error('Failed to write order: ' + writeResult.error);
    }
    
    // Assign IDs
    const sheet = getCachedSheet(writeResult.sheetName);
    const idResult = assignIDsToRow(sheet, writeResult.row);
    
    // Generate confirmation number
    const confirmationNumber = generateConfirmationNumber(orderData.organization);
    addConfirmationNumberToSheet(sheet, writeResult.row, confirmationNumber);
    
    // Log success
    logAction('MANUAL_ORDER_CREATED', {
      platform: platform,
      organization: orderData.organization,
      confirmationNumber: confirmationNumber
    });
    
    console.log(`✅ Manual order created: ${confirmationNumber}`);
    
    return {
      success: true,
      message: 'Order created successfully',
      confirmationNumber: confirmationNumber,
      orderId: idResult.assignedIds?.Historical_Submission_Id
    };
    
  } catch (error) {
    console.error('❌ Manual order creation failed:', error);
    
    logAction('MANUAL_ORDER_CREATION_FAILED', {
      error: error.message
    }, 'ERROR');
    
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ✅ NEW: Submit manual documents (with file upload support)
 * @param {Object} docData - Document submission data
 * @returns {Object} Submission result
 */
function submitManualDocuments(docData) {
  try {
    console.log('📄 Submitting manual documents');
    
    // Prepare form data
    const formData = {
      form_type: 'document-submission',
      First_Name: docData.firstName,
      Last_Name: docData.lastName,
      Email_Address: docData.email,
      Student_Organization: docData.organization,
      Event_Name: docData.eventName,
      Event_Date: docData.eventDate,
      'Form Submitter Notes': docData.notes || ''
    };
    
    // ✅ Add file data if present
    if (docData.fileData && Object.keys(docData.fileData).length > 0) {
      formData.fileData = docData.fileData;
      console.log(`📎 Processing ${Object.keys(docData.fileData).length} files`);
      
      // Process files BEFORE writing to sheet
      processFilesAndAddLinksToFormData(formData);
    }
    
    // Write to sheet
    const writeResult = writeToSheet(formData);
    
    if (!writeResult.success) {
      throw new Error('Failed to write document submission: ' + writeResult.error);
    }
    
    // Assign IDs
    const sheet = getCachedSheet(writeResult.sheetName);
    const idResult = assignIDsToRow(sheet, writeResult.row);
    
    // Generate confirmation number
    const confirmationNumber = generateConfirmationNumber(docData.organization);
    addConfirmationNumberToSheet(sheet, writeResult.row, confirmationNumber);
    
    // Log success
    logAction('MANUAL_DOCUMENT_SUBMITTED', {
      organization: docData.organization,
      confirmationNumber: confirmationNumber,
      filesUploaded: docData.fileData ? Object.keys(docData.fileData).length : 0
    });
    
    console.log(`✅ Manual document submission created: ${confirmationNumber}`);
    
    return {
      success: true,
      message: 'Document submission created successfully',
      confirmationNumber: confirmationNumber,
      docId: idResult.assignedIds?.Historical_Submission_Id,
      filesUploaded: docData.fileData ? Object.keys(docData.fileData).length : 0
    };
    
  } catch (error) {
    console.error('❌ Manual document submission failed:', error);
    
    logAction('MANUAL_DOCUMENT_SUBMISSION_FAILED', {
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
