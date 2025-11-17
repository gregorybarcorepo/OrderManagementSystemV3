/**
 * ============================================
 * SHARED UTILITIES - V3 with Logging
 * Common functions used across the application
 * ============================================
 */

// ========================================
// LOGGING SYSTEM (NEW!)
// ========================================

/**
 * Simple logging system - writes to a "Logs" sheet
 * @param {string} action - Action name (e.g., 'ORDER_CREATED')
 * @param {Object} details - Details object
 * @param {string} level - Log level: 'INFO', 'WARNING', 'ERROR'
 */
function logAction(action, details, level = 'INFO') {
  try {
    const ss = getCachedSpreadsheet();
    let logSheet = ss.getSheetByName('Logs');
    
    // Create Logs sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet('Logs');
      logSheet.getRange(1, 1, 1, 5).setValues([
        ['Timestamp', 'Level', 'Action', 'Details', 'User']
      ]);
      logSheet.setFrozenRows(1);
      
      // Format header
      logSheet.getRange(1, 1, 1, 5)
        .setBackground('#4a5568')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      
      console.log('✅ Logs sheet created');
    }
    
    // Add log entry
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();
    
    logSheet.appendRow([
      Utilities.formatDate(timestamp, 'America/New_York', 'yyyy-MM-dd HH:mm:ss'),
      level,
      action,
      JSON.stringify(details),
      user
    ]);
    
    // Keep only last 1000 logs (prevent sheet bloat)
    if (logSheet.getLastRow() > 1001) {
      logSheet.deleteRows(2, 100); // Delete oldest 100
    }
    
  } catch (error) {
    console.error('Logging failed:', error);
    // Don't throw - logging failures shouldn't break the app
  }
}

// ========================================
// SPREADSHEET ACCESS
// ========================================

let _cachedSpreadsheet = null;

function getCachedSpreadsheet() {
  if (!_cachedSpreadsheet) {
    const key = PropertiesService.getScriptProperties().getProperty('key');
    if (!key) {
      throw new Error('Spreadsheet key not found. Run masterSetup() first.');
    }
    _cachedSpreadsheet = SpreadsheetApp.openById(key);
  }
  return _cachedSpreadsheet;
}

function getCachedSheet(sheetName) {
  const spreadsheet = getCachedSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    console.log(`➕ Creating new sheet: ${sheetName}`);
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  return sheet;
}

// ========================================
// COLUMN UTILITIES
// ========================================

/**
 * Find column index by header name (case-insensitive, handles variations)
 * @param {Array} headers - Array of header strings
 * @param {string} targetHeader - Header name to find
 * @returns {number} Column index (0-based) or -1 if not found
 */
function findColumnIndex(headers, targetHeader) {
  if (!headers || !targetHeader) return -1;
  
  const target = targetHeader.toLowerCase().replace(/[_\s-]/g, '');
  
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().replace(/[_\s-]/g, '');
    if (header === target) {
      return i;
    }
  }
  
  return -1;
}

/**
 * Get all column indices that match a pattern
 * @param {Array} headers - Array of header strings
 * @param {string} pattern - Pattern to match (can include wildcards)
 * @returns {Array} Array of matching column indices
 */
function findColumnsByPattern(headers, pattern) {
  const matches = [];
  const regexPattern = new RegExp(pattern.replace('*', '.*'), 'i');
  
  for (let i = 0; i < headers.length; i++) {
    if (regexPattern.test(headers[i])) {
      matches.push(i);
    }
  }
  
  return matches;
}

// ========================================
// DATE/TIME UTILITIES
// ========================================

/**
 * Get current semester (Spring YYYY or Fall YYYY)
 * @returns {string} Semester string
 */
function getSemester() {
  const now = new Date();
  const month = now.getMonth() + 1; // January = 1
  const year = now.getFullYear();
  
  // January-May = Spring, June-December = Fall
  const season = (month <= 5) ? 'Spring' : 'Fall';
  return `${season} ${year}`;
}

/**
 * Get formatted date string (MM-DD-YYYY)
 * @param {Date} date - Optional date (defaults to now)
 * @returns {string} Formatted date string
 */
function getDateString(date) {
  const now = date || new Date();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const year = now.getFullYear();
  return `${month}-${day}-${year}`;
}

/**
 * Get formatted timestamp (EST timezone)
 * @param {Date} date - Optional date (defaults to now)
 * @returns {string} Formatted timestamp
 */
function getFormattedTimestamp(date) {
  const timestamp = date || new Date();
  return Utilities.formatDate(
    timestamp, 
    'America/New_York', 
    'MM/dd/yyyy hh:mm:ss a'
  );
}

/**
 * Check if a value matches HTML date format (YYYY-MM-DD)
 * @param {*} value - Value to check
 * @returns {boolean}
 */
function isHTMLDateFormat(value) {
  if (typeof value !== 'string') return false;
  const htmlDatePattern = /^\d{4}-\d{2}-\d{2}$/;
  return htmlDatePattern.test(value);
}

/**
 * Convert HTML date (YYYY-MM-DD) to readable format (MM/DD/YYYY)
 * @param {string} htmlDate - HTML date string
 * @returns {string} Formatted date
 */
function formatHTMLDate(htmlDate) {
  if (!isHTMLDateFormat(htmlDate)) return htmlDate;
  
  try {
    const dateParts = htmlDate.split('-');
    const year = dateParts[0];
    const month = dateParts[1];
    const day = dateParts[2];
    return `${month}/${day}/${year}`;
  } catch (error) {
    console.error('❌ Error formatting HTML date:', error);
    return htmlDate;
  }
}

// ========================================
// FILE UTILITIES
// ========================================

/**
 * Sanitize text for use in filenames
 * @param {string} text - Text to sanitize
 * @param {number} maxLength - Maximum length (default 20)
 * @returns {string} Sanitized text
 */
function sanitizeFileName(text, maxLength = 20) {
  if (!text) return 'Unknown';
  
  return text
    .replace(/[^a-zA-Z0-9\s]/g, '') // Remove special chars
    .replace(/\s+/g, '_')           // Replace spaces
    .substring(0, maxLength);       // Limit length
}

/**
 * Get file extension from filename
 * @param {string} fileName - Full filename
 * @returns {string} Extension (lowercase) or 'pdf' as fallback
 */
function getFileExtension(fileName) {
  if (!fileName) return 'pdf';
  
  const parts = fileName.split('.');
  return parts.length > 1 ? parts.pop().toLowerCase() : 'pdf';
}

/**
 * Get MIME type from file extension
 * @param {string} extension - File extension
 * @returns {string} MIME type
 */
function getMimeType(extension) {
  const mimeTypes = {
    'pdf': 'application/pdf',
    'doc': 'application/msword',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xls': 'application/vnd.ms-excel',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'png': 'image/png',
    'gif': 'image/gif'
  };
  
  return mimeTypes[extension.toLowerCase()] || 'application/octet-stream';
}

// ========================================
// STATUS UTILITIES
// ========================================

/**
 * Normalize status values for backend storage (display format)
 * @param {*} status - Raw status value
 * @returns {string} Normalized status for Google Sheets
 */
function normalizeStatus(status) {
  if (status === null || status === undefined) return 'New Order';
  
  const s = String(status).trim().toLowerCase();
  
  // Map all variations to display format
  const statusMap = {
    'new-order': 'New Order',
    'new order': 'New Order',
    'pending': 'New Order',
    'new': 'New Order',
    
    'ordered': 'Ordered',
    'processing': 'Ordered',
    'in-progress': 'Ordered',
    
    'delivered-to-mailroom': 'Delivered to Mailroom',
    'delivered to mailroom': 'Delivered to Mailroom',
    'delivered': 'Delivered to Mailroom',
    'received': 'Delivered to Mailroom',
    
    'completed': 'Completed',
    'complete': 'Completed',
    'done': 'Completed',
    
    'awaiting-club-response': 'Awaiting Club Response',
    'awaiting club response': 'Awaiting Club Response',
    'on-hold': 'Awaiting Club Response',
    'on hold': 'Awaiting Club Response',
    
    'cancelled': 'Cancelled',
    'canceled': 'Cancelled'
  };
  
  return statusMap[s] || 'New Order';
}

// ========================================
// NUMBER UTILITIES
// ========================================

/**
 * Parse money safely (supports $, commas, blanks)
 * @param {*} value - Value to parse
 * @returns {number} Parsed number or 0
 */
function parseMoneyValue(value) {
  if (value === null || value === undefined || value === '') return 0;
  
  const cleaned = String(value).replace(/[^\d.\-]/g, '');
  const parsed = parseFloat(cleaned);
  
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Extract maximum numeric ID from a column of data
 * @param {Array} columnData - 2D array from sheet range
 * @returns {number} Maximum ID found
 */
function getMaxIdFromColumn(columnData) {
  let maxId = 0;
  
  columnData.forEach(row => {
    const value = row[0];
    if (value && value !== '') {
      const numericValue = parseInt(value);
      if (!isNaN(numericValue) && numericValue > maxId) {
        maxId = numericValue;
      }
    }
  });
  
  return maxId;
}

// ========================================
// VALIDATION UTILITIES
// ========================================

/**
 * Check if email is valid
 * @param {string} email - Email to validate
 * @returns {boolean}
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

/**
 * Check if URL is valid
 * @param {string} url - URL to validate
 * @returns {boolean}
 */
function isValidURL(url) {
  if (!url) return false;
  try {
    new URL(url);
    return true;
  } catch (error) {
    return false;
  }
}
