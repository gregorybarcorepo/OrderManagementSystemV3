/**
 * ============================================
 * CONFIGURATION
 * Central configuration for the entire system
 * ============================================
 */

/**
 * System Configuration Object
 */
const SYSTEM_CONFIG = {
  // Google Drive folder for file uploads
  DRIVE_FOLDER_ID: "1go_LBeXdU2XbGhkFk24wsqURKnA-es4_",
  
  // Email configuration
  EMAIL: {
    FROM_NAME: 'Brooklyn College Central Depository',
    STAFF_EMAIL: 'greg@brooklyn.cuny.club',
    TRACKING_ENABLED: false // Set to true when tracking pixel deployed
  },
  
  // Sheet names mapping
  SHEETS: {
    TARGET_ORDERS: 'TargetOrders',
    AMAZON_ORDERS: 'AmazonOrders',
    DOCUMENT_SUBMISSIONS: 'DocumentSubmissions',
    CAMPUS_FEEDBACK: 'CampusClubSpaceFeedback',
    NAMES: 'Names'
  },
  
  // Form type mapping
  FORM_TYPES: {
    "target-order": "TargetOrders",
    "amazon-order": "AmazonOrders",
    "document-submission": "DocumentSubmissions",
    "campus-club-space-feedback": "CampusClubSpaceFeedback"
  },
  
  // File upload settings
  FILES: {
    MAX_SIZE_MB: 10,
    ALLOWED_MIME_TYPES: [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'image/jpeg',
      'image/png',
      'image/gif'
    ],
    ALLOWED_EXTENSIONS: ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'jpg', 'jpeg', 'png', 'gif']
  },
  
  // ID column names
  ID_COLUMNS: {
    HISTORICAL: 'Historical_Submission_Id',
    SEMESTER: 'Semester_Id',
    SUBMISSION: 'Submission_Id',
    CONFIRMATION: 'Confirmation_Number'
  },
  
  // Timezone
  TIMEZONE: 'America/New_York',
  
  // STATUSES
  STATUSES: {
  NEW_ORDER: 'New Order',
  ORDERED: 'Ordered',
  DELIVERED: 'Delivered to Mailroom',
  AWAITING_RESPONSE: 'Awaiting Club Response',
  CANCELLED: 'Cancelled',
  UNASSIGNED: 'Unassigned'
}
};

/**
 * Get configuration value
 * @param {string} path - Dot-notation path (e.g., 'EMAIL.STAFF_EMAIL')
 * @returns {*} Configuration value
 */
function getConfig(path) {
  const keys = path.split('.');
  let value = SYSTEM_CONFIG;
  
  for (const key of keys) {
    if (value && typeof value === 'object' && key in value) {
      value = value[key];
    } else {
      return null;
    }
  }
  
  return value;
}

/**
 * Update configuration value (use with caution in production)
 * @param {string} path - Dot-notation path
 * @param {*} newValue - New value
 * @returns {boolean} Success status
 */
function setConfig(path, newValue) {
  const keys = path.split('.');
  const lastKey = keys.pop();
  let obj = SYSTEM_CONFIG;
  
  for (const key of keys) {
    if (!(key in obj)) {
      obj[key] = {};
    }
    obj = obj[key];
  }
  
  obj[lastKey] = newValue;
  return true;
}

/**
 * Get all configuration (for debugging)
 * @returns {Object} Full configuration
 */
function getAllConfig() {
  return SYSTEM_CONFIG;
}
