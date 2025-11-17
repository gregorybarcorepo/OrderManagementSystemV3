/**
 * ============================================
 * EMAIL TEMPLATES & SENDING
 * Unified email system with tracking support
 * ============================================
 */

// Email configuration
const EMAIL_CONFIG = {
  FROM_NAME: 'Brooklyn College Central Depository',
  STAFF_EMAIL: 'greg@brooklyn.cuny.club',
  TRACKING_ENABLED: false // Set to true when tracking pixel is deployed
};

// ========================================
// MAIN EMAIL FUNCTIONS
// ========================================

/**
 * Send confirmation email to form submitter
 * @param {Object} formData - Form submission data
 * @param {string} formType - Form type identifier
 * @param {string} confirmationNumber - Confirmation number
 * @param {string} trackingPixel - Optional tracking pixel HTML
 * @returns {Object} Email sending result
 */
function sendConfirmationEmail(formData, formType, confirmationNumber, trackingPixel = '') {
  try {
    console.log('📧 Preparing confirmation email...');
    
    // Get submitter email
    const submitterEmail = formData.Submission_Email_Address || 
                          formData.Pickup_Person_Email ||
                          formData.email ||
                          formData.Email_Address;
    
    if (!submitterEmail || !isValidEmail(submitterEmail)) {
      console.log('❌ No valid email address found');
      return { 
        sent: false, 
        error: 'No valid email address provided' 
      };
    }
    
    // Generate email content
    const timestamp = getFormattedTimestamp();
    const emailContent = generateEmailContent(formData, formType, confirmationNumber, timestamp, trackingPixel);
    
    // Send email
    MailApp.sendEmail({
      to: submitterEmail,
      subject: emailContent.subject,
      htmlBody: emailContent.htmlBody,
      name: EMAIL_CONFIG.FROM_NAME
    });
    
    console.log('✅ Confirmation email sent to:', submitterEmail);
    
    // Send staff notification
    sendStaffNotification(formData, formType, confirmationNumber);
    
    return {
      sent: true,
      sentTo: submitterEmail,
      sentAt: timestamp
    };
    
  } catch (error) {
    console.error('❌ Error sending email:', error);
    return {
      sent: false,
      error: error.toString()
    };
  }
}

/**
 * Send staff notification email
 * @param {Object} formData - Form submission data
 * @param {string} formType - Form type identifier
 * @param {string} confirmationNumber - Confirmation number
 */
function sendStaffNotification(formData, formType, confirmationNumber) {
  try {
    const subject = `🔔 New ${formType} Submission - ${confirmationNumber}`;
    
    const body = `
New form submission received:

Confirmation Number: ${confirmationNumber}
Form Type: ${formType}
Organization: ${formData.Student_Organization || formData.Organization_Name || 'N/A'}
Event: ${formData.Event_Name || 'N/A'}
Event Date: ${formData.Event_Date || 'N/A'}
Submitter Email: ${formData.Submission_Email_Address || formData.Pickup_Person_Email || 'N/A'}

View in Order Management System:
${ScriptApp.getService().getUrl()}

Time: ${getFormattedTimestamp()}
    `.trim();
    
    MailApp.sendEmail(EMAIL_CONFIG.STAFF_EMAIL, subject, body);
    console.log('✅ Staff notification sent');
    
  } catch (error) {
    console.error('❌ Staff notification failed:', error);
  }
}

// ========================================
// EMAIL CONTENT GENERATION
// ========================================

/**
 * Generate email content based on form type
 * @param {Object} formData - Form data
 * @param {string} formType - Form type
 * @param {string} confirmationNumber - Confirmation number
 * @param {string} timestamp - Formatted timestamp
 * @param {string} trackingPixel - Optional tracking pixel
 * @returns {Object} Email subject and HTML body
 */
function generateEmailContent(formData, formType, confirmationNumber, timestamp, trackingPixel = '') {
  let subject, htmlBody;
  
  switch (formType) {
    case 'amazon-order':
      subject = `✅ Amazon Order Confirmation - ${confirmationNumber}`;
      htmlBody = generateAmazonEmailHTML(formData, confirmationNumber, timestamp, trackingPixel);
      break;
      
    case 'target-order':
      subject = `✅ Target Order Confirmation - ${confirmationNumber}`;
      htmlBody = generateTargetEmailHTML(formData, confirmationNumber, timestamp, trackingPixel);
      break;
      
    case 'document-submission':
      subject = `✅ Document Submission Confirmation - ${confirmationNumber}`;
      htmlBody = generateDocumentEmailHTML(formData, confirmationNumber, timestamp, trackingPixel);
      break;
      
    case 'campus-club-space-feedback':
      subject = `✅ Feedback Confirmation - ${confirmationNumber}`;
      htmlBody = generateFeedbackEmailHTML(formData, confirmationNumber, timestamp, trackingPixel);
      break;
      
    default:
      subject = `✅ Form Submission Confirmation - ${confirmationNumber}`;
      htmlBody = generateGenericEmailHTML(formData, confirmationNumber, timestamp, trackingPixel);
  }
  
  return { subject, htmlBody };
}

// ========================================
// EMAIL TEMPLATE STYLES (SHARED)
// ========================================

/**
 * Get shared CSS styles for email templates
 * @returns {string} CSS styles
 */
function getEmailStyles() {
  return `
    body {
      font-family: Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 600px;
      margin: 0 auto;
      background-color: #f5f5f5;
    }
    .email-container {
      background-color: #ffffff;
      margin: 20px auto;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .header {
      background-color: #7d2049;
      color: white;
      padding: 25px;
      text-align: center;
    }
    .header.amazon { background-color: #FF9900; }
    .header.target { background-color: #cc0000; }
    .header.document { background-color: #2C5AA0; }
    .header.feedback { background-color: #48A9A6; }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .header p {
      margin: 5px 0 0 0;
      font-size: 14px;
      opacity: 0.9;
    }
    .content {
      padding: 30px;
    }
    .confirmation-box {
      background-color: #d4edda;
      border: 2px solid #28a745;
      padding: 20px;
      text-align: center;
      border-radius: 8px;
      margin: 20px 0;
    }
    .confirmation-label {
      font-size: 12px;
      color: #155724;
      text-transform: uppercase;
      letter-spacing: 1px;
      margin-bottom: 5px;
      font-weight: 600;
    }
    .confirmation-number {
      font-size: 20px;
      font-weight: bold;
      color: #155724;
      letter-spacing: 1px;
      font-family: 'Courier New', monospace;
    }
    .confirmation-note {
      font-size: 11px;
      color: #155724;
      margin-top: 5px;
    }
    .info-box {
      background-color: #f8f9fa;
      padding: 20px;
      border-left: 4px solid #7d2049;
      margin: 20px 0;
      border-radius: 4px;
    }
    .info-box h3 {
      margin-top: 0;
      color: #7d2049;
      font-size: 16px;
      margin-bottom: 15px;
    }
    .info-row {
      padding: 8px 0;
      border-bottom: 1px solid #e9ecef;
      display: flex;
      flex-wrap: wrap;
    }
    .info-row:last-child {
      border-bottom: none;
    }
    .info-label {
      font-weight: 600;
      color: #495057;
      min-width: 120px;
      margin-right: 10px;
    }
    .info-value {
      color: #212529;
      flex: 1;
      word-break: break-word;
    }
    .next-steps {
      background-color: #fff3cd;
      border-left: 4px solid #ffc107;
      padding: 20px;
      margin: 20px 0;
      border-radius: 4px;
    }
    .next-steps h3 {
      margin-top: 0;
      color: #856404;
      font-size: 16px;
    }
    .button {
      background-color: #7d2049;
      color: white !important;
      padding: 12px 30px;
      text-decoration: none;
      border-radius: 5px;
      display: inline-block;
      margin: 15px 0;
      font-weight: 600;
    }
    .footer {
      background-color: #f8f9fa;
      padding: 25px;
      text-align: center;
      color: #6c757d;
      font-size: 13px;
      border-top: 1px solid #e9ecef;
    }
    .footer p {
      margin: 5px 0;
    }
    .footer-brand {
      font-weight: bold;
      color: #7d2049;
      font-size: 14px;
    }
    a {
      color: #7d2049;
      word-break: break-all;
    }
  `;
}

// ========================================
// AMAZON EMAIL TEMPLATE
// ========================================

/**
 * Generate Amazon order confirmation email
 */
function generateAmazonEmailHTML(formData, confirmationNumber, timestamp, trackingPixel) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>${getEmailStyles()}</style>
</head>
<body>
  ${trackingPixel}
  <div class="email-container">
    <div class="header amazon">
      <h1>🛒 Amazon Order Request Received</h1>
      <p>Brooklyn College Central Depository</p>
    </div>
    
    <div class="content">
      <p>Dear <strong>${formData.Student_Organization || 'Student Organization'}</strong>,</p>
      
      <p>Thank you for submitting your Amazon order request. Your submission has been received and will be processed by our team.</p>
      
      <div class="confirmation-box">
        <div class="confirmation-label">Confirmation Number</div>
        <div class="confirmation-number">${confirmationNumber}</div>
        <div class="confirmation-note">Please save this number for your records</div>
      </div>
      
      <div class="info-box">
        <h3>📋 Submission Details</h3>
        <div class="info-row">
          <span class="info-label">Submitted:</span>
          <span class="info-value">${timestamp}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Event Name:</span>
          <span class="info-value">${formData.Event_Name || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Event Date:</span>
          <span class="info-value">${formData.Event_Date || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Organization:</span>
          <span class="info-value">${formData.Student_Organization || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Wishlist Link:</span>
          <span class="info-value"><a href="${formData.Wishlist_Link || '#'}">${formData.Wishlist_Link || 'Not provided'}</a></span>
        </div>
      </div>
      
      <div class="info-box">
        <h3>👤 Pickup Contact Information</h3>
        <div class="info-row">
          <span class="info-label">Name:</span>
          <span class="info-value">${formData.Pickup_Person_Name || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Email:</span>
          <span class="info-value">${formData.Pickup_Person_Email || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Phone:</span>
          <span class="info-value">${formData.Pickup_Person_Phone || 'Not specified'}</span>
        </div>
      </div>
      
      <div class="next-steps">
        <h3>📞 Next Steps</h3>
        <p>Our team will review your Amazon order request and contact you within <strong>2-3 business days</strong>.</p>
      </div>
      
      <div style="text-align: center;">
        <a href="mailto:${EMAIL_CONFIG.STAFF_EMAIL}" class="button">Contact Us</a>
      </div>
    </div>
    
    <div class="footer">
      <p class="footer-brand">Brooklyn College Central Depository</p>
      <p>Room 314, Student Center<br>2705 Campus Rd, Brooklyn, NY 11210</p>
      <p><a href="https://brooklyn.cuny.club/cd">Visit our website</a></p>
    </div>
  </div>
</body>
</html>
  `;
}

// ========================================
// TARGET EMAIL TEMPLATE
// ========================================

/**
 * Generate Target order confirmation email
 */
function generateTargetEmailHTML(formData, confirmationNumber, timestamp, trackingPixel) {
  // Count items
  let itemCount = 0;
  const itemFields = ['First_Item_Url', 'Second_Item_Url', 'Third_Item_Url', 'Fourth_Item_Url', 'Fifth_Item_Url',
                     'Sixth_Item_Url', 'Seventh_Item_Url', 'Eighth_Item_Url', 'Ninth_Item_Url', 'Tenth_Item_Url'];
  itemFields.forEach(field => {
    if (formData[field] && formData[field].trim() !== '') itemCount++;
  });
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>${getEmailStyles()}</style>
</head>
<body>
  ${trackingPixel}
  <div class="email-container">
    <div class="header target">
      <h1>🎯 Target Order Request Received</h1>
      <p>Brooklyn College Central Depository</p>
    </div>
    
    <div class="content">
      <p>Dear <strong>${formData.Student_Organization || 'Student Organization'}</strong>,</p>
      
      <p>Thank you for submitting your Target order request. Your submission has been received and will be processed by our team.</p>
      
      <div class="confirmation-box">
        <div class="confirmation-label">Confirmation Number</div>
        <div class="confirmation-number">${confirmationNumber}</div>
        <div class="confirmation-note">Please save this number for your records</div>
      </div>
      
      <div class="info-box">
        <h3>📋 Submission Details</h3>
        <div class="info-row">
          <span class="info-label">Submitted:</span>
          <span class="info-value">${timestamp}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Event Name:</span>
          <span class="info-value">${formData.Event_Name || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Event Date:</span>
          <span class="info-value">${formData.Event_Date || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Items Requested:</span>
          <span class="info-value">${itemCount} items</span>
        </div>
        <div class="info-row">
          <span class="info-label">Estimated Total:</span>
          <span class="info-value">$${formData.Cart_Total || '0.00'}</span>
        </div>
      </div>
      
      <div class="info-box">
        <h3>👤 Pickup Contact Information</h3>
        <div class="info-row">
          <span class="info-label">Name:</span>
          <span class="info-value">${formData.Pickup_Person_Name || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Email:</span>
          <span class="info-value">${formData.Pickup_Person_Email || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Phone:</span>
          <span class="info-value">${formData.Pickup_Person_Phone || 'Not specified'}</span>
        </div>
      </div>
      
      <div class="next-steps">
        <h3>📞 Next Steps</h3>
        <p>Our team will review your Target order request and contact you within <strong>2-3 business days</strong>. We will process your order and notify you when items are ready for pickup.</p>
      </div>
      
      <div style="text-align: center;">
        <a href="mailto:${EMAIL_CONFIG.STAFF_EMAIL}" class="button">Contact Us</a>
      </div>
    </div>
    
    <div class="footer">
      <p class="footer-brand">Brooklyn College Central Depository</p>
      <p>Room 314, Student Center<br>2705 Campus Rd, Brooklyn, NY 11210</p>
      <p><a href="https://brooklyn.cuny.club/cd">Visit our website</a></p>
    </div>
  </div>
</body>
</html>
  `;
}

// ========================================
// DOCUMENT SUBMISSION EMAIL TEMPLATE
// ========================================

/**
 * Generate document submission confirmation email
 */
function generateDocumentEmailHTML(formData, confirmationNumber, timestamp, trackingPixel) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>${getEmailStyles()}</style>
</head>
<body>
  ${trackingPixel}
  <div class="email-container">
    <div class="header document">
      <h1>📄 Document Submission Received</h1>
      <p>Brooklyn College Central Depository</p>
    </div>
    
    <div class="content">
      <p>Dear <strong>${formData.Student_Organization || formData.Organization_Name || 'Student Organization'}</strong>,</p>
      
      <p>Your document submission has been successfully received.</p>
      
      <div class="confirmation-box">
        <div class="confirmation-label">Confirmation Number</div>
        <div class="confirmation-number">${confirmationNumber}</div>
        <div class="confirmation-note">Please save this number for your records</div>
      </div>
      
      <div class="info-box">
        <h3>📋 Submission Details</h3>
        <div class="info-row">
          <span class="info-label">Submitted:</span>
          <span class="info-value">${timestamp}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Event Name:</span>
          <span class="info-value">${formData.Event_Name || 'Not specified'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Event Date:</span>
          <span class="info-value">${formData.Event_Date || 'Not specified'}</span>
        </div>
      </div>
      
      <div class="next-steps">
        <h3>📞 Next Steps</h3>
        <p>We will review your submission within <strong>3-5 business days</strong> and contact you if we need any additional information.</p>
      </div>
      
      <div style="text-align: center;">
        <a href="mailto:${EMAIL_CONFIG.STAFF_EMAIL}" class="button">Contact Us</a>
      </div>
    </div>
    
    <div class="footer">
      <p class="footer-brand">Brooklyn College Central Depository</p>
      <p>Room 314, Student Center<br>2705 Campus Rd, Brooklyn, NY 11210</p>
    </div>
  </div>
</body>
</html>
  `;
}

// ========================================
// FEEDBACK EMAIL TEMPLATE
// ========================================

/**
 * Generate feedback confirmation email
 */
function generateFeedbackEmailHTML(formData, confirmationNumber, timestamp, trackingPixel) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>${getEmailStyles()}</style>
</head>
<body>
  ${trackingPixel}
  <div class="email-container">
    <div class="header feedback">
      <h1>💬 Feedback Received</h1>
      <p>Brooklyn College Central Depository</p>
    </div>
    
    <div class="content">
      <p>Thank you for your valuable feedback!</p>
      
      <div class="confirmation-box">
        <div class="confirmation-label">Confirmation Number</div>
        <div class="confirmation-number">${confirmationNumber}</div>
      </div>
      
      <p>Your input helps us improve our services for the Brooklyn College community.</p>
      
      <div style="text-align: center; margin-top: 30px;">
        <a href="mailto:${EMAIL_CONFIG.STAFF_EMAIL}" class="button">Contact Us</a>
      </div>
    </div>
    
    <div class="footer">
      <p class="footer-brand">Brooklyn College Central Depository</p>
      <p>Room 314, Student Center<br>2705 Campus Rd, Brooklyn, NY 11210</p>
    </div>
  </div>
</body>
</html>
  `;
}

// ========================================
// GENERIC EMAIL TEMPLATE
// ========================================

/**
 * Generate generic confirmation email
 */
function generateGenericEmailHTML(formData, confirmationNumber, timestamp, trackingPixel) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>${getEmailStyles()}</style>
</head>
<body>
  ${trackingPixel}
  <div class="email-container">
    <div class="header">
      <h1>✅ Form Submission Received</h1>
      <p>Brooklyn College Central Depository</p>
    </div>
    
    <div class="content">
    <p>Your form submission has been successfully received.</p>
      
      <div class="confirmation-box">
        <div class="confirmation-label">Confirmation Number</div>
        <div class="confirmation-number">${confirmationNumber}</div>
        <div class="confirmation-note">Please save this number for your records</div>
      </div>
      
      <div class="info-box">
        <h3>📋 Submission Details</h3>
        <div class="info-row">
          <span class="info-label">Submitted:</span>
          <span class="info-value">${timestamp}</span>
        </div>
      </div>
      
      <div class="next-steps">
        <h3>📞 Next Steps</h3>
        <p>Our team will review your submission and contact you if additional information is needed.</p>
      </div>
      
      <div style="text-align: center;">
        <a href="mailto:${EMAIL_CONFIG.STAFF_EMAIL}" class="button">Contact Us</a>
      </div>
    </div>
    
    <div class="footer">
      <p class="footer-brand">Brooklyn College Central Depository</p>
      <p>Room 314, Student Center<br>2705 Campus Rd, Brooklyn, NY 11210</p>
    </div>
  </div>
</body>
</html>
  `;
}
