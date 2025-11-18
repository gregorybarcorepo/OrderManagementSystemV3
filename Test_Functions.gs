/**
 * ============================================
 * TEST FUNCTIONS
 * Temporary functions for testing before deployment
 * You can delete this entire file after deployment
 * ============================================
 */

/**
 * Test confirmation number handling and email sending
 * Run this BEFORE deploying to verify everything works
 */
function testConfirmationAndEmail() {
  console.log('🧪 Testing submission flow...');
  console.log('==========================================');
  
  const testData = {
    form_type: "target-order",
    Student_Organization: "Test Club",
    Event_Name: "Test Event",
    Event_Date: "2025-12-31",
    Submission_Email_Address: "greg@brooklyn.cuny.club", // ← CHANGE THIS
    Pickup_Person_Name: "Test Person",
    Pickup_Person_Email: "greg@brooklyn.cuny.club",      // ← CHANGE THIS
    Pickup_Person_Phone: "555-1234",
    Confirmation_Number: "CD-T-01182025-TEST-TESTCLUB", // Frontend sends this
    First_Item_Url: "https://www.target.com/test",
    First_Item_Quantity: "1",
    Cart_Total: "25.00"
  };
  
  const mockEvent = {
    postData: { contents: JSON.stringify(testData) }
  };
  
  try {
    console.log('1️⃣ Submitting form...');
    const response = doPost(mockEvent);
    const result = JSON.parse(response.getContent());
    
    console.log('\n2️⃣ Testing confirmation number...');
    console.log('   Frontend sent:', testData.Confirmation_Number);
    console.log('   Backend returned:', result.confirmationNumber);
    const confMatch = result.confirmationNumber === testData.Confirmation_Number;
    console.log('   Match:', confMatch ? '✅ PASS' : '❌ FAIL');
    
    console.log('\n3️⃣ Testing email...');
    console.log('   Email sent:', result.emailSent ? '✅ YES' : '❌ NO');
    console.log('   Sent to:', testData.Submission_Email_Address);
    
    console.log('\n4️⃣ Checking spreadsheet...');
    const sheet = getCachedSheet('TargetOrders');
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const confCol = headers.indexOf('Confirmation_Number');
    
    if (confCol !== -1) {
      const sheetConfNum = sheet.getRange(lastRow, confCol + 1).getValue();
      console.log('   Spreadsheet has:', sheetConfNum);
      const sheetMatch = sheetConfNum === testData.Confirmation_Number;
      console.log('   Match:', sheetMatch ? '✅ PASS' : '❌ FAIL');
    }
    
    console.log('\n==========================================');
    if (confMatch && result.emailSent) {
      console.log('🎉 ALL TESTS PASSED!');
      console.log('✅ Ready to deploy');
      console.log('📧 Check your email inbox now');
      return { success: true };
    } else {
      console.log('❌ TESTS FAILED - DO NOT DEPLOY');
      return { success: false };
    }
    
  } catch (error) {
    console.error('\n❌ TEST ERROR:', error);
    console.error('Stack:', error.stack);
    return { success: false, error: error.toString() };
  }
}

/**
 * Test just the email sending (simpler test)
 */
function testEmailOnly() {
  console.log('📧 Testing email sending only...');
  
  const testFormData = {
    form_type: 'target-order',
    Student_Organization: 'Test Organization',
    Event_Name: 'Email Test Event',
    Event_Date: '2025-12-31',
    Submission_Email_Address: 'greg@brooklyn.cuny.club', // ← CHANGE THIS
    Pickup_Person_Email: 'greg@brooklyn.cuny.club',      // ← CHANGE THIS
    Confirmation_Number: 'CD-T-TEST-EMAIL-123'
  };
  
  try {
    const timestamp = getFormattedTimestamp();
    const result = sendConfirmationEmail(
      testFormData,
      testFormData.form_type,
      testFormData.Confirmation_Number,
      timestamp
    );
    
    console.log('Result:', result);
    
    if (result.sent) {
      console.log('✅ Email sent successfully');
      console.log('📧 Check inbox:', testFormData.Submission_Email_Address);
      return { success: true };
    } else {
      console.log('❌ Email failed:', result.error);
      return { success: false, error: result.error };
    }
    
  } catch (error) {
    console.error('❌ Error:', error);
    return { success: false, error: error.toString() };
  }
}
