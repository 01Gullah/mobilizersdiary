/**
 * MD_22_Paystack_Engine / MD_13_WEB_API Extension
 * Verifies Paystack payments and logs them to the Google Sheet database.
 */
function adminVerifyPaystackPayment(data) {
  var reference = data.reference;
  if (!reference) {
    throw new Error('Transaction reference is missing.');
  }

  // Get Secret Key from Script Properties
  var secretKey = PropertiesService.getScriptProperties().getProperty("PAYSTACK_SECRET_KEY");
  if (!secretKey) {
    throw new Error('Paystack secret key is not configured in Script Properties.');
  }

  // Call Paystack Verification API
  var url = "https://api.paystack.co/transaction/verify/" + encodeURIComponent(reference);
  var options = {
    "method": "GET",
    "headers": {
      "Authorization": "Bearer " + secretKey
    },
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json.status && json.data && json.data.status === "success") {
    var txn = json.data;
    
    // Open or create the "Transactions" sheet to store payments
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Transactions");
    if (!sheet) {
      sheet = ss.insertSheet("Transactions");
      sheet.appendRow(["Timestamp", "Reference", "Payer Email", "Amount (NGN)", "Currency", "Status", "Channel"]);
    }
    
    // Check if transaction was already logged to prevent duplicates
    var existingRefs = sheet.getRange("B:B").getValues().flat();
    if (existingRefs.indexOf(txn.reference) === -1) {
      sheet.appendRow([
        new Date(),
        txn.reference,
        txn.customer ? txn.customer.email : 'N/A',
        txn.amount / 100, // Convert kobo to NGN
        txn.currency,
        txn.status,
        txn.channel || 'N/A'
      ]);
    }

    return {
      ok: true,
      message: "Payment of ₦" + (txn.amount / 100) + " verified and logged successfully!",
      reference: txn.reference
    };
  } else {
    throw new Error('Payment verification failed on Paystack.');
  }
}
