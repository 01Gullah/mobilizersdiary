/*
 * THE MOBILIZER'S DIARY — PAYSTACK VERIFICATION ENGINE
 *
 * File Name: MD_22_Paystack_Engine.gs
 * Handles Paystack API verification calls and logs payments to Google Sheets.
 */

function MD22_verifyPaystackPayment_(d) {
  d = d || {};
  var reference = String(d.reference || '').trim();
  if (!reference) {
    throw new Error('Transaction reference is missing.');
  }

  // Retrieve Secret Key from Script Properties
  var props = PropertiesService.getScriptProperties();
  var secretKey = props.getProperty('PAYSTACK_SECRET_KEY');
  if (!secretKey) {
    throw new Error('Paystack secret key is not configured in Script Properties.');
  }

  // Query Paystack Transaction Verification Endpoint
  var url = 'https://api.paystack.co/transaction/verify/' + encodeURIComponent(reference);
  var options = {
    'method': 'GET',
    'headers': {
      'Authorization': 'Bearer ' + secretKey
    },
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json.status && json.data && json.data.status === 'success') {
    var txn = json.data;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transactions');
    if (!sheet) {
      sheet = ss.insertSheet('Transactions');
      sheet.appendRow(['Timestamp', 'Reference', 'Payer Email', 'Amount (NGN)', 'Currency', 'Status', 'Channel']);
    }

    // Duplicate check using column B (Reference)
    var lastRow = sheet.getLastRow();
    var isDuplicate = false;
    if (lastRow > 1) {
      var refs = sheet.getRange(1, 2, lastRow, 1).getValues().flat();
      if (refs.indexOf(txn.reference) !== -1) {
        isDuplicate = true;
      }
    }

    if (!isDuplicate) {
      sheet.appendRow([
        new Date(),
        txn.reference,
        txn.customer ? txn.customer.email : 'N/A',
        txn.amount / 100, // Kobo to NGN
        txn.currency,
        txn.status,
        txn.channel || 'N/A'
      ]);
    }

    return {
      ok: true,
      message: 'Payment of ₦' + (txn.amount / 100) + ' verified and logged successfully.',
      reference: txn.reference,
      amount: txn.amount / 100
    };
  } else {
    throw new Error('Paystack verification failed: Transaction not successful.');
  }
}
