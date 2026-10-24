
/**
 * THE MOBILIZER'S DIARY — PAYSTACK DONATION ENGINE
 *
 * Required Script Property:
 * PAYSTACK_SECRET_KEY
 *
 * Creates donation references, records pending payments, and verifies
 * transactions directly with Paystack before confirming a donation.
 */

function donStart(d) {
  d = d || {};

  if (d.hp) return { ref: '' };
  if (d.consent !== true) {
    throw new Error('Donation consent is required.');
  }

  var amount = Number(d.amount);
  var email = String(d.email || '').trim().toLowerCase();
  var name = String(d.name || '').trim().substring(0, 150);

  if (!isFinite(amount) || Math.floor(amount) !== amount || amount < 100) {
    throw new Error('Enter a valid donation amount of at least ₦100.');
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Enter a valid email address.');
  }

  var props = PropertiesService.getScriptProperties();
  var secretKey = String(props.getProperty('PAYSTACK_SECRET_KEY') || '').trim();

  if (!secretKey) {
    throw new Error('Paystack is not configured. Please contact the site administrator.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  var reference;
  try {
    var year = new Date().getFullYear();
    var counterKey = 'DONATION_REF_COUNTER_' + year;
    var counter = Number(props.getProperty(counterKey) || 0);

    var sheet = donSheet_();
    var lastRow = sheet.getLastRow();

    // Continue above any existing references for this year.
    if (lastRow > 1) {
      var refs = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
      refs.forEach(function (row) {
        var match = String(row[0]).match(
          new RegExp('^DON-' + year + '-(\\d{4})$')
        );
        if (match) counter = Math.max(counter, Number(match[1]));
      });
    }

    counter++;
    if (counter > 9999) {
      throw new Error('Donation reference limit reached for this year.');
    }

    props.setProperty(counterKey, String(counter));
    reference = 'DON-' + year + '-' +
      ('0000' + counter).slice(-4);

    sheet.appendRow([
      new Date(),
      reference,
      email,
      amount,
      'NGN',
      'PENDING',
      '',
      name,
      d.anonymous === true ? 'Yes' : 'No'
    ]);
  } finally {
    lock.releaseLock();
  }

  return { ref: reference };
}



/**
 * Verify a donation reference with Paystack.
 * A browser callback alone never confirms a payment.
 */
function donVerify_(reference) {
  reference = String(reference || '').trim();

if (!/^DON-\d{4}-\d{4}$/.test(reference)) {


  var props = PropertiesService.getScriptProperties();
  var secretKey = String(props.getProperty('PAYSTACK_SECRET_KEY') || '').trim();

  if (!secretKey) {
    throw new Error('Paystack secret key is not configured.');
  }

  var sheet = donSheet_();
  var rowNumber = donFindRow_(sheet, reference);

  if (!rowNumber) return { status: 'unknown' };

  var row = sheet.getRange(rowNumber, 1, 1, 9).getValues()[0];
  var expectedAmountKobo = Math.round(Number(row[3]) * 100);
  var email = String(row[2] || '').trim().toLowerCase();

  var response = UrlFetchApp.fetch(
    'https://api.paystack.co/transaction/verify/' +
      encodeURIComponent(reference),
    {
      method: 'get',
      headers: { Authorization: 'Bearer ' + secretKey },
      muteHttpExceptions: true
    }
  );

  var httpCode = response.getResponseCode();
  var result;

  try {
    result = JSON.parse(response.getContentText());
  } catch (e) {
    throw new Error('Paystack returned an unreadable verification response.');
  }

  if (httpCode < 200 || httpCode >= 300 || !result || !result.status) {
    throw new Error('Paystack could not verify this transaction.');
  }

  var txn = result.data || {};
  var status = String(txn.status || '').toLowerCase();

  // Confirm only a successful transaction with matching reference,
  // expected amount, currency, and payer email.
  if (status === 'success') {
    var valid =
      String(txn.reference || '') === reference &&
      Number(txn.amount) === expectedAmountKobo &&
      String(txn.currency || '').toUpperCase() === 'NGN' &&
      String((txn.customer || {}).email || '').trim().toLowerCase() === email;

    if (!valid) {
      sheet.getRange(rowNumber, 6).setValue('REVIEW');
      sheet.getRange(rowNumber, 7).setValue(
        'Verification mismatch; manual review required'
      );
      return { status: 'pending', ref: reference };
    }

    sheet.getRange(rowNumber, 6).setValue('SUCCESS');
    sheet.getRange(rowNumber, 7).setValue(txn.channel || '');
    return { status: 'success', ref: reference };
  }

  if (status === 'abandoned') {
    sheet.getRange(rowNumber, 6).setValue('ABANDONED');
    return { status: 'abandoned', ref: reference };
  }

  if (status === 'failed') {
    sheet.getRange(rowNumber, 6).setValue('FAILED');
    return { status: 'failed', ref: reference };
  }

  return { status: 'pending', ref: reference };
}


/**
 * Compatibility function retained from the previous engine.
 */
function MD22_verifyPaystackPayment_(d) {
  d = d || {};
  var reference = String(d.reference || d.ref || '').trim();
  if (!reference) throw new Error('Transaction reference is missing.');

  var result = donVerify_(reference);

  if (result.status !== 'success') {
    throw new Error('Payment has not been verified as successful.');
  }

  var sheet = donSheet_();
  var rowNumber = donFindRow_(sheet, reference);
  var row = sheet.getRange(rowNumber, 1, 1, 9).getValues()[0];

  return {
    ok: true,
    message: 'Payment verified and logged successfully.',
    reference: reference,
    amount: Number(row[3])
  };
}


/**
 * Creates the donation ledger if it does not already exist.
 */
function donSheet_() {
  var ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('CMS_ID')
  );

  var sheet = ss.getSheetByName('Transactions');

  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'Reference',
      'Payer Email',
      'Amount (NGN)',
      'Currency',
      'Status',
      'Channel',
      'Donor Name',
      'Anonymous'
    ]);
  }

  return sheet;
}


function donFindRow_(sheet, reference) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var found = sheet.getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(reference)
    .matchEntireCell(true)
    .findNext();

  return found ? found.getRow() : 0;
}
