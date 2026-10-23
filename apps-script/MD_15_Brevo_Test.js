/**
 * One-recipient, Brevo-only test sender for the launch-edition email.
 *
 * Create the HTML file MD_Launch_Edition_Email in the same Apps Script project
 * and paste the matching template. Set these Script Properties privately:
 *   BREVO_API_KEY
 *   BREVO_SENDER_EMAIL       (verified in Brevo)
 *   NEWSLETTER_TEST_EMAIL    (your own address only)
 * Optional:
 *   BREVO_REPLY_TO_EMAIL
 *   NEWSLETTER_PORTAL_URL
 *
 * The founder portrait is the verified public Google Drive image URL below.
 * This function sends ONE message directly through Brevo. It does not use Gmail,
 * the subscriber list, or the campaign queue. Do not run it with a student list.
 */
var MD15_FOUNDER_IMAGE_URL_ = 'https://drive.google.com/uc?export=view&id=1NGT6YyH_qNS_qYeFlUPKcRCll4hbkn9P';

function sendLaunchEditionTestViaBrevo( ) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BREVO_API_KEY');
  var senderEmail = String(props.getProperty('BREVO_SENDER_EMAIL') || '').trim();
  var testEmail = String(props.getProperty('NEWSLETTER_TEST_EMAIL') || '').trim();
  var founderImageUrl = MD15_FOUNDER_IMAGE_URL_;
  var replyTo = String(props.getProperty('BREVO_REPLY_TO_EMAIL') || '').trim();
  var portalUrl = String(props.getProperty('NEWSLETTER_PORTAL_URL') || 'https://01gullah.github.io/mobilizersdiary/?v=4' ).trim();

  if (!apiKey) throw new Error('Set BREVO_API_KEY in Apps Script Script Properties first.');
  if (!MD15_isEmail_(senderEmail)) throw new Error('Set BREVO_SENDER_EMAIL to a valid sender address verified in Brevo.');
  if (!MD15_isEmail_(testEmail)) throw new Error('Set NEWSLETTER_TEST_EMAIL to your own valid test address.');
  if (founderImageUrl.indexOf('https://' ) !== 0) throw new Error('The configured founder portrait URL must use HTTPS.');
  if (portalUrl.indexOf('https://' ) !== 0) throw new Error('NEWSLETTER_PORTAL_URL must use HTTPS.');
  if (replyTo && !MD15_isEmail_(replyTo)) throw new Error('BREVO_REPLY_TO_EMAIL must be a valid email address.');

  var template;
  try {
    template = HtmlService.createTemplateFromFile('MD_Launch_Edition_Email');
  } catch (templateError) {
    throw new Error('Could not load HTML file MD_Launch_Edition_Email. Create that HTML file and paste the supplied template first.');
  }
  template.founderImageUrl = founderImageUrl;
  template.portalUrl = portalUrl;
  template.unsubscribeUrl = '';
  template.isTest = true;
  var htmlBody = template.evaluate().getContent();
  var textBody =
    "TEST SEND — The Mobilizer's Diary Launch Edition\n\n" +
    'From lived student experience to public voice\n\n' +
    'A note from Comr. (Engr.) Lukman Aminu Gullah, Founder, The Mobilizer\'s Diary.\n\n' +
    'During my time at the Federal University of Technology, Minna, I experienced and witnessed concerns around the cost of living, student safety, accountability, and representation. Those experiences left me asking: Where can student realities be documented carefully and heard constructively?\n\n' +
    'We built The Mobilizer\'s Diary as one answer: an independent, non-partisan platform for students to share experiences, propose solutions, and help create a record of the issues that matter.\n\n' +
    'We have taken our first step by building this platform. Now we invite you to take yours: share an experience, suggest a practical solution, nominate a student voice, or pass the platform on to a classmate.\n\n' +
    'Our aim is to listen, handle submissions respectfully, and amplify student voices responsibly.\n\n' +
    'Report. Propose solutions. Be part of the record.\n' + portalUrl + '\n\n' +
    'TEST ONLY — this message is not a live campaign.';

  var payload = {
    sender: { name: "The Mobilizer's Diary", email: senderEmail },
    to: [{ email: testEmail }],
    subject: '[TEST] Students deserve to be heard — Launch Edition',
    htmlContent: htmlBody,
    textContent: textBody
  };
  if (replyTo) payload.replyTo = { email: replyTo };

  var response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'api-key': apiKey, 'accept': 'application/json' },
    payload: JSON.stringify(payload ),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    var detail = String(response.getContentText() || '').substring(0, 500);
    throw new Error('Brevo rejected the test message (HTTP ' + code + '): ' + detail);
  }
  Logger.log('Brevo accepted one launch-edition test email (HTTP ' + code + '). Check the test inbox and Spam folder.');
  return { ok: true, provider: 'Brevo', httpStatus: code };
}

function MD15_isEmail_(value ) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}
