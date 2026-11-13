const SHEET_NAME = 'Responses';
const HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];
const PAYMENT_AMOUNT = 10.00;
const CASHFREE_API_VERSION = '2023-08-01';

function getResponsesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveResponse_(params) {
  const name = (params.name || '').toString().trim();
  const email = (params.email || '').toString().trim();
  const phone = (params.phone || '').toString().trim();

  if (!name || !email || !phone) {
    return jsonOutput_({ success: false, error: 'name, email and phone are required' });
  }

  getResponsesSheet_().appendRow([new Date(), name, email, phone]);
  return jsonOutput_({ success: true });
}

function getCashfreeConfig_() {
  // Store CASHFREE_CLIENT_ID, CASHFREE_CLIENT_SECRET and CASHFREE_ENV in Script Properties.
  const properties = PropertiesService.getScriptProperties();
  const clientId = properties.getProperty('CASHFREE_CLIENT_ID');
  const clientSecret = properties.getProperty('CASHFREE_CLIENT_SECRET');
  const environment = (properties.getProperty('CASHFREE_ENV') || 'sandbox').toLowerCase();

  if (!clientId || !clientSecret) {
    throw new Error('Cashfree is not configured. Add CASHFREE_CLIENT_ID and CASHFREE_CLIENT_SECRET in Apps Script project settings.');
  }
  if (environment !== 'sandbox' && environment !== 'production') {
    throw new Error('CASHFREE_ENV must be either sandbox or production.');
  }

  return {
    clientId: clientId,
    clientSecret: clientSecret,
    apiUrl: environment === 'production'
      ? 'https://api.cashfree.com/pg'
      : 'https://sandbox.cashfree.com/pg',
    mode: environment === 'production' ? 'production' : 'sandbox'
  };
}

function cashfreeRequest_(config, method, path, body) {
  const options = {
    method: method,
    headers: {
      'x-client-id': config.clientId,
      'x-client-secret': config.clientSecret,
      'x-api-version': CASHFREE_API_VERSION,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  if (body) {
    options.payload = JSON.stringify(body);
  }

  const response = UrlFetchApp.fetch(config.apiUrl + path, options);
  const status = response.getResponseCode();
  const responseText = response.getContentText();
  let result;
  try {
    result = JSON.parse(responseText);
  } catch (error) {
    throw new Error('Cashfree returned an unreadable response (HTTP ' + status + ').');
  }
  if (status < 200 || status >= 300) {
    const message = result.message || result.type || ('HTTP ' + status);
    throw new Error('Cashfree could not create or verify the payment: ' + message);
  }
  return result;
}

function escapeHtml_(value) {
  return String(value).replace(/[&<>"']/g, function (character) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[character];
  });
}

function paymentPage_(title, message) {
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + escapeHtml_(title) + '</title>' +
    '<style>body{font:16px system-ui,sans-serif;background:#07090f;color:#e7eaf2;display:grid;place-items:center;min-height:100vh;margin:0;padding:24px;box-sizing:border-box}.card{max-width:520px;padding:32px;border:1px solid #303747;border-radius:16px;background:#111722}h1{font-size:24px}p{color:#b9c1d2;line-height:1.6}</style>' +
    '</head><body><main class="card"><h1>' + escapeHtml_(title) + '</h1><p>' + escapeHtml_(message) + '</p></main></body></html>'
  );
}

function createPayment_(params) {
  const name = (params.name || '').toString().trim();
  const email = (params.email || '').toString().trim();
  const phone = (params.phone || '').toString().trim();

  if (name.length < 2 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !/^[6-9]\d{9}$/.test(phone)) {
    return paymentPage_('Check your details', 'Please return to the enquiry form and enter a valid name, email and 10-digit Indian mobile number.');
  }

  const config = getCashfreeConfig_();
  const orderId = 'reg_' + Date.now() + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  const order = cashfreeRequest_(config, 'post', '/orders', {
    order_id: orderId,
    order_amount: PAYMENT_AMOUNT,
    order_currency: 'INR',
    customer_details: {
      customer_id: 'customer_' + Utilities.getUuid().replace(/-/g, '').slice(0, 24),
      customer_name: name,
      customer_email: email,
      customer_phone: phone
    },
    order_meta: {
      return_url: ScriptApp.getService().getUrl() +
        '?payment_return=1&order_id=' + encodeURIComponent(orderId)
    }
  });

  if (!order.payment_session_id) {
    throw new Error('Cashfree did not return a payment session. Please try again.');
  }

  getResponsesSheet_().appendRow([new Date(), name, email, phone]);
  const sessionId = JSON.stringify(order.payment_session_id).replace(/</g, '\\u003c');
  const mode = JSON.stringify(config.mode);
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Secure payment</title>' +
    '<style>body{font:16px system-ui,sans-serif;background:#07090f;color:#e7eaf2;display:grid;place-items:center;min-height:100vh;margin:0;padding:24px;box-sizing:border-box}.card{max-width:520px;padding:32px;border:1px solid #303747;border-radius:16px;background:#111722}button{padding:12px 20px;border:0;border-radius:10px;background:#38bdf8;color:#04121d;font:600 16px system-ui,sans-serif;cursor:pointer}</style>' +
    '</head><body><main class="card"><p>Your fresh ₹10.00 payment session is ready.</p><button id="payNow" type="button">Continue to secure payment</button></main>' +
    '<script src="https://sdk.cashfree.com/js/v3/cashfree.js"></script>' +
    '<script>const cashfree=Cashfree({mode:' + mode + '});' +
    'document.getElementById("payNow").addEventListener("click",function(){' +
    'cashfree.checkout({paymentSessionId:' + sessionId + ',redirectTarget:"_top"});});</script>' +
    '</body></html>'
  );
}

function paymentReturn_(orderId) {
  if (!orderId) {
    return paymentPage_('Payment status unavailable', 'No payment order was provided. Please return to the enquiry form to start a new payment.');
  }

  const config = getCashfreeConfig_();
  const order = cashfreeRequest_(config, 'get', '/orders/' + encodeURIComponent(orderId));
  if (order.order_status === 'PAID') {
    return paymentPage_('Payment successful', 'Your ₹10.00 payment was received. You can close this page.');
  }
  if (order.order_status === 'ACTIVE') {
    return paymentPage_('Payment not completed', 'This payment is still pending or was not completed. Return to the enquiry form to start a fresh payment.');
  }
  return paymentPage_('Payment not completed', 'Payment status: ' + (order.order_status || 'unknown') + '. Return to the enquiry form to start a fresh payment.');
}

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.payment_return === '1') {
      return paymentReturn_(params.order_id);
    }
    return saveResponse_(params);
  } catch (err) {
    return paymentPage_('Unable to process payment', err.message || err.toString());
  }
}

function doPost(e) {
  try {
    let params = (e && e.parameter) || {};
    if (e && e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);
        params = Object.assign({}, params, body);
      } catch (parseErr) {
        // Form submissions use e.parameter; non-JSON bodies need no extra parsing.
      }
    }
    if (params.action === 'createPayment') {
      return createPayment_(params);
    }
    return saveResponse_(params);
  } catch (err) {
    return paymentPage_('Unable to process payment', err.message || err.toString());
  }
}
