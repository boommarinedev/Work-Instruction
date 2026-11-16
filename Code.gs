/*******************************************************
 * SALES ORDER → TELEGRAM NOTIFICATION
 *******************************************************/

const CONFIG = {
  SHEET_NAME: 'Sales Orders',

  // Column numbers
  COL_DATE_PLACED: 1,       // A
  COL_WO_NO: 2,             // B
  COL_CUSTOMER: 3,          // C
  COL_ORDER: 4,             // D
  COL_SPECIFICATIONS: 5,    // E
  COL_DATE_NEEDED: 6,       // F
  COL_PLACE_ORDER: 7,       // G
  COL_WI_REF_NO: 8,          // H
  MENTIONS: '@<a href="tg://user?id=8710796814">Cheryl</a>'
};


/**
 * RUN THIS ONCE.
 *
 * Put your Telegram Bot Token and Chat ID here.
 */
function setTelegramConfig() {

  const token = '8886831274:AAFnoyOF5brEgKk6_vvv9GP5yK5CUFjRiHo';
  const chatIds = '-5315631926';
  PropertiesService.getScriptProperties().setProperties({
    TELEGRAM_BOT_TOKEN: token,
    TELEGRAM_CHAT_ID: chatIds
  });

  Logger.log('Telegram configuration saved.');
}


/**
 * TEST TELEGRAM CONNECTION
 */
function testTelegram() {

  const message =
    '🔔 <b>Telegram Test Successful</b>\n\n' +
    'Sales Orders Google Sheet is successfully connected to Telegram.';

  sendTelegramMessage(message);
}


/**
 * INSTALLABLE ON-EDIT TRIGGER
 *
 * Runs when Place Order (Column G)
 * is changed to "Notify".
 */

/**
 * INSTALLABLE ON-EDIT TRIGGER
 *
 * Runs when Date Needed (Column F) is filled in.
 * Sends the notification once per order.
 */
function handleSalesOrderEdit(e) {

  const range = e.range;
  const sheet = range.getSheet();

  // Only Sales Orders sheet
  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    return;
  }

  // Only when the edit touches Column F (Date Needed)
  const firstCol = range.getColumn();
  const lastCol = range.getLastColumn();
  if (CONFIG.COL_DATE_NEEDED < firstCol || CONFIG.COL_DATE_NEEDED > lastCol) {
    return;
  }

  // Avoid two rapid edits sending duplicates
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return;
  }

  try {
    // Handles single edits and pasted multi-row ranges
    for (let row = range.getRow(); row <= range.getLastRow(); row++) {

      // Skip header row
      if (row === 1) continue;

      // Skip if Date Needed is empty (e.g. cell was cleared)
      const dateNeeded = sheet
        .getRange(row, CONFIG.COL_DATE_NEEDED)
        .getDisplayValue()
        .trim();
      if (dateNeeded === '') continue;

      // Skip if already notified (prevents re-sending when the date is corrected later)
      const status = sheet
        .getRange(row, CONFIG.COL_PLACE_ORDER)
        .getDisplayValue()
        .trim();
      if (status === 'Notification Sent') continue;

      // Make sure the key fields are filled before sending
      const woNo = sheet.getRange(row, CONFIG.COL_WO_NO).getDisplayValue().trim();
      const customer = sheet.getRange(row, CONFIG.COL_CUSTOMER).getDisplayValue().trim();
      const order = sheet.getRange(row, CONFIG.COL_ORDER).getDisplayValue().trim();

      if (!woNo || !customer || !order) {
        sheet.getParent().toast(
          '⚠️ Row ' + row + ': WO No., Customer, and Order must be filled before notifying.',
          'Sales Order',
          8
        );
        continue;
      }

      processNotification(sheet, row);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * PROCESS SALES ORDER NOTIFICATION
 */
function processNotification(sheet, row) {

  const placeOrderCell =
    sheet.getRange(row, CONFIG.COL_PLACE_ORDER);

  // Read order information
  const datePlaced =
    sheet.getRange(row, CONFIG.COL_DATE_PLACED)
      .getDisplayValue();

  const woNo =
    sheet.getRange(row, CONFIG.COL_WO_NO)
      .getDisplayValue();

  const customer =
    sheet.getRange(row, CONFIG.COL_CUSTOMER)
      .getDisplayValue();

  const order =
    sheet.getRange(row, CONFIG.COL_ORDER)
      .getDisplayValue();

  const specifications =
    sheet.getRange(row, CONFIG.COL_SPECIFICATIONS)
      .getDisplayValue();

  const dateNeeded =
    sheet.getRange(row, CONFIG.COL_DATE_NEEDED)
      .getDisplayValue();

  const wiRefNo =
    sheet.getRange(row, CONFIG.COL_WI_REF_NO)
      .getDisplayValue();

const notificationDateTime =
  Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'MMMM d, yyyy — h:mm a'
  );

  // Build Telegram message
  const message = '━━━━━━━━━━━━━━━━━━━━\n' +
  '🔔 <b>NEW SALES ORDER</b>\n' +
  '🕐 ' + notificationDateTime + '\n\n' +

  '📋 <b>WO No.:</b> ' +
  escapeHtml(woNo) + '\n' +

  '👤 <b>Customer:</b> ' +
  escapeHtml(customer) + '\n\n' +

  '📦 <b>Order:</b>\n' +
  escapeHtml(order) + '\n\n' +

  '📅 <b>Date Needed:</b> ' +
  escapeHtml(dateNeeded) +'\n\n' +

    '👥 ' + CONFIG.MENTIONS;


  try {

    const result =
      sendTelegramMessage(message);

    if (result && result.ok) {

      // Mark as successfully notified
      placeOrderCell.setValue('Notification Sent');

      sheet.getParent().toast(
        '✅ Telegram notification sent.',
        'Sales Order',
        5
      );

    } else {

      sheet.getParent().toast(
        '❌ Telegram notification failed.',
        'Sales Order',
        8
      );
    }

  } catch (error) {

    // Keep "Notify" if sending fails
    sheet.getParent().toast(
      '❌ Telegram error: ' + error.message,
      'Sales Order',
      8
    );

    console.error(error);
  }
}


/**
 * SEND TELEGRAM MESSAGE
 */
function sendTelegramMessage(message) {

  const properties =
    PropertiesService.getScriptProperties();

  const token =
    properties.getProperty('TELEGRAM_BOT_TOKEN');

  const chatIds =
    properties.getProperty('TELEGRAM_CHAT_ID');


  if (!token || !chatIds) {

    throw new Error(
      'Telegram configuration is missing. ' +
      'Run setTelegramConfig() first.'
    );
  }


  const url =
    'https://api.telegram.org/bot' +
    token +
    '/sendMessage';


  const payload = {
    chat_id: chatIds,
    text: message,
    parse_mode: 'HTML'
  };


  const response =
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });


  const result =
    JSON.parse(response.getContentText());

  Logger.log(result);

  return result;
}


/**
 * ESCAPE TELEGRAM HTML CHARACTERS
 */
function escapeHtml(value) {

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}