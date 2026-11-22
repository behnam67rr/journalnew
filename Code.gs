/**
 * Zehn Trader 2.0 — Google Sheets read-only API bridge
 *
 * این اسکریپت را در یک پروژه مستقل Google Apps Script قرار بده.
 * آن را به پروژه فعلی ربات MT5 اضافه نکن؛ تا doGet/doPost ربات دست‌نخورده بماند.
 *
 * سازگار با index.html ذهن تریدر 2.0:
 *   GET /exec?action=journal&token=...&sheet=Journal&callback=__ztSheetCallback
 *
 * دسترسی این API فقط خواندن است و هیچ سلولی را تغییر نمی‌دهد.
 */

// شناسه فایل Google Sheets را از URL، بین /d/ و /edit، اینجا قرار بده.
const ZT_SPREADSHEET_ID = '1rRX8-P3UeA0RhrBh8TA3D07TJ0sNtJwlaMk0PneoKRI';
const ZT_DEFAULT_SHEET_NAME = 'Journal';
const ZT_TOKEN_PROPERTY = 'ZT_API_TOKEN';
const ZT_ALLOWED_CALLBACK = '__ztSheetCallback';

/**
 * راه‌اندازی اولیه:
 * 1) شناسه شیت را در بالای فایل قرار بده.
 * 2) این تابع را یک بار از محیط Apps Script اجرا کن.
 * 3) مجوز Google Sheets را تأیید کن.
 * 4) کلید چاپ‌شده در Execution log را در ذهن تریدر وارد کن.
 *
 * اگر کلید قبلاً ساخته شده باشد، دوباره همان کلید را نمایش می‌دهد.
 */
function ztInitialize() {
  ztValidateConfiguration_();

  const spreadsheet = SpreadsheetApp.openById(ZT_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(ZT_DEFAULT_SHEET_NAME);
  if (!sheet) {
    throw new Error('شیت «' + ZT_DEFAULT_SHEET_NAME + '» در فایل Google Sheets پیدا نشد. نام شیت را بررسی کن.');
  }

  const properties = PropertiesService.getScriptProperties();
  let token = properties.getProperty(ZT_TOKEN_PROPERTY);

  if (!token) {
    token = ztGenerateToken_();
    properties.setProperty(ZT_TOKEN_PROPERTY, token);
  }

  Logger.log('Zehn Trader API Token: ' + token);
  Logger.log('Spreadsheet: ' + spreadsheet.getName());
  Logger.log('Sheet: ' + ZT_DEFAULT_SHEET_NAME);
  Logger.log('Initialization completed. Copy the API token into Zehn Trader settings.');

  return 'راه‌اندازی انجام شد. کلید API را از Execution log کپی کن.';
}

/**
 * ساخت کلید جدید و باطل کردن کلید قبلی.
 * فقط زمانی اجرا کن که قصد تعویض کلید را داری؛ سپس کلید جدید را در برنامه وارد کن.
 */
function ztRotateToken() {
  ztValidateConfiguration_();
  const spreadsheet = SpreadsheetApp.openById(ZT_SPREADSHEET_ID);
  if (!spreadsheet.getSheetByName(ZT_DEFAULT_SHEET_NAME)) {
    throw new Error('شیت Journal پیدا نشد.');
  }

  const token = ztGenerateToken_();
  PropertiesService.getScriptProperties().setProperty(ZT_TOKEN_PROPERTY, token);
  Logger.log('NEW Zehn Trader API Token: ' + token);
  Logger.log('کلید قبلی باطل شد؛ کلید جدید را در تنظیمات ذهن تریدر جایگزین کن.');
  return 'کلید جدید ساخته شد. آن را از Execution log کپی کن.';
}

/** بررسی دسترسی اسکریپت به فایل و شیت؛ بدون نمایش کلید API. */
function ztTestConnection() {
  ztValidateConfiguration_();
  const spreadsheet = SpreadsheetApp.openById(ZT_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(ZT_DEFAULT_SHEET_NAME);
  if (!sheet) throw new Error('شیت Journal پیدا نشد.');

  const result = {
    spreadsheet: spreadsheet.getName(),
    sheet: sheet.getName(),
    rows: sheet.getLastRow(),
    columns: sheet.getLastColumn(),
    tokenConfigured: Boolean(PropertiesService.getScriptProperties().getProperty(ZT_TOKEN_PROPERTY))
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/** نقطه ورود وب‌اپ؛ پاسخ JSON یا JSONP مطابق نیاز مرورگر برنامه. */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const callback = String(params.callback || '');
  let payload;

  try {
    ztValidateConfiguration_();

    const savedToken = PropertiesService.getScriptProperties().getProperty(ZT_TOKEN_PROPERTY);
    const suppliedToken = String(params.token || '');

    if (!savedToken) {
      throw new Error('API token تنظیم نشده است. ابتدا تابع ztInitialize را اجرا کن.');
    }
    if (!suppliedToken || !ztSafeEqual_(suppliedToken, savedToken)) {
      throw new Error('کلید دسترسی نامعتبر است. API Token را در تنظیمات ذهن تریدر بررسی کن.');
    }

    const action = String(params.action || 'journal').toLowerCase();
    if (action !== 'journal') {
      throw new Error('درخواست API ناشناخته است.');
    }

    const sheetName = String(params.sheet || ZT_DEFAULT_SHEET_NAME).trim();
    if (!sheetName || sheetName.length > 80) {
      throw new Error('نام شیت معتبر نیست.');
    }

    payload = ztReadJournal_(sheetName);
  } catch (error) {
    payload = {
      status: 'error',
      message: String(error && error.message ? error.message : error),
      trades: [],
      count: 0
    };
  }

  // فقط callback ثابت و مورد انتظار برنامه پذیرفته می‌شود.
  if (callback) {
    if (callback !== ZT_ALLOWED_CALLBACK) {
      return ContentService
        .createTextOutput('/* Invalid callback */')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    const jsonp = JSON.stringify(payload)
      .replace(/</g, '\\u003c')
      .replace(/>/g, '\\u003e')
      .replace(/&/g, '\\u0026')
      .replace(/\u2028/g, '\\u2028')
      .replace(/\u2029/g, '\\u2029');

    return ContentService
      .createTextOutput('window.' + ZT_ALLOWED_CALLBACK + ' && window.' + ZT_ALLOWED_CALLBACK + '(' + jsonp + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** خواندن سطرهای شیت و تبدیل آن‌ها به آرایه‌ای از objectها. */
function ztReadJournal_(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(ZT_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    return {
      status: 'error',
      message: 'شیت «' + sheetName + '» پیدا نشد.',
      sheet: sheetName,
      trades: [],
      count: 0
    };
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return {
      status: 'success',
      sheet: sheetName,
      headers: [],
      trades: [],
      count: 0,
      syncedAt: new Date().toISOString()
    };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(function (header, index) {
    const name = header === null || header === undefined || String(header).trim() === ''
      ? 'column_' + (index + 1)
      : String(header).trim();
    return name;
  });

  const trades = [];
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const isEmpty = row.every(function (value) {
      return value === '' || value === null || value === undefined;
    });
    if (isEmpty) continue;

    const record = {};
    headers.forEach(function (header, columnIndex) {
      record[header] = ztSerializeValue_(row[columnIndex]);
    });
    trades.push(record);
  }

  return {
    status: 'success',
    sheet: sheetName,
    headers: headers,
    trades: trades,
    count: trades.length,
    syncedAt: new Date().toISOString()
  };
}

/** تبدیل مقادیر Date به ISO؛ سایر مقادیر JSON-compatible باقی می‌مانند. */
function ztSerializeValue_(value) {
  if (value instanceof Date) return value.toISOString();
  if (value === null || value === undefined) return '';
  return value;
}

/** تولید توکن تصادفی ۲۵۶ بیتی با استفاده از UUIDهای Google Apps Script. */
function ztGenerateToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

/** مقایسه زمان‌ثابت برای جلوگیری از مقایسه ساده رشته کلید. */
function ztSafeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let difference = 0;
  for (let i = 0; i < a.length; i++) {
    difference |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return difference === 0;
}

/** بررسی تنظیم شناسه فایل قبل از اجرای توابع اصلی. */
function ztValidateConfiguration_() {
  if (!ZT_SPREADSHEET_ID || ZT_SPREADSHEET_ID === 'PASTE_GOOGLE_SHEET_ID_HERE') {
    throw new Error('ابتدا مقدار ZT_SPREADSHEET_ID را در ابتدای فایل با شناسه Google Sheets خودت جایگزین کن.');
  }
}
