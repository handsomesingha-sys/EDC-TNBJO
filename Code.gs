/**
 * EDC PAYMENT SYSTEM - Google Apps Script Backend
 * Version: 1.0.0
 * Timezone: Asia/Bangkok
 *
 * FRONTEND CONTRACT (index.html)
 * --------------------------------
 * Send POST requests to the deployed Apps Script Web App URL.
 * Use Content-Type: text/plain;charset=utf-8 to avoid browser preflight
 * when calling the Apps Script web app from a GitHub-hosted frontend.
 * Request JSON:
 *   { action: 'login', data: { accessCode: '...' } }
 *   { action: 'createPayment', token: 'SESSION_TOKEN', data: { amount: 100 } }
 * All responses use:
 *   { ok: true, message: '...', data: {...} }
 * or
 *   { ok: false, message: 'ข้อความสำหรับผู้ใช้', code: 'SAFE_CODE' }
 *
 * INITIAL SETUP
 * --------------------------------
 * 1) Create/open the Google Spreadsheet used as the database.
 * 2) Open Extensions > Apps Script and paste this entire file as Code.gs.
 * 3) In Project Settings > Script Properties, set:
 *      SPREADSHEET_ID       = Google Spreadsheet ID
 *      ADMIN_INITIAL_CODE  = initial Admin login code (the code previously
 *                             specified for this project)
 *      THUNDER_API_KEY     = Thunder API key
 *      IMGBB_API_KEY       = ImgBB API key
 * 4) Run setupDatabase() manually once and authorize the requested scopes.
 * 5) Configure Settings sheet values for the PromptPay identifier and store.
 * 6) Deploy as Web App. Execute as the deploying account. Restrict access as
 *    appropriate for the intended users and test from the actual GitHub site.
 *
 * SECURITY NOTES
 * --------------------------------
 * - API secrets stay in Script Properties; they are never returned to frontend.
 * - Access codes are stored as salted SHA-256 hashes, never as plain text.
 * - Session tokens are stored as hashes in the Sessions sheet.
 * - Payment approval happens only on the backend after Thunder verification.
 * - ImgBB URLs are public image URLs. Do not upload sensitive images unless
 *   the business accepts this privacy model. The app follows the requested
 *   ImgBB storage design.
 * - Google Sheets is a lightweight database. For high-concurrency / high-volume
 *   payment production workloads, a transactional database is recommended.
 * - This backend creates refund requests and records approvals. It does NOT
 *   claim to move money back to a payer unless a supported refund API is later
 *   integrated and independently verified.
 */

'use strict';

/* ========================================================================== */
/*  1. GLOBAL CONFIGURATION                                                    */
/* ========================================================================== */

const APP_CONFIG = Object.freeze({
  APP_NAME: 'EDC PAYMENT SYSTEM',
  VERSION: '1.0.0',
  TIMEZONE: 'Asia/Bangkok',
  SESSION_HOURS: 12,
  LOGIN_MAX_ATTEMPTS: 8,
  LOGIN_LOCK_MINUTES: 10,
  MAX_SLIP_BYTES: 4 * 1024 * 1024,
  MAX_GENERAL_IMAGE_BYTES: 4 * 1024 * 1024,
  SUPPORTED_IMAGE_MIMES: ['image/jpeg', 'image/png', 'image/webp'],
  THUNDER_QR_URL: 'https://api.thunder.in.th/v1/qr/generate',
  THUNDER_VERIFY_URL: 'https://api.thunder.in.th/v2/verify/bank',
  IMGBB_UPLOAD_URL: 'https://api.imgbb.com/1/upload',
  DEFAULT_QR_EXPIRATION_MINUTES: 5,
  DEFAULT_CURRENCY: 'THB',
  DEFAULT_DEVICE_NAME: 'EDC DEVICE',
  TOKEN_PREFIX: 'EDC',
  ID_PREFIX: 'EDC'
});

const SHEET_SCHEMAS = Object.freeze({
  Users: [
    'UserID', 'Prefix', 'FullName', 'AccessCodeHash', 'AccessCodeSalt',
    'Role', 'Status', 'CodePrefix', 'CreatedAt', 'UpdatedAt',
    'LastLoginAt', 'FailedLoginCount', 'LockedUntil'
  ],
  Transactions: [
    'TransactionID', 'Amount', 'Currency', 'Status', 'Note',
    'CashierID', 'CashierName', 'DeviceID', 'ShiftID', 'QRType',
    'QRPayload', 'QRMime', 'QRCreatedAt', 'QRExpiresAt', 'PaidAt',
    'SlipImageURL', 'SlipViewerURL', 'ImgBBImageID', 'ThunderTransRef',
    'LastSlipCheckID', 'CreatedAt', 'UpdatedAt', 'CancelledAt',
    'CancelReason'
  ],
  SlipChecks: [
    'CheckID', 'TransactionID', 'AttemptNo', 'ImageURL', 'ViewerURL',
    'ImgBBImageID', 'ThunderSuccess', 'IsDuplicate', 'IsAmountMatched',
    'IsAccountMatched', 'AmountInSlip', 'AmountInOrder', 'TransRef',
    'SenderBank', 'SenderName', 'ReceiverBank', 'ReceiverName',
    'ReceiverAccount', 'SlipDate', 'Result', 'FailureReason',
    'RawResponse', 'CheckedBy', 'CreatedAt'
  ],
  Shifts: [
    'ShiftID', 'UserID', 'UserName', 'DeviceID', 'Status', 'OpenedAt',
    'ClosedAt', 'OpeningNote', 'ClosingNote', 'TransactionCount',
    'TotalAmount', 'CreatedAt', 'UpdatedAt'
  ],
  Devices: [
    'DeviceID', 'DeviceName', 'Status', 'AssignedUserID', 'LastSeenAt',
    'CreatedAt', 'UpdatedAt', 'Notes'
  ],
  Refunds: [
    'RefundID', 'TransactionID', 'Amount', 'Reason', 'Status',
    'RequestedBy', 'RequestedByName', 'RequestedAt', 'ReviewedBy',
    'ReviewedByName', 'ReviewedAt', 'ReviewNote', 'ProviderReference',
    'ProcessedAt', 'CreatedAt', 'UpdatedAt'
  ],
  DailyClosings: [
    'ClosingID', 'BusinessDate', 'DeviceID', 'UserID', 'TransactionCount',
    'SuccessAmount', 'RefundAmount', 'NetAmount', 'ClosedBy',
    'ClosedByName', 'ClosedAt', 'Notes'
  ],
  AuditLogs: [
    'LogID', 'UserID', 'UserName', 'Role', 'Action', 'EntityType',
    'EntityID', 'Details', 'CreatedAt'
  ],
  Notifications: [
    'NotificationID', 'Type', 'Title', 'Message', 'RecipientID',
    'IsRead', 'CreatedAt', 'ReadAt', 'RelatedID'
  ],
  RolePermissions: [
    'Role', 'PermissionCode', 'IsAllowed', 'UpdatedAt'
  ],
  Settings: [
    'SettingKey', 'SettingValue', 'Description', 'UpdatedAt'
  ],
  SystemLogs: [
    'LogID', 'Level', 'PublicMessage', 'InternalDetail',
    'Action', 'CreatedAt'
  ],
  APIUsage: [
    'APIUsageID', 'Provider', 'Endpoint', 'TransactionID',
    'HttpStatus', 'DurationMs', 'Success', 'CreatedAt'
  ],
  MediaFiles: [
    'MediaID', 'ImageType', 'RelatedID', 'ImageURL', 'ViewerURL',
    'ThumbnailURL', 'ImgBBImageID', 'DeleteURL', 'UploadedBy',
    'UploadedAt', 'UploadStatus'
  ],
  Sessions: [
    'SessionID', 'TokenHash', 'UserID', 'Role', 'DeviceID',
    'CreatedAt', 'ExpiresAt', 'LastSeenAt', 'RevokedAt'
  ]
});

const DEFAULT_SETTINGS = Object.freeze([
  ['BUSINESS_NAME', 'EDC PAYMENT', 'ชื่อร้านค้า'],
  ['BUSINESS_ADDRESS', '', 'ที่อยู่ร้านค้า'],
  ['BUSINESS_PHONE', '', 'เบอร์โทรศัพท์ร้านค้า'],
  ['BUSINESS_LOGO_URL', '', 'URL โลโก้ร้านค้า ImgBB'],
  ['QR_EXPIRATION_MINUTES', '5', 'อายุ QR Code เป็นนาที'],
  ['THUNDER_QR_TYPE', 'PROMPTPAY', 'ประเภท QR: PROMPTPAY, KSHOP, MAE_MANEE, TUNGNGERN'],
  ['PROMPTPAY_IDENTIFIER_TYPE', 'msisdn', 'ชนิดตัวระบุ PromptPay: msisdn, natId, eWalletId, billerId'],
  ['PROMPTPAY_IDENTIFIER', '', 'หมายเลข PromptPay หรือ Biller ID'],
  ['PROMPTPAY_BILLER_REF1', '', 'Reference 1 สำหรับ PromptPay แบบ Biller ID'],
  ['PROMPTPAY_BILLER_REF2', '', 'Reference 2 สำหรับ PromptPay แบบ Biller ID'],
  ['THUNDER_MERCHANT_NAME', '', 'ชื่อร้านสำหรับ QR ประเภท Tungngern'],
  ['THUNDER_MATCH_ACCOUNT', 'true', 'ตรวจสอบบัญชีผู้รับกับบัญชีที่ลงทะเบียนใน Thunder'],
  ['THUNDER_CHECK_DUPLICATE', 'true', 'ตรวจสอบสลิปซ้ำ'],
  ['REQUIRE_OPEN_SHIFT', 'false', 'บังคับให้พนักงานเปิดกะก่อนรับชำระ'],
  ['DEFAULT_DEVICE_NAME', 'EDC DEVICE', 'ชื่อเครื่องเริ่มต้น'],
  ['IMG_BB_EXPIRATION_SECONDS', '0', '0 = ไม่กำหนดเวลาหมดอายุภาพ; หรือ 60-15552000 วินาที'],
  ['RECEIPT_FOOTER', 'ขอบคุณที่ใช้บริการ', 'ข้อความท้ายใบเสร็จ'],
  ['SOUND_ENABLED', 'true', 'เปิดเสียงแจ้งเตือน'],
  ['LINE_NOTIFICATION_ENABLED', 'false', 'เปิดการแจ้งเตือน LINE เมื่อเชื่อมต่อระบบแล้ว'],
  ['ALLOW_CASHIER_VIEW_OWN_HISTORY', 'true', 'ให้พนักงานดูประวัติรายการของตนเอง']
]);

const DEFAULT_PERMISSIONS = Object.freeze([
  ['ADMIN', 'payment.create', true],
  ['ADMIN', 'payment.cancel', true],
  ['ADMIN', 'transaction.read.all', true],
  ['ADMIN', 'transaction.read.own', true],
  ['ADMIN', 'staff.read', true],
  ['ADMIN', 'staff.manage', true],
  ['ADMIN', 'reports.read', true],
  ['ADMIN', 'settlement.read', true],
  ['ADMIN', 'shift.manage', true],
  ['ADMIN', 'refund.request', true],
  ['ADMIN', 'refund.review', true],
  ['ADMIN', 'device.manage', true],
  ['ADMIN', 'settings.read', true],
  ['ADMIN', 'settings.manage', true],
  ['ADMIN', 'audit.read', true],
  ['ADMIN', 'systemlogs.read', true],
  ['ADMIN', 'notifications.read', true],
  ['ADMIN', 'media.upload', true],
  ['MANAGER', 'payment.create', true],
  ['MANAGER', 'payment.cancel', true],
  ['MANAGER', 'transaction.read.all', true],
  ['MANAGER', 'transaction.read.own', true],
  ['MANAGER', 'staff.read', false],
  ['MANAGER', 'staff.manage', false],
  ['MANAGER', 'reports.read', true],
  ['MANAGER', 'settlement.read', true],
  ['MANAGER', 'shift.manage', true],
  ['MANAGER', 'refund.request', true],
  ['MANAGER', 'refund.review', false],
  ['MANAGER', 'device.manage', true],
  ['MANAGER', 'settings.read', false],
  ['MANAGER', 'settings.manage', false],
  ['MANAGER', 'audit.read', false],
  ['MANAGER', 'systemlogs.read', false],
  ['MANAGER', 'notifications.read', true],
  ['MANAGER', 'media.upload', true],
  ['CASHIER', 'payment.create', true],
  ['CASHIER', 'payment.cancel', false],
  ['CASHIER', 'transaction.read.all', false],
  ['CASHIER', 'transaction.read.own', true],
  ['CASHIER', 'staff.read', false],
  ['CASHIER', 'staff.manage', false],
  ['CASHIER', 'reports.read', false],
  ['CASHIER', 'settlement.read', false],
  ['CASHIER', 'shift.manage', true],
  ['CASHIER', 'refund.request', false],
  ['CASHIER', 'refund.review', false],
  ['CASHIER', 'device.manage', false],
  ['CASHIER', 'settings.read', false],
  ['CASHIER', 'settings.manage', false],
  ['CASHIER', 'audit.read', false],
  ['CASHIER', 'systemlogs.read', false],
  ['CASHIER', 'notifications.read', true],
  ['CASHIER', 'media.upload', false]
]);

/* ========================================================================== */
/*  2. WEB APP ENTRY POINTS                                                    */
/* ========================================================================== */

function doGet(e) {
  const output = {
    ok: true,
    message: 'EDC Payment API is ready.',
    data: {
      app: APP_CONFIG.APP_NAME,
      version: APP_CONFIG.VERSION,
      timestamp: new Date().toISOString()
    }
  };
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let action = 'unknown';
  try {
    const rawBody = e && e.postData && e.postData.contents
      ? e.postData.contents
      : '{}';
    const request = JSON.parse(rawBody);
    action = String(request.action || '').trim();

    if (!action) {
      return jsonOutput_(errorResponse_('ไม่พบคำสั่งที่ต้องการดำเนินการ', 'MISSING_ACTION'));
    }

    const result = dispatchAction_(action, request);
    return jsonOutput_(result);
  } catch (error) {
    logSystemError_('doPost', error, action);
    return jsonOutput_(errorResponse_(publicMessageForError_(error), 'REQUEST_FAILED'));
  }
}

function jsonOutput_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatchAction_(action, request) {
  const data = request.data || request.payload || {};

  if (action === 'health') {
    return successResponse_('ระบบพร้อมใช้งาน', {
      app: APP_CONFIG.APP_NAME,
      version: APP_CONFIG.VERSION,
      timestamp: new Date().toISOString()
    });
  }

  if (action === 'login') {
    return apiLogin_(data);
  }

  const session = authenticateRequest_(request.token);

  switch (action) {
    case 'logout':
      return apiLogout_(session);
    case 'getBootstrap':
      return apiGetBootstrap_(session);
    case 'createPayment':
      return apiCreatePayment_(session, data);
    case 'verifySlip':
      return apiVerifySlip_(session, data);
    case 'cancelPayment':
      return apiCancelPayment_(session, data);
    case 'listTransactions':
      return apiListTransactions_(session, data);
    case 'getTransaction':
      return apiGetTransaction_(session, data);
    case 'listStaff':
      return apiListStaff_(session, data);
    case 'createStaff':
      return apiCreateStaff_(session, data);
    case 'updateStaff':
      return apiUpdateStaff_(session, data);
    case 'resetStaffCode':
      return apiResetStaffCode_(session, data);
    case 'setStaffStatus':
      return apiSetStaffStatus_(session, data);
    case 'getDashboard':
      return apiGetDashboard_(session, data);
    case 'getReports':
      return apiGetReports_(session, data);
    case 'openShift':
      return apiOpenShift_(session, data);
    case 'closeShift':
      return apiCloseShift_(session, data);
    case 'getMyShift':
      return apiGetMyShift_(session);
    case 'listShifts':
      return apiListShifts_(session, data);
    case 'requestRefund':
      return apiRequestRefund_(session, data);
    case 'reviewRefund':
      return apiReviewRefund_(session, data);
    case 'listRefunds':
      return apiListRefunds_(session, data);
    case 'listDevices':
      return apiListDevices_(session, data);
    case 'saveDevice':
      return apiSaveDevice_(session, data);
    case 'setDeviceStatus':
      return apiSetDeviceStatus_(session, data);
    case 'getSettings':
      return apiGetSettings_(session);
    case 'saveSettings':
      return apiSaveSettings_(session, data);
    case 'listRolePermissions':
      return apiListRolePermissions_(session);
    case 'saveRolePermissions':
      return apiSaveRolePermissions_(session, data);
    case 'changeOwnAccessCode':
      return apiChangeOwnAccessCode_(session, data);
    case 'uploadImage':
      return apiUploadImage_(session, data);
    case 'listMediaFiles':
      return apiListMediaFiles_(session, data);
    case 'getNotifications':
      return apiGetNotifications_(session, data);
    case 'markNotificationRead':
      return apiMarkNotificationRead_(session, data);
    case 'listAuditLogs':
      return apiListAuditLogs_(session, data);
    case 'listSystemLogs':
      return apiListSystemLogs_(session, data);
    case 'listApiUsage':
      return apiListApiUsage_(session, data);
    case 'closeBusinessDay':
      return apiCloseBusinessDay_(session, data);
    case 'getDatabaseStatus':
      return apiGetDatabaseStatus_(session);
    default:
      return errorResponse_('ไม่พบเมนูหรือคำสั่งที่ร้องขอ', 'UNKNOWN_ACTION');
  }
}

/* ========================================================================== */
/*  3. DATABASE INITIALIZATION                                                 */
/* ========================================================================== */

/** Run this function manually from the Apps Script editor once. */
function setupDatabase() {
  const spreadsheet = getDatabaseSpreadsheet_();
  if (!getScriptProperty_('SPREADSHEET_ID')) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());
  }
  spreadsheet.setSpreadsheetTimeZone(APP_CONFIG.TIMEZONE);
  ensureDatabaseStructure_();
  seedDefaultSettings_();
  seedDefaultPermissions_();
  const adminSeed = seedInitialAdmin_();

  logAuditSystem_('SYSTEM', 'ระบบ', 'DATABASE_SETUP', 'DATABASE', spreadsheet.getId(), {
    spreadsheetName: spreadsheet.getName(),
    version: APP_CONFIG.VERSION
  });

  return {
    success: true,
    message: 'สร้างและตรวจสอบโครงสร้างฐานข้อมูลเรียบร้อยแล้ว',
    spreadsheetId: spreadsheet.getId(),
    sheets: Object.keys(SHEET_SCHEMAS),
    adminSeed: adminSeed
  };
}

/**
 * Safe schema migration. Adds missing sheets and missing header columns.
 * Existing rows and existing columns are not deleted or rearranged.
 */
function ensureDatabaseStructure() {
  return ensureDatabaseStructure_();
}

function ensureDatabaseStructure_() {
  const spreadsheet = getDatabaseSpreadsheet_();
  const schemaNames = Object.keys(SHEET_SCHEMAS);

  schemaNames.forEach(function (sheetName) {
    const expectedHeaders = SHEET_SCHEMAS[sheetName];
    let sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      formatHeaderRow_(sheet, expectedHeaders.length);
      return;
    }

    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      .map(function (value) { return String(value || '').trim(); });
    const cleanedHeaders = currentHeaders.filter(function (value) { return value !== ''; });

    if (cleanedHeaders.length === 0) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      formatHeaderRow_(sheet, expectedHeaders.length);
      return;
    }

    const missingHeaders = expectedHeaders.filter(function (header) {
      return cleanedHeaders.indexOf(header) === -1;
    });

    if (missingHeaders.length > 0) {
      const startColumn = Math.max(sheet.getLastColumn(), cleanedHeaders.length) + 1;
      sheet.getRange(1, startColumn, 1, missingHeaders.length).setValues([missingHeaders]);
      formatHeaderRow_(sheet, startColumn + missingHeaders.length - 1);
    }

    sheet.setFrozenRows(1);
  });

  return { success: true, sheetCount: schemaNames.length };
}

function formatHeaderRow_(sheet, columnCount) {
  const range = sheet.getRange(1, 1, 1, columnCount);
  range
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#17324D')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, Math.min(columnCount, 12));
}

function seedDefaultSettings_() {
  const existing = getSheetObjects_('Settings');
  const existingKeys = {};
  existing.forEach(function (row) { existingKeys[String(row.SettingKey)] = true; });

  const now = new Date();
  const rowsToAdd = DEFAULT_SETTINGS
    .filter(function (row) { return !existingKeys[row[0]]; })
    .map(function (row) {
      return {
        SettingKey: row[0],
        SettingValue: row[1],
        Description: row[2],
        UpdatedAt: now
      };
    });

  if (rowsToAdd.length) {
    appendObjects_('Settings', rowsToAdd);
  }
}

function seedDefaultPermissions_() {
  const existing = getSheetObjects_('RolePermissions');
  const map = {};
  existing.forEach(function (row) {
    map[String(row.Role) + '|' + String(row.PermissionCode)] = true;
  });

  const now = new Date();
  const rowsToAdd = DEFAULT_PERMISSIONS
    .filter(function (row) {
      return !map[String(row[0]) + '|' + String(row[1])];
    })
    .map(function (row) {
      return {
        Role: row[0],
        PermissionCode: row[1],
        IsAllowed: row[2],
        UpdatedAt: now
      };
    });

  if (rowsToAdd.length) {
    appendObjects_('RolePermissions', rowsToAdd);
  }
}

function seedInitialAdmin_() {
  const users = getSheetObjects_('Users');
  const adminExists = users.some(function (user) {
    return String(user.Role).toUpperCase() === 'ADMIN';
  });

  if (adminExists) {
    return { created: false, reason: 'ADMIN_ALREADY_EXISTS' };
  }

  ensureAccessCodePepper_();
  const initialCode = getScriptProperty_('ADMIN_INITIAL_CODE');
  if (!initialCode) {
    logSystem_('WARN', 'ยังไม่ได้ตั้งค่ารหัสเริ่มต้นของผู้ดูแลระบบ',
      'Set Script Property ADMIN_INITIAL_CODE, then run setupDatabase() again.',
      'seedInitialAdmin_');
    return { created: false, reason: 'ADMIN_INITIAL_CODE_NOT_SET' };
  }

  const salt = createRandomToken_();
  const now = new Date();
  appendObject_('Users', {
    UserID: 'ADMIN-001',
    Prefix: '',
    FullName: 'System Administrator',
    AccessCodeHash: hashAccessCode_(initialCode, salt),
    AccessCodeSalt: salt,
    Role: 'ADMIN',
    Status: 'ACTIVE',
    CodePrefix: 'ADMIN',
    CreatedAt: now,
    UpdatedAt: now,
    LastLoginAt: '',
    FailedLoginCount: 0,
    LockedUntil: ''
  });

  PropertiesService.getScriptProperties().deleteProperty('ADMIN_INITIAL_CODE');
  return { created: true, userId: 'ADMIN-001' };
}

function getDatabaseStatus_() {
  const spreadsheet = getDatabaseSpreadsheet_();
  const results = [];
  Object.keys(SHEET_SCHEMAS).forEach(function (sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    results.push({
      sheet: sheetName,
      exists: !!sheet,
      rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0,
      columns: sheet ? sheet.getLastColumn() : 0
    });
  });
  return results;
}

/* ========================================================================== */
/*  4. AUTHENTICATION, SESSION AND PERMISSIONS                                 */
/* ========================================================================== */

function apiLogin_(data) {
  const accessCode = String(data.accessCode || '').trim();
  const deviceId = String(data.deviceId || '').trim();

  if (!accessCode) {
    return errorResponse_('กรุณากรอกรหัสเข้าใช้งาน', 'ACCESS_CODE_REQUIRED');
  }

  const attemptCacheKey = 'LOGIN_FAIL_' + hashToken_(accessCode).substring(0, 30);
  const attemptCount = Number(CacheService.getScriptCache().get(attemptCacheKey) || 0);
  if (attemptCount >= APP_CONFIG.LOGIN_MAX_ATTEMPTS) {
    return errorResponse_('ไม่สามารถเข้าสู่ระบบได้ในขณะนี้ กรุณาลองใหม่ภายหลัง', 'LOGIN_TEMPORARILY_LOCKED');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const users = getSheetObjects_('Users');
    const now = new Date();
    let matchedUser = null;

    for (let i = 0; i < users.length; i += 1) {
      const user = users[i];
      if (String(user.Status).toUpperCase() !== 'ACTIVE') continue;
      const lockedUntil = parseDate_(user.LockedUntil);
      if (lockedUntil && lockedUntil.getTime() > now.getTime()) continue;

      const candidateHash = hashAccessCode_(accessCode, String(user.AccessCodeSalt || ''));
      if (constantTimeEquals_(candidateHash, String(user.AccessCodeHash || ''))) {
        matchedUser = user;
        break;
      }
    }

    if (!matchedUser) {
      recordLoginFailure_(accessCode);
      return errorResponse_('รหัสเข้าใช้งานไม่ถูกต้อง หรือบัญชีถูกระงับ', 'LOGIN_INVALID');
    }

    const userRow = findRowByKey_('Users', 'UserID', matchedUser.UserID);
    if (userRow) {
      updateRowObject_('Users', userRow.rowIndex, {
        FailedLoginCount: 0,
        LockedUntil: '',
        LastLoginAt: now,
        UpdatedAt: now
      });
    }

    const token = APP_CONFIG.TOKEN_PREFIX + '_' + createRandomToken_() + createRandomToken_();
    const expiresAt = new Date(now.getTime() + APP_CONFIG.SESSION_HOURS * 60 * 60 * 1000);
    const sessionId = createId_('SES');

    appendObject_('Sessions', {
      SessionID: sessionId,
      TokenHash: hashToken_(token),
      UserID: matchedUser.UserID,
      Role: String(matchedUser.Role).toUpperCase(),
      DeviceID: deviceId,
      CreatedAt: now,
      ExpiresAt: expiresAt,
      LastSeenAt: now,
      RevokedAt: ''
    });

    if (deviceId) touchDevice_(deviceId, matchedUser.UserID);

    CacheService.getScriptCache().remove(attemptCacheKey);
    logAuditSystem_(matchedUser.UserID, matchedUser.FullName, 'LOGIN', 'USER',
      matchedUser.UserID, { deviceId: deviceId });

    return successResponse_('เข้าสู่ระบบสำเร็จ', {
      token: token,
      expiresAt: expiresAt.toISOString(),
      user: publicUser_(matchedUser)
    });
  } finally {
    lock.releaseLock();
  }
}

function recordLoginFailure_(accessCode) {
  const key = 'LOGIN_FAIL_' + hashToken_(accessCode).substring(0, 30);
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(count), APP_CONFIG.LOGIN_LOCK_MINUTES * 60);
  if (count >= APP_CONFIG.LOGIN_MAX_ATTEMPTS) {
    logSystem_('WARN', 'มีการพยายามเข้าสู่ระบบไม่สำเร็จหลายครั้ง',
      'Repeated login failures. Attempt counter key=' + key,
      'apiLogin_');
  }
}

function authenticateRequest_(token) {
  const sessionToken = String(token || '').trim();
  if (!sessionToken) {
    throw createAppError_('กรุณาเข้าสู่ระบบใหม่อีกครั้ง', 'SESSION_REQUIRED');
  }

  const tokenHash = hashToken_(sessionToken);
  const sessionRecord = findRowByKey_('Sessions', 'TokenHash', tokenHash);
  if (!sessionRecord) {
    throw createAppError_('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่', 'SESSION_INVALID');
  }

  const session = sessionRecord.object;
  if (session.RevokedAt) {
    throw createAppError_('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่', 'SESSION_REVOKED');
  }

  const expiresAt = parseDate_(session.ExpiresAt);
  if (!expiresAt || expiresAt.getTime() <= Date.now()) {
    updateRowObject_('Sessions', sessionRecord.rowIndex, { RevokedAt: new Date() });
    throw createAppError_('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่', 'SESSION_EXPIRED');
  }

  const userRecord = findRowByKey_('Users', 'UserID', session.UserID);
  if (!userRecord || String(userRecord.object.Status).toUpperCase() !== 'ACTIVE') {
    throw createAppError_('บัญชีนี้ไม่สามารถใช้งานได้', 'USER_INACTIVE');
  }

  updateRowObject_('Sessions', sessionRecord.rowIndex, { LastSeenAt: new Date() });

  return {
    sessionId: session.SessionID,
    userId: userRecord.object.UserID,
    fullName: userRecord.object.FullName,
    role: String(userRecord.object.Role).toUpperCase(),
    deviceId: session.DeviceID || '',
    user: userRecord.object
  };
}

function apiLogout_(session) {
  const record = findRowByKey_('Sessions', 'SessionID', session.sessionId);
  if (record) {
    updateRowObject_('Sessions', record.rowIndex, { RevokedAt: new Date() });
  }
  logAudit_(session, 'LOGOUT', 'SESSION', session.sessionId, {});
  return successResponse_('ออกจากระบบเรียบร้อยแล้ว', {});
}

function requirePermission_(session, permissionCode) {
  if (session.role === 'ADMIN') return true;
  const permission = findPermission_(session.role, permissionCode);
  if (!permission) {
    throw createAppError_('คุณไม่มีสิทธิ์ดำเนินการนี้', 'PERMISSION_DENIED');
  }
  return true;
}

function findPermission_(role, permissionCode) {
  const records = getSheetObjects_('RolePermissions');
  const item = records.find(function (row) {
    return String(row.Role).toUpperCase() === role &&
      String(row.PermissionCode) === permissionCode;
  });
  return !!(item && toBoolean_(item.IsAllowed));
}

function requireAdmin_(session) {
  if (session.role !== 'ADMIN') {
    throw createAppError_('คุณไม่มีสิทธิ์ดำเนินการนี้', 'ADMIN_REQUIRED');
  }
}

function publicUser_(user) {
  return {
    userId: user.UserID,
    prefix: user.Prefix || '',
    fullName: user.FullName,
    role: String(user.Role).toUpperCase(),
    status: user.Status,
    codePrefix: user.CodePrefix || ''
  };
}

/* ========================================================================== */
/*  5. BOOTSTRAP AND SETTINGS                                                  */
/* ========================================================================== */

function apiGetBootstrap_(session) {
  const allowed = getAllowedPermissions_(session.role);
  const settings = getPublicSettings_();
  const shift = getOpenShiftForUser_(session.userId);

  return successResponse_('โหลดข้อมูลระบบเรียบร้อยแล้ว', {
    appName: APP_CONFIG.APP_NAME,
    version: APP_CONFIG.VERSION,
    timezone: APP_CONFIG.TIMEZONE,
    user: publicUser_(session.user),
    permissions: allowed,
    settings: settings,
    openShift: shift ? publicRecord_(shift) : null,
    navigation: [
      { code: 'payment', label: 'รับชำระเงิน', permission: 'payment.create' },
      { code: 'transactions', label: 'ประวัติรายการ', permission: 'transaction.read.own' },
      { code: 'reports', label: 'รายงาน', permission: 'reports.read' },
      { code: 'staff', label: 'พนักงาน', permission: 'staff.read' },
      { code: 'more', label: 'เพิ่มเติม', permission: null }
    ]
  });
}

function getAllowedPermissions_(role) {
  if (role === 'ADMIN') {
    return DEFAULT_PERMISSIONS
      .filter(function (row) { return row[0] === 'ADMIN'; })
      .map(function (row) { return row[1]; });
  }
  return getSheetObjects_('RolePermissions')
    .filter(function (row) {
      return String(row.Role).toUpperCase() === role && toBoolean_(row.IsAllowed);
    })
    .map(function (row) { return row.PermissionCode; });
}

function getPublicSettings_() {
  const settings = getSettingsMap_();
  return {
    businessName: settings.BUSINESS_NAME || APP_CONFIG.APP_NAME,
    businessAddress: settings.BUSINESS_ADDRESS || '',
    businessPhone: settings.BUSINESS_PHONE || '',
    businessLogoUrl: settings.BUSINESS_LOGO_URL || '',
    qrExpirationMinutes: Number(settings.QR_EXPIRATION_MINUTES || APP_CONFIG.DEFAULT_QR_EXPIRATION_MINUTES),
    soundEnabled: toBoolean_(settings.SOUND_ENABLED),
    receiptFooter: settings.RECEIPT_FOOTER || 'ขอบคุณที่ใช้บริการ'
  };
}

function apiGetSettings_(session) {
  requirePermission_(session, 'settings.read');
  const values = getSheetObjects_('Settings').map(function (row) {
    return {
      key: row.SettingKey,
      value: row.SettingValue,
      description: row.Description || ''
    };
  });
  return successResponse_('โหลดการตั้งค่าเรียบร้อยแล้ว', { settings: values });
}

function apiSaveSettings_(session, data) {
  requirePermission_(session, 'settings.manage');
  const entries = Array.isArray(data.settings) ? data.settings : [];
  if (!entries.length) {
    return errorResponse_('ไม่มีข้อมูลการตั้งค่าที่ต้องการบันทึก', 'SETTINGS_EMPTY');
  }

  const allowedKeys = {};
  DEFAULT_SETTINGS.forEach(function (row) { allowedKeys[row[0]] = true; });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    entries.forEach(function (entry) {
      const key = String(entry.key || '').trim();
      if (!allowedKeys[key]) return;
      const value = String(entry.value === undefined || entry.value === null ? '' : entry.value);
      validateSettingValue_(key, value);
      const record = findRowByKey_('Settings', 'SettingKey', key);
      if (record) {
        updateRowObject_('Settings', record.rowIndex, {
          SettingValue: value,
          UpdatedAt: now
        });
      } else {
        appendObject_('Settings', {
          SettingKey: key,
          SettingValue: value,
          Description: '',
          UpdatedAt: now
        });
      }
    });

    logAudit_(session, 'SETTINGS_UPDATED', 'SETTINGS', 'GENERAL', {
      updatedKeys: entries.map(function (entry) { return entry.key; })
    });
    return successResponse_('บันทึกการตั้งค่าเรียบร้อยแล้ว', {});
  } finally {
    lock.releaseLock();
  }
}

function apiListRolePermissions_(session) {
  requireAdmin_(session);
  const permissions = getSheetObjects_('RolePermissions');
  return successResponse_('โหลดสิทธิ์การใช้งานเรียบร้อยแล้ว', {
    permissions: permissions.map(function (row) {
      return {
        role: row.Role,
        permissionCode: row.PermissionCode,
        isAllowed: toBoolean_(row.IsAllowed),
        updatedAt: toIsoString_(row.UpdatedAt)
      };
    })
  });
}

function apiSaveRolePermissions_(session, data) {
  requireAdmin_(session);
  const entries = Array.isArray(data.permissions) ? data.permissions : [];
  if (!entries.length) return errorResponse_('ไม่มีสิทธิ์ที่ต้องการบันทึก', 'PERMISSION_DATA_EMPTY');

  const allowedRoles = ['ADMIN', 'MANAGER', 'CASHIER'];
  const allowedCodes = {};
  DEFAULT_PERMISSIONS.forEach(function (row) { allowedCodes[row[1]] = true; });
  const now = new Date();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    entries.forEach(function (entry) {
      const role = String(entry.role || '').toUpperCase();
      const permissionCode = String(entry.permissionCode || '');
      if (allowedRoles.indexOf(role) === -1 || !allowedCodes[permissionCode]) return;
      if (role === 'ADMIN') return; // Admin permission is always enabled.

      const record = findRowByPredicate_('RolePermissions', function (row) {
        return String(row.Role).toUpperCase() === role && String(row.PermissionCode) === permissionCode;
      });
      const value = toBoolean_(entry.isAllowed);
      if (record) {
        updateRowObject_('RolePermissions', record.rowIndex, { IsAllowed: value, UpdatedAt: now });
      } else {
        appendObject_('RolePermissions', {
          Role: role,
          PermissionCode: permissionCode,
          IsAllowed: value,
          UpdatedAt: now
        });
      }
    });
    logAudit_(session, 'ROLE_PERMISSIONS_UPDATED', 'ROLE_PERMISSIONS', 'ALL', {
      changedCount: entries.length
    });
    return successResponse_('บันทึกสิทธิ์การใช้งานเรียบร้อยแล้ว', {});
  } finally {
    lock.releaseLock();
  }
}

function apiChangeOwnAccessCode_(session, data) {
  const oldCode = String(data.oldAccessCode || '').trim();
  const newCode = String(data.newAccessCode || '').trim();
  if (!oldCode || !newCode) return errorResponse_('กรุณากรอกรหัสเดิมและรหัสใหม่', 'ACCESS_CODE_REQUIRED');
  if (newCode.length < 8 || newCode.length > 64) {
    return errorResponse_('รหัสใหม่ต้องมีความยาว 8 ถึง 64 ตัวอักษร', 'ACCESS_CODE_LENGTH_INVALID');
  }

  const record = findRowByKey_('Users', 'UserID', session.userId);
  if (!record) return errorResponse_('ไม่พบบัญชีผู้ใช้งาน', 'USER_NOT_FOUND');
  const currentHash = hashAccessCode_(oldCode, String(record.object.AccessCodeSalt || ''));
  if (!constantTimeEquals_(currentHash, String(record.object.AccessCodeHash || ''))) {
    return errorResponse_('รหัสเดิมไม่ถูกต้อง', 'OLD_ACCESS_CODE_INVALID');
  }

  const salt = createRandomToken_();
  updateRowObject_('Users', record.rowIndex, {
    AccessCodeHash: hashAccessCode_(newCode, salt),
    AccessCodeSalt: salt,
    UpdatedAt: new Date(),
    FailedLoginCount: 0,
    LockedUntil: ''
  });
  revokeSessionsForUser_(session.userId);
  logAudit_(session, 'OWN_ACCESS_CODE_CHANGED', 'USER', session.userId, {});
  return successResponse_('เปลี่ยนรหัสเข้าใช้งานเรียบร้อยแล้ว กรุณาเข้าสู่ระบบใหม่', {
    requireLogin: true
  });
}

function validateSettingValue_(key, value) {
  if (key === 'QR_EXPIRATION_MINUTES') {
    const number = Number(value);
    if (!Number.isInteger(number) || number < 1 || number > 60) {
      throw createAppError_('เวลาหมดอายุ QR ต้องอยู่ระหว่าง 1 ถึง 60 นาที', 'INVALID_SETTING');
    }
  }
  if (key === 'IMG_BB_EXPIRATION_SECONDS') {
    const number = Number(value);
    if (number !== 0 && (number < 60 || number > 15552000)) {
      throw createAppError_('ระยะเวลาจัดเก็บรูปภาพไม่ถูกต้อง', 'INVALID_SETTING');
    }
  }
  if (key === 'THUNDER_QR_TYPE') {
    const allowed = ['PROMPTPAY', 'KSHOP', 'MAE_MANEE', 'TUNGNGERN'];
    if (allowed.indexOf(value.toUpperCase()) === -1) {
      throw createAppError_('ประเภท QR Code ไม่ถูกต้อง', 'INVALID_SETTING');
    }
  }
}

function getSettingsMap_() {
  const map = {};
  getSheetObjects_('Settings').forEach(function (row) {
    map[String(row.SettingKey)] = String(row.SettingValue === undefined ? '' : row.SettingValue);
  });
  DEFAULT_SETTINGS.forEach(function (row) {
    if (map[row[0]] === undefined) map[row[0]] = row[1];
  });
  return map;
}

/* ========================================================================== */
/*  6. PAYMENT - THUNDER QR GENERATION                                         */
/* ========================================================================== */

function apiCreatePayment_(session, data) {
  requirePermission_(session, 'payment.create');

  const amount = roundMoney_(data.amount);
  if (!isFinite(amount) || amount <= 0 || amount > 9999999999.99) {
    return errorResponse_('กรุณาระบุจำนวนเงินให้ถูกต้อง', 'INVALID_AMOUNT');
  }

  const settings = getSettingsMap_();
  const openShift = getOpenShiftForUser_(session.userId);
  if (toBoolean_(settings.REQUIRE_OPEN_SHIFT) && !openShift) {
    return errorResponse_('กรุณาเปิดกะทำงานก่อนรับชำระเงิน', 'SHIFT_REQUIRED');
  }

  const deviceId = String(data.deviceId || session.deviceId || '').trim();
  const transactionId = createTransactionId_();
  const qrConfig = buildThunderQrRequest_(settings, amount, transactionId);
  const qrResult = callThunderQr_(qrConfig, transactionId);

  const createdAt = new Date();
  const expirationMinutes = Number(settings.QR_EXPIRATION_MINUTES || APP_CONFIG.DEFAULT_QR_EXPIRATION_MINUTES);
  const expiresAt = new Date(createdAt.getTime() + expirationMinutes * 60 * 1000);
  const qrType = String(qrConfig.type || 'PROMPTPAY');

  const transaction = {
    TransactionID: transactionId,
    Amount: amount,
    Currency: APP_CONFIG.DEFAULT_CURRENCY,
    Status: 'PENDING',
    Note: String(data.note || '').trim().substring(0, 250),
    CashierID: session.userId,
    CashierName: session.fullName,
    DeviceID: deviceId,
    ShiftID: openShift ? openShift.ShiftID : '',
    QRType: qrType,
    QRPayload: qrResult.payload,
    QRMime: qrResult.mime,
    QRCreatedAt: createdAt,
    QRExpiresAt: expiresAt,
    PaidAt: '',
    SlipImageURL: '',
    SlipViewerURL: '',
    ImgBBImageID: '',
    ThunderTransRef: '',
    LastSlipCheckID: '',
    CreatedAt: createdAt,
    UpdatedAt: createdAt,
    CancelledAt: '',
    CancelReason: ''
  };

  appendObject_('Transactions', transaction);
  if (deviceId) touchDevice_(deviceId, session.userId);
  logAudit_(session, 'PAYMENT_CREATED', 'TRANSACTION', transactionId, {
    amount: amount,
    qrType: qrType,
    deviceId: deviceId
  });
  createNotification_('PAYMENT_PENDING', 'มีรายการรอชำระเงิน',
    'รายการ ' + transactionId + ' รอการชำระเงิน', session.userId, transactionId);

  return successResponse_('สร้าง QR Code เรียบร้อยแล้ว', {
    transaction: publicRecord_(transaction),
    qr: {
      image: qrResult.image,
      mime: qrResult.mime,
      payload: qrResult.payload
    },
    expiresAt: expiresAt.toISOString(),
    expirationSeconds: expirationMinutes * 60
  });
}

function buildThunderQrRequest_(settings, amount, transactionId) {
  const type = String(settings.THUNDER_QR_TYPE || 'PROMPTPAY').toUpperCase();
  const request = { type: type, amount: amount };
  const ref1 = transactionId.replace(/[^A-Za-z0-9]/g, '').substring(0, 20);

  if (type === 'PROMPTPAY') {
    const identifierType = String(settings.PROMPTPAY_IDENTIFIER_TYPE || 'msisdn');
    const identifier = String(settings.PROMPTPAY_IDENTIFIER || '').trim();
    if (!identifier) {
      throw createAppError_('กรุณาตั้งค่าข้อมูล PromptPay ก่อนสร้าง QR Code', 'PROMPTPAY_NOT_CONFIGURED');
    }
    if (['msisdn', 'natId', 'eWalletId', 'billerId'].indexOf(identifierType) === -1) {
      throw createAppError_('ประเภทบัญชี PromptPay ไม่ถูกต้อง', 'PROMPTPAY_CONFIG_INVALID');
    }
    request[identifierType] = identifier;
    if (identifierType === 'billerId') {
      const billerRef1 = String(settings.PROMPTPAY_BILLER_REF1 || ref1).trim();
      if (!billerRef1) {
        throw createAppError_('กรุณาตั้งค่า Reference 1 สำหรับ Biller ID', 'BILLER_REF_REQUIRED');
      }
      request.ref1 = billerRef1.substring(0, 20);
      const ref2 = String(settings.PROMPTPAY_BILLER_REF2 || '').trim();
      if (ref2) request.ref2 = ref2.substring(0, 20);
    }
    return request;
  }

  if (type === 'KSHOP' || type === 'MAE_MANEE') {
    request.ref1 = ref1;
    return request;
  }

  if (type === 'TUNGNGERN') {
    request.ref1 = ref1;
    const merchantName = String(settings.THUNDER_MERCHANT_NAME || '').trim();
    if (merchantName) request.merchantName = merchantName.substring(0, 25);
    return request;
  }

  throw createAppError_('ประเภท QR Code ยังไม่รองรับ', 'QR_TYPE_UNSUPPORTED');
}

function callThunderQr_(requestBody, transactionId) {
  const apiKey = getScriptProperty_('THUNDER_API_KEY');
  if (!apiKey) {
    throw createAppError_('ยังไม่ได้ตั้งค่าระบบรับชำระเงิน', 'THUNDER_NOT_CONFIGURED');
  }

  const started = Date.now();
  let response;
  let parsed;
  try {
    response = UrlFetchApp.fetch(APP_CONFIG.THUNDER_QR_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    parsed = parseJsonResponse_(response.getContentText());
  } catch (error) {
    recordApiUsage_('THUNDER', APP_CONFIG.THUNDER_QR_URL, transactionId, 0,
      Date.now() - started, false);
    throw createAppError_('ไม่สามารถสร้าง QR Code ได้', 'THUNDER_QR_REQUEST_FAILED', error);
  }

  const statusCode = response.getResponseCode();
  const success = statusCode >= 200 && statusCode < 300 &&
    Number(parsed.status || statusCode) === 200 && parsed.data && parsed.data.image;
  recordApiUsage_('THUNDER', APP_CONFIG.THUNDER_QR_URL, transactionId,
    statusCode, Date.now() - started, !!success);

  if (!success) {
    logSystem_('ERROR', 'ไม่สามารถสร้าง QR Code ได้',
      'Thunder QR response: ' + safeStringify_(parsed), 'callThunderQr_');
    throw createAppError_('ไม่สามารถสร้าง QR Code ได้ กรุณาลองใหม่อีกครั้ง', 'THUNDER_QR_REJECTED');
  }

  return {
    image: parsed.data.image,
    mime: parsed.data.mime || 'image/png',
    payload: parsed.data.payload || ''
  };
}

function apiCancelPayment_(session, data) {
  requirePermission_(session, 'payment.cancel');
  const transactionId = String(data.transactionId || '').trim();
  if (!transactionId) return errorResponse_('ไม่พบรายการที่ต้องการยกเลิก', 'TRANSACTION_REQUIRED');

  const record = findRowByKey_('Transactions', 'TransactionID', transactionId);
  if (!record) return errorResponse_('ไม่พบรายการที่ต้องการ', 'TRANSACTION_NOT_FOUND');
  const transaction = record.object;

  if (String(transaction.Status) !== 'PENDING') {
    return errorResponse_('รายการนี้ไม่สามารถยกเลิกได้', 'TRANSACTION_NOT_PENDING');
  }

  if (session.role !== 'ADMIN' && transaction.CashierID !== session.userId) {
    return errorResponse_('คุณไม่มีสิทธิ์ยกเลิกรายการนี้', 'PERMISSION_DENIED');
  }

  const now = new Date();
  updateRowObject_('Transactions', record.rowIndex, {
    Status: 'CANCELLED',
    CancelledAt: now,
    CancelReason: String(data.reason || 'ยกเลิกโดยผู้ใช้งาน').trim().substring(0, 250),
    UpdatedAt: now
  });

  logAudit_(session, 'PAYMENT_CANCELLED', 'TRANSACTION', transactionId, {
    reason: String(data.reason || '').trim()
  });
  return successResponse_('ยกเลิกรายการเรียบร้อยแล้ว', { transactionId: transactionId });
}

/* ========================================================================== */
/*  7. PAYMENT - SLIP UPLOAD AND THUNDER VERIFICATION                          */
/* ========================================================================== */

function apiVerifySlip_(session, data) {
  requirePermission_(session, 'payment.create');

  const transactionId = String(data.transactionId || '').trim();
  const imageData = normalizeBase64Image_(data.imageBase64 || data.image || '', data.mimeType || '');
  if (!transactionId) return errorResponse_('ไม่พบรายการชำระเงิน', 'TRANSACTION_REQUIRED');
  if (!imageData.base64) return errorResponse_('กรุณาถ่ายภาพสลิปก่อนส่งตรวจสอบ', 'SLIP_IMAGE_REQUIRED');
  if (imageData.bytes > APP_CONFIG.MAX_SLIP_BYTES) {
    return errorResponse_('ขนาดภาพสลิปต้องไม่เกิน 4 MB', 'SLIP_TOO_LARGE');
  }
  if (APP_CONFIG.SUPPORTED_IMAGE_MIMES.indexOf(imageData.mimeType) === -1) {
    return errorResponse_('รองรับเฉพาะภาพ JPG, PNG และ WebP', 'SLIP_FORMAT_UNSUPPORTED');
  }

  const transactionRecord = findRowByKey_('Transactions', 'TransactionID', transactionId);
  if (!transactionRecord) return errorResponse_('ไม่พบรายการชำระเงิน', 'TRANSACTION_NOT_FOUND');
  const transaction = transactionRecord.object;

  if (String(transaction.Status) === 'PAID') {
    return errorResponse_('รายการนี้ชำระเงินสำเร็จแล้ว', 'TRANSACTION_ALREADY_PAID');
  }
  if (String(transaction.Status) !== 'PENDING') {
    return errorResponse_('รายการนี้ไม่อยู่ในสถานะรอชำระเงิน', 'TRANSACTION_NOT_PENDING');
  }
  if (session.role !== 'ADMIN' && transaction.CashierID !== session.userId) {
    return errorResponse_('คุณไม่มีสิทธิ์ตรวจสอบรายการนี้', 'PERMISSION_DENIED');
  }

  const now = new Date();
  const qrExpiresAt = parseDate_(transaction.QRExpiresAt);
  if (!qrExpiresAt || qrExpiresAt.getTime() <= now.getTime()) {
    updateRowObject_('Transactions', transactionRecord.rowIndex, {
      Status: 'EXPIRED', UpdatedAt: now
    });
    return errorResponse_('QR Code หมดอายุแล้ว กรุณาสร้างรายการใหม่', 'QR_EXPIRED');
  }

  const attemptNo = getNextSlipAttempt_(transactionId);
  const mediaName = 'SLIP_' + transactionId + '_' + attemptNo;
  const media = uploadToImgBb_(imageData.base64, imageData.mimeType, mediaName, 'SLIP', transactionId, session);

  const settings = getSettingsMap_();
  const matchAccount = toBoolean_(settings.THUNDER_MATCH_ACCOUNT);
  const checkDuplicate = toBoolean_(settings.THUNDER_CHECK_DUPLICATE);
  const thunderResult = callThunderVerifyByUrl_(
    media.imageUrl,
    Number(transaction.Amount),
    transactionId,
    matchAccount,
    checkDuplicate
  );

  const decision = evaluateThunderSlip_(thunderResult, Number(transaction.Amount), matchAccount);
  const transRef = decision.transRef;
  let localDuplicate = false;
  if (transRef) {
    localDuplicate = hasSuccessfulSlipReference_(transRef);
  }
  if (localDuplicate) {
    decision.passed = false;
    decision.reason = 'สลิปนี้เคยถูกใช้กับรายการอื่นแล้ว';
    decision.isDuplicate = true;
  }

  const checkId = createId_('CHK');
  const checkedAt = new Date();
  const dataResult = thunderResult && thunderResult.data ? thunderResult.data : {};
  const rawSlip = dataResult.rawSlip || {};
  const sender = rawSlip.sender || {};
  const receiver = rawSlip.receiver || {};
  const senderAccount = sender.account || {};
  const receiverAccount = receiver.account || {};
  const senderBank = sender.bank || {};
  const receiverBank = receiver.bank || {};
  const senderName = senderAccount.name || {};
  const receiverName = receiverAccount.name || {};
  const receiverBankAccount = receiverAccount.bank || {};

  const slipCheck = {
    CheckID: checkId,
    TransactionID: transactionId,
    AttemptNo: attemptNo,
    ImageURL: media.imageUrl,
    ViewerURL: media.viewerUrl,
    ImgBBImageID: media.imageId,
    ThunderSuccess: !!(thunderResult && thunderResult.success),
    IsDuplicate: !!decision.isDuplicate,
    IsAmountMatched: !!decision.isAmountMatched,
    IsAccountMatched: !!decision.isAccountMatched,
    AmountInSlip: decision.amountInSlip,
    AmountInOrder: Number(transaction.Amount),
    TransRef: transRef,
    SenderBank: senderBank.name || senderBank.short || '',
    SenderName: senderName.th || senderName.en || '',
    ReceiverBank: receiverBank.name || receiverBank.short || '',
    ReceiverName: receiverName.th || receiverName.en || '',
    ReceiverAccount: receiverBankAccount.account || '',
    SlipDate: rawSlip.date || '',
    Result: decision.passed ? 'PASSED' : (decision.apiError ? 'API_ERROR' : 'FAILED'),
    FailureReason: decision.reason || '',
    RawResponse: safeStringify_(thunderResult),
    CheckedBy: session.userId,
    CreatedAt: checkedAt
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const latestRecord = findRowByKey_('Transactions', 'TransactionID', transactionId);
    if (!latestRecord) {
      return errorResponse_('ไม่พบรายการชำระเงิน', 'TRANSACTION_NOT_FOUND');
    }
    if (String(latestRecord.object.Status) === 'PAID') {
      return errorResponse_('รายการนี้ชำระเงินสำเร็จแล้ว', 'TRANSACTION_ALREADY_PAID');
    }
    if (String(latestRecord.object.Status) !== 'PENDING') {
      return errorResponse_('รายการนี้ไม่อยู่ในสถานะรอชำระเงิน', 'TRANSACTION_NOT_PENDING');
    }

    if (decision.passed && transRef && hasSuccessfulSlipReference_(transRef)) {
      decision.passed = false;
      decision.isDuplicate = true;
      decision.reason = 'สลิปนี้เคยถูกใช้กับรายการอื่นแล้ว';
      slipCheck.IsDuplicate = true;
      slipCheck.Result = 'FAILED';
      slipCheck.FailureReason = decision.reason;
    }

    appendObject_('SlipChecks', slipCheck);

    if (decision.passed) {
      updateRowObject_('Transactions', latestRecord.rowIndex, {
        Status: 'PAID',
        PaidAt: checkedAt,
        SlipImageURL: media.imageUrl,
        SlipViewerURL: media.viewerUrl,
        ImgBBImageID: media.imageId,
        ThunderTransRef: transRef,
        LastSlipCheckID: checkId,
        UpdatedAt: checkedAt
      });

      if (latestRecord.object.ShiftID) {
        updateShiftTotals_(latestRecord.object.ShiftID);
      }

      logAudit_(session, 'PAYMENT_VERIFIED', 'TRANSACTION', transactionId, {
        amount: transaction.Amount,
        thunderTransRef: transRef,
        checkId: checkId
      });
      createNotification_('PAYMENT_SUCCESS', 'รับชำระเงินสำเร็จ',
        'รายการ ' + transactionId + ' ชำระเงินสำเร็จ', transaction.CashierID, transactionId);

      return successResponse_('ตรวจสอบการชำระเงินสำเร็จ', {
        verified: true,
        canProceed: true,
        transactionId: transactionId,
        amount: Number(transaction.Amount),
        slipCheck: publicSlipCheck_(slipCheck)
      });
    }

    updateRowObject_('Transactions', latestRecord.rowIndex, {
      SlipImageURL: media.imageUrl,
      SlipViewerURL: media.viewerUrl,
      ImgBBImageID: media.imageId,
      LastSlipCheckID: checkId,
      UpdatedAt: checkedAt
    });
    logAudit_(session, 'SLIP_VERIFICATION_FAILED', 'TRANSACTION', transactionId, {
      attemptNo: attemptNo,
      reason: decision.reason,
      checkId: checkId
    });

    return successResponse_('ตรวจสอบสลิปไม่ผ่าน กรุณาถ่ายใหม่', {
      verified: false,
      canRetry: true,
      canProceed: false,
      reason: decision.reason || 'ไม่สามารถยืนยันการชำระเงินได้',
      attemptNo: attemptNo,
      slipCheck: publicSlipCheck_(slipCheck)
    });
  } finally {
    lock.releaseLock();
  }
}

function callThunderVerifyByUrl_(imageUrl, expectedAmount, transactionId, matchAccount, checkDuplicate) {
  const apiKey = getScriptProperty_('THUNDER_API_KEY');
  if (!apiKey) {
    return { success: false, apiError: true, error: { message: 'THUNDER_API_KEY_NOT_CONFIGURED' } };
  }

  const payload = {
    url: imageUrl,
    remark: transactionId.substring(0, 255),
    matchAccount: !!matchAccount,
    matchAmount: Number(expectedAmount),
    checkDuplicate: !!checkDuplicate
  };

  const started = Date.now();
  let response;
  let parsed;
  try {
    response = UrlFetchApp.fetch(APP_CONFIG.THUNDER_VERIFY_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    parsed = parseJsonResponse_(response.getContentText());
  } catch (error) {
    recordApiUsage_('THUNDER', APP_CONFIG.THUNDER_VERIFY_URL, transactionId, 0,
      Date.now() - started, false);
    logSystem_('ERROR', 'ไม่สามารถตรวจสอบสลิปได้', error.stack || String(error), 'callThunderVerifyByUrl_');
    return { success: false, apiError: true, error: { message: 'THUNDER_REQUEST_FAILED' } };
  }

  const statusCode = response.getResponseCode();
  const success = statusCode >= 200 && statusCode < 300 && parsed.success === true && parsed.data;
  recordApiUsage_('THUNDER', APP_CONFIG.THUNDER_VERIFY_URL, transactionId,
    statusCode, Date.now() - started, !!success);

  if (!success) {
    logSystem_('WARN', 'ตรวจสอบสลิปไม่สำเร็จ', 'Thunder response: ' + safeStringify_(parsed), 'callThunderVerifyByUrl_');
    return {
      success: false,
      apiError: true,
      error: parsed.error || { message: parsed.message || 'THUNDER_VERIFY_FAILED' },
      httpStatus: statusCode
    };
  }

  return parsed;
}

function evaluateThunderSlip_(result, expectedAmount, requireAccountMatch) {
  const output = {
    passed: false,
    apiError: false,
    reason: '',
    isDuplicate: false,
    isAmountMatched: false,
    isAccountMatched: false,
    amountInSlip: 0,
    transRef: ''
  };

  if (!result || result.success !== true || !result.data) {
    output.apiError = !!(result && result.apiError);
    output.reason = output.apiError
      ? 'ระบบตรวจสอบไม่พร้อมใช้งาน กรุณาลองใหม่อีกครั้ง'
      : 'ไม่สามารถตรวจสอบข้อมูลจากสลิปได้';
    return output;
  }

  const data = result.data;
  output.isDuplicate = data.isDuplicate === true;
  output.amountInSlip = Number(data.amountInSlip || 0);
  output.isAmountMatched = typeof data.isAmountMatched === 'boolean'
    ? data.isAmountMatched
    : Math.abs(output.amountInSlip - Number(expectedAmount)) < 0.005;
  output.isAccountMatched = !!data.matchedAccount;
  output.transRef = String((data.rawSlip && data.rawSlip.transRef) || '').trim();

  if (output.isDuplicate) {
    output.reason = 'สลิปนี้ถูกใช้หรือถูกตรวจสอบแล้ว';
    return output;
  }
  if (!output.isAmountMatched) {
    output.reason = 'ยอดเงินในสลิปไม่ตรงกับยอดที่ต้องชำระ';
    return output;
  }
  if (requireAccountMatch && !output.isAccountMatched) {
    output.reason = 'ไม่พบข้อมูลบัญชีผู้รับที่ตรงกับร้านค้า';
    return output;
  }
  if (!output.transRef) {
    output.reason = 'ไม่พบเลขอ้างอิงธุรกรรมในสลิป';
    return output;
  }

  output.passed = true;
  output.reason = '';
  return output;
}

function getNextSlipAttempt_(transactionId) {
  const rows = getSheetObjects_('SlipChecks');
  return rows.filter(function (row) {
    return String(row.TransactionID) === transactionId;
  }).length + 1;
}

function hasSuccessfulSlipReference_(transRef) {
  if (!transRef) return false;
  const rows = getSheetObjects_('SlipChecks');
  return rows.some(function (row) {
    return String(row.TransRef) === String(transRef) && String(row.Result) === 'PASSED';
  });
}

function publicSlipCheck_(check) {
  return {
    checkId: check.CheckID,
    transactionId: check.TransactionID,
    attemptNo: check.AttemptNo,
    imageUrl: check.ImageURL,
    result: check.Result,
    isDuplicate: toBoolean_(check.IsDuplicate),
    isAmountMatched: toBoolean_(check.IsAmountMatched),
    isAccountMatched: toBoolean_(check.IsAccountMatched),
    amountInSlip: Number(check.AmountInSlip || 0),
    reason: check.FailureReason || ''
  };
}

/* ========================================================================== */
/*  8. IMGBB IMAGE STORAGE                                                     */
/* ========================================================================== */

function apiUploadImage_(session, data) {
  requirePermission_(session, 'media.upload');
  const imageData = normalizeBase64Image_(data.imageBase64 || data.image || '', data.mimeType || '');
  if (!imageData.base64) return errorResponse_('กรุณาเลือกภาพที่ต้องการอัปโหลด', 'IMAGE_REQUIRED');
  if (imageData.bytes > APP_CONFIG.MAX_GENERAL_IMAGE_BYTES) {
    return errorResponse_('ขนาดรูปภาพต้องไม่เกิน 4 MB', 'IMAGE_TOO_LARGE');
  }
  if (APP_CONFIG.SUPPORTED_IMAGE_MIMES.indexOf(imageData.mimeType) === -1) {
    return errorResponse_('รองรับเฉพาะภาพ JPG, PNG และ WebP', 'IMAGE_FORMAT_UNSUPPORTED');
  }

  const imageType = String(data.imageType || 'GENERAL').toUpperCase();
  const relatedId = String(data.relatedId || '').trim();
  const name = String(data.fileName || imageType + '_' + Date.now()).replace(/[^A-Za-z0-9_-]/g, '_');
  const result = uploadToImgBb_(imageData.base64, imageData.mimeType, name, imageType, relatedId, session);
  logAudit_(session, 'IMAGE_UPLOADED', 'MEDIA', result.mediaId, {
    imageType: imageType,
    relatedId: relatedId
  });
  return successResponse_('อัปโหลดรูปภาพเรียบร้อยแล้ว', result);
}

function uploadToImgBb_(base64, mimeType, fileName, imageType, relatedId, session) {
  const apiKey = getScriptProperty_('IMGBB_API_KEY');
  if (!apiKey) {
    throw createAppError_('ยังไม่ได้ตั้งค่าระบบจัดเก็บรูปภาพ', 'IMGBB_NOT_CONFIGURED');
  }

  const settings = getSettingsMap_();
  const expiration = Number(settings.IMG_BB_EXPIRATION_SECONDS || 0);
  const query = ['key=' + encodeURIComponent(apiKey)];
  if (expiration > 0) {
    if (expiration < 60 || expiration > 15552000) {
      throw createAppError_('การตั้งค่าเวลาจัดเก็บรูปภาพไม่ถูกต้อง', 'IMGBB_EXPIRATION_INVALID');
    }
    query.push('expiration=' + encodeURIComponent(String(expiration)));
  }

  const endpoint = APP_CONFIG.IMGBB_UPLOAD_URL + '?' + query.join('&');
  const started = Date.now();
  let response;
  let parsed;
  try {
    response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      payload: {
        image: base64,
        name: String(fileName || 'edc-image').substring(0, 100)
      },
      muteHttpExceptions: true
    });
    parsed = parseJsonResponse_(response.getContentText());
  } catch (error) {
    recordApiUsage_('IMGBB', APP_CONFIG.IMGBB_UPLOAD_URL, relatedId, 0,
      Date.now() - started, false);
    logSystem_('ERROR', 'ไม่สามารถอัปโหลดรูปภาพได้', error.stack || String(error), 'uploadToImgBb_');
    throw createAppError_('ไม่สามารถอัปโหลดรูปภาพได้ กรุณาลองใหม่', 'IMGBB_UPLOAD_FAILED', error);
  }

  const statusCode = response.getResponseCode();
  const success = statusCode >= 200 && statusCode < 300 && parsed.success === true && parsed.data;
  recordApiUsage_('IMGBB', APP_CONFIG.IMGBB_UPLOAD_URL, relatedId, statusCode,
    Date.now() - started, !!success);

  if (!success) {
    logSystem_('ERROR', 'ไม่สามารถอัปโหลดรูปภาพได้', 'ImgBB response: ' + safeStringify_(parsed), 'uploadToImgBb_');
    throw createAppError_('ไม่สามารถอัปโหลดรูปภาพได้ กรุณาลองใหม่', 'IMGBB_UPLOAD_REJECTED');
  }

  const mediaId = createId_('MED');
  const uploadedAt = new Date();
  const media = {
    MediaID: mediaId,
    ImageType: imageType,
    RelatedID: relatedId,
    ImageURL: parsed.data.url || (parsed.data.image && parsed.data.image.url) || '',
    ViewerURL: parsed.data.url_viewer || '',
    ThumbnailURL: parsed.data.thumb && parsed.data.thumb.url ? parsed.data.thumb.url : '',
    ImgBBImageID: parsed.data.id || '',
    DeleteURL: parsed.data.delete_url || '',
    UploadedBy: session ? session.userId : 'SYSTEM',
    UploadedAt: uploadedAt,
    UploadStatus: 'SUCCESS'
  };

  if (!media.ImageURL) {
    throw createAppError_('ไม่สามารถรับลิงก์รูปภาพจากบริการจัดเก็บได้', 'IMGBB_URL_MISSING');
  }

  appendObject_('MediaFiles', media);
  return {
    mediaId: mediaId,
    imageUrl: media.ImageURL,
    viewerUrl: media.ViewerURL,
    thumbnailUrl: media.ThumbnailURL,
    imageId: media.ImgBBImageID,
    imageType: imageType,
    relatedId: relatedId
  };
}

function normalizeBase64Image_(input, mimeType) {
  let value = String(input || '').trim();
  let detectedMime = String(mimeType || '').toLowerCase().trim();

  if (value.indexOf('data:') === 0) {
    const match = value.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/s);
    if (!match) {
      throw createAppError_('รูปภาพมีรูปแบบไม่ถูกต้อง', 'IMAGE_DATA_INVALID');
    }
    detectedMime = match[1].toLowerCase();
    value = match[2];
  }

  value = value.replace(/\s/g, '');
  if (!detectedMime) detectedMime = 'image/jpeg';
  const bytes = Math.floor(value.length * 3 / 4);
  if (!value || !/^[A-Za-z0-9+/]*={0,2}$/.test(value)) {
    throw createAppError_('ข้อมูลรูปภาพไม่ถูกต้อง', 'IMAGE_BASE64_INVALID');
  }

  return {
    base64: value,
    mimeType: detectedMime,
    bytes: bytes
  };
}

function apiListMediaFiles_(session, data) {
  requirePermission_(session, 'settings.read');
  const limit = clampNumber_(data.limit, 1, 500, 100);
  const imageType = String(data.imageType || '').toUpperCase();
  let records = getSheetObjects_('MediaFiles');
  if (imageType) {
    records = records.filter(function (row) { return String(row.ImageType).toUpperCase() === imageType; });
  }
  records.sort(function (a, b) { return dateValue_(b.UploadedAt) - dateValue_(a.UploadedAt); });
  return successResponse_('โหลดรายการรูปภาพเรียบร้อยแล้ว', {
    media: records.slice(0, limit).map(function (row) {
      return {
        mediaId: row.MediaID,
        imageType: row.ImageType,
        relatedId: row.RelatedID,
        imageUrl: row.ImageURL,
        viewerUrl: row.ViewerURL,
        thumbnailUrl: row.ThumbnailURL,
        uploadedBy: row.UploadedBy,
        uploadedAt: toIsoString_(row.UploadedAt),
        uploadStatus: row.UploadStatus
      };
    })
  });
}

/* ========================================================================== */
/*  9. STAFF MANAGEMENT                                                       */
/* ========================================================================== */

function apiListStaff_(session, data) {
  requirePermission_(session, 'staff.read');
  const statusFilter = String(data.status || '').toUpperCase();
  const query = String(data.query || '').toLowerCase().trim();
  let users = getSheetObjects_('Users').filter(function (user) {
    return String(user.Role).toUpperCase() !== 'ADMIN';
  });

  if (statusFilter) {
    users = users.filter(function (user) { return String(user.Status).toUpperCase() === statusFilter; });
  }
  if (query) {
    users = users.filter(function (user) {
      return String(user.FullName).toLowerCase().indexOf(query) !== -1 ||
        String(user.UserID).toLowerCase().indexOf(query) !== -1 ||
        String(user.CodePrefix).toLowerCase().indexOf(query) !== -1;
    });
  }

  users.sort(function (a, b) { return String(a.FullName).localeCompare(String(b.FullName), 'th'); });
  return successResponse_('โหลดข้อมูลพนักงานเรียบร้อยแล้ว', {
    staff: users.map(function (user) {
      const safe = publicUser_(user);
      safe.createdAt = toIsoString_(user.CreatedAt);
      safe.lastLoginAt = toIsoString_(user.LastLoginAt);
      return safe;
    })
  });
}

function apiCreateStaff_(session, data) {
  requirePermission_(session, 'staff.manage');
  const prefix = String(data.prefix || '').trim();
  const fullName = String(data.fullName || '').trim();
  const codePrefix = String(data.codePrefix || 'TN').trim().toUpperCase();
  const role = String(data.role || 'CASHIER').toUpperCase();

  if (['นาย', 'นาง', 'นางสาว'].indexOf(prefix) === -1) {
    return errorResponse_('กรุณาเลือกคำนำหน้าให้ถูกต้อง', 'STAFF_PREFIX_INVALID');
  }
  if (!fullName || fullName.length > 120) {
    return errorResponse_('กรุณากรอกชื่อ-สกุลให้ถูกต้อง', 'STAFF_NAME_INVALID');
  }
  if (!/^[A-Z]{1,6}$/.test(codePrefix)) {
    return errorResponse_('ตัวขึ้นต้นรหัสต้องเป็นตัวอักษรอังกฤษ 1-6 ตัว', 'STAFF_CODE_PREFIX_INVALID');
  }
  if (['CASHIER', 'MANAGER'].indexOf(role) === -1) {
    return errorResponse_('ตำแหน่งพนักงานไม่ถูกต้อง', 'STAFF_ROLE_INVALID');
  }

  const newCode = generateUniqueStaffCode_(codePrefix);
  const salt = createRandomToken_();
  const userId = createId_('USR');
  const now = new Date();

  appendObject_('Users', {
    UserID: userId,
    Prefix: prefix,
    FullName: fullName,
    AccessCodeHash: hashAccessCode_(newCode, salt),
    AccessCodeSalt: salt,
    Role: role,
    Status: 'ACTIVE',
    CodePrefix: codePrefix,
    CreatedAt: now,
    UpdatedAt: now,
    LastLoginAt: '',
    FailedLoginCount: 0,
    LockedUntil: ''
  });

  logAudit_(session, 'STAFF_CREATED', 'USER', userId, {
    prefix: prefix,
    fullName: fullName,
    role: role,
    codePrefix: codePrefix
  });

  return successResponse_('เพิ่มพนักงานเรียบร้อยแล้ว กรุณาคัดลอกรหัสเข้าใช้งาน', {
    staff: {
      userId: userId,
      prefix: prefix,
      fullName: fullName,
      role: role,
      status: 'ACTIVE'
    },
    accessCode: newCode,
    showCodeOnce: true
  });
}

function generateUniqueStaffCode_(prefix) {
  for (let attempt = 0; attempt < 100; attempt += 1) {
    const randomDigits = generateRandomDigits_(8);
    const code = prefix + randomDigits;
    const collision = getSheetObjects_('Users').some(function (user) {
      const salt = String(user.AccessCodeSalt || '');
      const hash = hashAccessCode_(code, salt);
      return constantTimeEquals_(hash, String(user.AccessCodeHash || ''));
    });
    if (!collision) return code;
  }
  throw createAppError_('ไม่สามารถสร้างรหัสพนักงานได้ กรุณาลองใหม่', 'STAFF_CODE_GENERATION_FAILED');
}

function apiUpdateStaff_(session, data) {
  requirePermission_(session, 'staff.manage');
  const userId = String(data.userId || '').trim();
  if (!userId) return errorResponse_('ไม่พบพนักงานที่ต้องการแก้ไข', 'USER_ID_REQUIRED');

  const record = findRowByKey_('Users', 'UserID', userId);
  if (!record || String(record.object.Role).toUpperCase() === 'ADMIN') {
    return errorResponse_('ไม่พบพนักงานที่ต้องการแก้ไข', 'STAFF_NOT_FOUND');
  }

  const updates = {};
  if (data.prefix !== undefined) {
    const prefix = String(data.prefix).trim();
    if (['นาย', 'นาง', 'นางสาว'].indexOf(prefix) === -1) {
      return errorResponse_('คำนำหน้าไม่ถูกต้อง', 'STAFF_PREFIX_INVALID');
    }
    updates.Prefix = prefix;
  }
  if (data.fullName !== undefined) {
    const fullName = String(data.fullName).trim();
    if (!fullName || fullName.length > 120) return errorResponse_('ชื่อ-สกุลไม่ถูกต้อง', 'STAFF_NAME_INVALID');
    updates.FullName = fullName;
  }
  if (data.role !== undefined) {
    const role = String(data.role).toUpperCase();
    if (['CASHIER', 'MANAGER'].indexOf(role) === -1) return errorResponse_('ตำแหน่งไม่ถูกต้อง', 'STAFF_ROLE_INVALID');
    updates.Role = role;
  }
  updates.UpdatedAt = new Date();

  updateRowObject_('Users', record.rowIndex, updates);
  logAudit_(session, 'STAFF_UPDATED', 'USER', userId, updates);
  return successResponse_('แก้ไขข้อมูลพนักงานเรียบร้อยแล้ว', {});
}

function apiResetStaffCode_(session, data) {
  requirePermission_(session, 'staff.manage');
  const userId = String(data.userId || '').trim();
  const record = findRowByKey_('Users', 'UserID', userId);
  if (!record || String(record.object.Role).toUpperCase() === 'ADMIN') {
    return errorResponse_('ไม่พบพนักงานที่ต้องการเปลี่ยนรหัส', 'STAFF_NOT_FOUND');
  }

  const prefix = String(data.codePrefix || record.object.CodePrefix || 'TN').trim().toUpperCase();
  if (!/^[A-Z]{1,6}$/.test(prefix)) {
    return errorResponse_('ตัวขึ้นต้นรหัสไม่ถูกต้อง', 'STAFF_CODE_PREFIX_INVALID');
  }

  const newCode = generateUniqueStaffCode_(prefix);
  const salt = createRandomToken_();
  updateRowObject_('Users', record.rowIndex, {
    AccessCodeHash: hashAccessCode_(newCode, salt),
    AccessCodeSalt: salt,
    CodePrefix: prefix,
    UpdatedAt: new Date(),
    FailedLoginCount: 0,
    LockedUntil: ''
  });
  revokeSessionsForUser_(userId);
  logAudit_(session, 'STAFF_CODE_RESET', 'USER', userId, { codePrefix: prefix });

  return successResponse_('สร้างรหัสเข้าใช้งานใหม่เรียบร้อยแล้ว', {
    userId: userId,
    accessCode: newCode,
    showCodeOnce: true
  });
}

function apiSetStaffStatus_(session, data) {
  requirePermission_(session, 'staff.manage');
  const userId = String(data.userId || '').trim();
  const status = String(data.status || '').toUpperCase();
  if (['ACTIVE', 'INACTIVE'].indexOf(status) === -1) {
    return errorResponse_('สถานะพนักงานไม่ถูกต้อง', 'STAFF_STATUS_INVALID');
  }
  const record = findRowByKey_('Users', 'UserID', userId);
  if (!record || String(record.object.Role).toUpperCase() === 'ADMIN') {
    return errorResponse_('ไม่พบพนักงานที่ต้องการ', 'STAFF_NOT_FOUND');
  }

  updateRowObject_('Users', record.rowIndex, { Status: status, UpdatedAt: new Date() });
  if (status === 'INACTIVE') revokeSessionsForUser_(userId);
  logAudit_(session, 'STAFF_STATUS_CHANGED', 'USER', userId, { status: status });
  return successResponse_('เปลี่ยนสถานะพนักงานเรียบร้อยแล้ว', { userId: userId, status: status });
}

function revokeSessionsForUser_(userId) {
  const rows = getSheetRows_('Sessions');
  const now = new Date();
  rows.records.forEach(function (record) {
    if (String(record.object.UserID) === String(userId) && !record.object.RevokedAt) {
      updateRowObject_('Sessions', record.rowIndex, { RevokedAt: now });
    }
  });
}

/* ========================================================================== */
/*  10. TRANSACTION HISTORY                                                    */
/* ========================================================================== */

function apiListTransactions_(session, data) {
  const historySettings = getSettingsMap_();
  const canReadAll = session.role === 'ADMIN' || findPermission_(session.role, 'transaction.read.all');
  const canReadOwn = session.role === 'ADMIN' || (
    findPermission_(session.role, 'transaction.read.own') &&
    toBoolean_(historySettings.ALLOW_CASHIER_VIEW_OWN_HISTORY)
  );
  if (!canReadAll && !canReadOwn) {
    throw createAppError_('คุณไม่มีสิทธิ์ดูประวัติรายการ', 'PERMISSION_DENIED');
  }

  const query = String(data.query || '').toLowerCase().trim();
  const status = String(data.status || '').toUpperCase().trim();
  const cashierId = String(data.cashierId || '').trim();
  const fromDate = data.fromDate ? new Date(data.fromDate) : null;
  const toDate = data.toDate ? endOfDay_(new Date(data.toDate)) : null;
  const limit = clampNumber_(data.limit, 1, 500, 100);
  const offset = clampNumber_(data.offset, 0, 1000000, 0);

  let records = getSheetObjects_('Transactions');
  if (!canReadAll) records = records.filter(function (row) { return String(row.CashierID) === session.userId; });
  if (status) records = records.filter(function (row) { return String(row.Status).toUpperCase() === status; });
  if (cashierId && canReadAll) records = records.filter(function (row) { return String(row.CashierID) === cashierId; });
  if (fromDate && !isNaN(fromDate.getTime())) records = records.filter(function (row) { return dateValue_(row.CreatedAt) >= fromDate.getTime(); });
  if (toDate && !isNaN(toDate.getTime())) records = records.filter(function (row) { return dateValue_(row.CreatedAt) <= toDate.getTime(); });
  if (query) {
    records = records.filter(function (row) {
      return String(row.TransactionID).toLowerCase().indexOf(query) !== -1 ||
        String(row.CashierName).toLowerCase().indexOf(query) !== -1 ||
        String(row.Amount).toLowerCase().indexOf(query) !== -1 ||
        String(row.ThunderTransRef).toLowerCase().indexOf(query) !== -1;
    });
  }

  records.sort(function (a, b) { return dateValue_(b.CreatedAt) - dateValue_(a.CreatedAt); });
  const total = records.length;
  const pageRecords = records.slice(offset, offset + limit).map(publicTransaction_);
  return successResponse_('โหลดประวัติรายการเรียบร้อยแล้ว', {
    transactions: pageRecords,
    total: total,
    limit: limit,
    offset: offset
  });
}

function apiGetTransaction_(session, data) {
  const transactionId = String(data.transactionId || '').trim();
  const record = findRowByKey_('Transactions', 'TransactionID', transactionId);
  if (!record) return errorResponse_('ไม่พบรายการที่ต้องการ', 'TRANSACTION_NOT_FOUND');

  if (session.role !== 'ADMIN' && !findPermission_(session.role, 'transaction.read.all') &&
      String(record.object.CashierID) !== session.userId) {
    return errorResponse_('คุณไม่มีสิทธิ์ดูรายการนี้', 'PERMISSION_DENIED');
  }

  const checks = getSheetObjects_('SlipChecks').filter(function (row) {
    return String(row.TransactionID) === transactionId;
  }).sort(function (a, b) { return Number(a.AttemptNo) - Number(b.AttemptNo); });

  return successResponse_('โหลดรายละเอียดรายการเรียบร้อยแล้ว', {
    transaction: publicTransaction_(record.object),
    slipChecks: checks.map(publicSlipCheck_)
  });
}

function publicTransaction_(row) {
  return {
    transactionId: row.TransactionID,
    amount: Number(row.Amount || 0),
    currency: row.Currency || 'THB',
    status: row.Status,
    note: row.Note || '',
    cashierId: row.CashierID,
    cashierName: row.CashierName,
    deviceId: row.DeviceID,
    shiftId: row.ShiftID,
    qrType: row.QRType,
    qrPayload: row.QRPayload,
    qrMime: row.QRMime,
    qrCreatedAt: toIsoString_(row.QRCreatedAt),
    qrExpiresAt: toIsoString_(row.QRExpiresAt),
    paidAt: toIsoString_(row.PaidAt),
    slipImageUrl: row.SlipImageURL || '',
    slipViewerUrl: row.SlipViewerURL || '',
    thunderTransRef: row.ThunderTransRef || '',
    createdAt: toIsoString_(row.CreatedAt),
    updatedAt: toIsoString_(row.UpdatedAt),
    cancelledAt: toIsoString_(row.CancelledAt),
    cancelReason: row.CancelReason || ''
  };
}

/* ========================================================================== */
/*  11. DASHBOARD AND REPORTS                                                  */
/* ========================================================================== */

function apiGetDashboard_(session, data) {
  requirePermission_(session, 'reports.read');
  const targetDate = data.date ? new Date(data.date) : new Date();
  const dayStart = startOfDay_(targetDate).getTime();
  const dayEnd = endOfDay_(targetDate).getTime();
  const transactions = getSheetObjects_('Transactions').filter(function (row) {
    const timestamp = dateValue_(row.CreatedAt);
    return timestamp >= dayStart && timestamp <= dayEnd;
  });

  const summary = summarizeTransactions_(transactions);
  const monthStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1).getTime();
  const monthTransactions = getSheetObjects_('Transactions').filter(function (row) {
    return dateValue_(row.CreatedAt) >= monthStart && dateValue_(row.CreatedAt) <= dayEnd;
  });
  const monthSummary = summarizeTransactions_(monthTransactions);

  return successResponse_('โหลด Dashboard เรียบร้อยแล้ว', {
    businessDate: formatDate_(targetDate),
    today: summary,
    month: monthSummary,
    recentTransactions: transactions
      .sort(function (a, b) { return dateValue_(b.CreatedAt) - dateValue_(a.CreatedAt); })
      .slice(0, 10)
      .map(publicTransaction_)
  });
}

function apiGetReports_(session, data) {
  requirePermission_(session, 'reports.read');
  const fromDate = data.fromDate ? startOfDay_(new Date(data.fromDate)) : startOfDay_(new Date());
  const toDate = data.toDate ? endOfDay_(new Date(data.toDate)) : endOfDay_(new Date());
  if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime()) || fromDate > toDate) {
    return errorResponse_('ช่วงวันที่รายงานไม่ถูกต้อง', 'REPORT_DATE_INVALID');
  }

  let transactions = getSheetObjects_('Transactions').filter(function (row) {
    const timestamp = dateValue_(row.CreatedAt);
    return timestamp >= fromDate.getTime() && timestamp <= toDate.getTime();
  });

  if (session.role !== 'ADMIN' && !findPermission_(session.role, 'transaction.read.all')) {
    transactions = transactions.filter(function (row) { return String(row.CashierID) === session.userId; });
  }

  const dailyMap = {};
  const staffMap = {};
  const deviceMap = {};

  transactions.forEach(function (row) {
    const day = formatDate_(parseDate_(row.CreatedAt) || new Date());
    if (!dailyMap[day]) dailyMap[day] = { date: day, count: 0, amount: 0, successCount: 0, successAmount: 0 };
    dailyMap[day].count += 1;
    if (String(row.Status) === 'PAID') {
      dailyMap[day].successCount += 1;
      dailyMap[day].successAmount += Number(row.Amount || 0);
      dailyMap[day].amount += Number(row.Amount || 0);
    }

    const staffKey = String(row.CashierID || 'UNKNOWN');
    if (!staffMap[staffKey]) staffMap[staffKey] = {
      cashierId: staffKey, cashierName: row.CashierName || '', transactionCount: 0, successCount: 0, successAmount: 0
    };
    staffMap[staffKey].transactionCount += 1;
    if (String(row.Status) === 'PAID') {
      staffMap[staffKey].successCount += 1;
      staffMap[staffKey].successAmount += Number(row.Amount || 0);
    }

    const deviceKey = String(row.DeviceID || 'UNASSIGNED');
    if (!deviceMap[deviceKey]) deviceMap[deviceKey] = { deviceId: deviceKey, transactionCount: 0, successAmount: 0 };
    deviceMap[deviceKey].transactionCount += 1;
    if (String(row.Status) === 'PAID') deviceMap[deviceKey].successAmount += Number(row.Amount || 0);
  });

  return successResponse_('สร้างรายงานเรียบร้อยแล้ว', {
    fromDate: fromDate.toISOString(),
    toDate: toDate.toISOString(),
    summary: summarizeTransactions_(transactions),
    daily: Object.keys(dailyMap).sort().map(function (key) { return dailyMap[key]; }),
    byStaff: Object.keys(staffMap).map(function (key) { return staffMap[key]; }),
    byDevice: Object.keys(deviceMap).map(function (key) { return deviceMap[key]; })
  });
}

function summarizeTransactions_(transactions) {
  const summary = {
    totalCount: transactions.length,
    successCount: 0,
    pendingCount: 0,
    expiredCount: 0,
    cancelledCount: 0,
    failedCount: 0,
    successAmount: 0
  };
  transactions.forEach(function (row) {
    const status = String(row.Status || '').toUpperCase();
    if (status === 'PAID') {
      summary.successCount += 1;
      summary.successAmount += Number(row.Amount || 0);
    } else if (status === 'PENDING') summary.pendingCount += 1;
    else if (status === 'EXPIRED') summary.expiredCount += 1;
    else if (status === 'CANCELLED') summary.cancelledCount += 1;
    else summary.failedCount += 1;
  });
  summary.successAmount = roundMoney_(summary.successAmount);
  return summary;
}

/* ========================================================================== */
/*  12. SHIFT MANAGEMENT                                                       */
/* ========================================================================== */

function apiOpenShift_(session, data) {
  requirePermission_(session, 'shift.manage');
  const deviceId = String(data.deviceId || session.deviceId || '').trim();
  const existing = getOpenShiftForUser_(session.userId);
  if (existing) {
    return errorResponse_('คุณมีกะที่เปิดใช้งานอยู่แล้ว', 'SHIFT_ALREADY_OPEN', {
      shiftId: existing.ShiftID
    });
  }

  const now = new Date();
  const shift = {
    ShiftID: createId_('SFT'),
    UserID: session.userId,
    UserName: session.fullName,
    DeviceID: deviceId,
    Status: 'OPEN',
    OpenedAt: now,
    ClosedAt: '',
    OpeningNote: String(data.note || '').trim().substring(0, 250),
    ClosingNote: '',
    TransactionCount: 0,
    TotalAmount: 0,
    CreatedAt: now,
    UpdatedAt: now
  };
  appendObject_('Shifts', shift);
  logAudit_(session, 'SHIFT_OPENED', 'SHIFT', shift.ShiftID, { deviceId: deviceId });
  return successResponse_('เปิดกะทำงานเรียบร้อยแล้ว', { shift: publicRecord_(shift) });
}

function apiCloseShift_(session, data) {
  requirePermission_(session, 'shift.manage');
  const shiftId = String(data.shiftId || '').trim();
  const record = shiftId
    ? findRowByKey_('Shifts', 'ShiftID', shiftId)
    : findOpenShiftRecordForUser_(session.userId);
  if (!record) return errorResponse_('ไม่พบกะที่ต้องการปิด', 'SHIFT_NOT_FOUND');
  if (session.role !== 'ADMIN' && String(record.object.UserID) !== session.userId) {
    return errorResponse_('คุณไม่มีสิทธิ์ปิดกะนี้', 'PERMISSION_DENIED');
  }
  if (String(record.object.Status) !== 'OPEN') return errorResponse_('กะนี้ปิดไปแล้ว', 'SHIFT_NOT_OPEN');

  const summary = getShiftSummary_(record.object.ShiftID);
  const now = new Date();
  updateRowObject_('Shifts', record.rowIndex, {
    Status: 'CLOSED',
    ClosedAt: now,
    ClosingNote: String(data.note || '').trim().substring(0, 250),
    TransactionCount: summary.transactionCount,
    TotalAmount: summary.totalAmount,
    UpdatedAt: now
  });

  logAudit_(session, 'SHIFT_CLOSED', 'SHIFT', record.object.ShiftID, summary);
  return successResponse_('ปิดกะทำงานเรียบร้อยแล้ว', {
    shiftId: record.object.ShiftID,
    summary: summary,
    closedAt: now.toISOString()
  });
}

function apiGetMyShift_(session) {
  const shift = getOpenShiftForUser_(session.userId);
  if (!shift) return successResponse_('ยังไม่มีกะที่เปิดใช้งาน', { shift: null });
  return successResponse_('โหลดข้อมูลกะเรียบร้อยแล้ว', {
    shift: publicRecord_(shift),
    summary: getShiftSummary_(shift.ShiftID)
  });
}

function apiListShifts_(session, data) {
  requirePermission_(session, 'settlement.read');
  let records = getSheetObjects_('Shifts');
  if (session.role !== 'ADMIN' && !findPermission_(session.role, 'transaction.read.all')) {
    records = records.filter(function (row) { return String(row.UserID) === session.userId; });
  }
  const status = String(data.status || '').toUpperCase();
  if (status) records = records.filter(function (row) { return String(row.Status).toUpperCase() === status; });
  records.sort(function (a, b) { return dateValue_(b.OpenedAt) - dateValue_(a.OpenedAt); });
  return successResponse_('โหลดข้อมูลกะเรียบร้อยแล้ว', {
    shifts: records.slice(0, clampNumber_(data.limit, 1, 500, 100)).map(publicRecord_)
  });
}

function getOpenShiftForUser_(userId) {
  const records = getSheetObjects_('Shifts');
  return records.find(function (row) {
    return String(row.UserID) === String(userId) && String(row.Status) === 'OPEN';
  }) || null;
}

function findOpenShiftRecordForUser_(userId) {
  return findRowByPredicate_('Shifts', function (row) {
    return String(row.UserID) === String(userId) && String(row.Status) === 'OPEN';
  });
}

function getShiftSummary_(shiftId) {
  const rows = getSheetObjects_('Transactions').filter(function (row) {
    return String(row.ShiftID) === String(shiftId);
  });
  const paid = rows.filter(function (row) { return String(row.Status) === 'PAID'; });
  return {
    transactionCount: rows.length,
    successCount: paid.length,
    totalAmount: roundMoney_(paid.reduce(function (sum, row) { return sum + Number(row.Amount || 0); }, 0))
  };
}

function updateShiftTotals_(shiftId) {
  const record = findRowByKey_('Shifts', 'ShiftID', shiftId);
  if (!record) return;
  const summary = getShiftSummary_(shiftId);
  updateRowObject_('Shifts', record.rowIndex, {
    TransactionCount: summary.transactionCount,
    TotalAmount: summary.totalAmount,
    UpdatedAt: new Date()
  });
}

/* ========================================================================== */
/*  13. REFUND REQUEST AND REVIEW                                              */
/* ========================================================================== */

function apiRequestRefund_(session, data) {
  requirePermission_(session, 'refund.request');
  const transactionId = String(data.transactionId || '').trim();
  const reason = String(data.reason || '').trim();
  if (!transactionId || !reason) return errorResponse_('กรุณาระบุรายการและเหตุผลการคืนเงิน', 'REFUND_DATA_REQUIRED');

  const transaction = findRowByKey_('Transactions', 'TransactionID', transactionId);
  if (!transaction || String(transaction.object.Status) !== 'PAID') {
    return errorResponse_('ไม่พบรายการชำระเงินที่สามารถขอคืนเงินได้', 'REFUND_TRANSACTION_INVALID');
  }

  const existing = getSheetObjects_('Refunds').some(function (row) {
    return String(row.TransactionID) === transactionId &&
      ['REQUESTED', 'APPROVED_PENDING_PROVIDER', 'PROCESSING', 'REFUNDED'].indexOf(String(row.Status)) !== -1;
  });
  if (existing) return errorResponse_('รายการนี้มีคำขอคืนเงินอยู่แล้ว', 'REFUND_ALREADY_EXISTS');

  const now = new Date();
  const refundId = createId_('RFD');
  appendObject_('Refunds', {
    RefundID: refundId,
    TransactionID: transactionId,
    Amount: Number(transaction.object.Amount),
    Reason: reason.substring(0, 500),
    Status: 'REQUESTED',
    RequestedBy: session.userId,
    RequestedByName: session.fullName,
    RequestedAt: now,
    ReviewedBy: '',
    ReviewedByName: '',
    ReviewedAt: '',
    ReviewNote: '',
    ProviderReference: '',
    ProcessedAt: '',
    CreatedAt: now,
    UpdatedAt: now
  });
  logAudit_(session, 'REFUND_REQUESTED', 'REFUND', refundId, { transactionId: transactionId, amount: transaction.object.Amount });
  createNotification_('REFUND_REQUESTED', 'มีคำขอคืนเงิน', 'รายการ ' + transactionId + ' มีคำขอคืนเงิน', '', refundId);
  return successResponse_('ส่งคำขอคืนเงินเรียบร้อยแล้ว รอผู้ดูแลตรวจสอบ', { refundId: refundId });
}

function apiReviewRefund_(session, data) {
  requirePermission_(session, 'refund.review');
  const refundId = String(data.refundId || '').trim();
  const decision = String(data.decision || '').toUpperCase();
  if (['APPROVE', 'REJECT'].indexOf(decision) === -1) {
    return errorResponse_('คำสั่งอนุมัติไม่ถูกต้อง', 'REFUND_DECISION_INVALID');
  }
  const record = findRowByKey_('Refunds', 'RefundID', refundId);
  if (!record || String(record.object.Status) !== 'REQUESTED') {
    return errorResponse_('ไม่พบคำขอคืนเงินที่รอพิจารณา', 'REFUND_NOT_PENDING');
  }

  const now = new Date();
  const status = decision === 'APPROVE' ? 'APPROVED_PENDING_PROVIDER' : 'REJECTED';
  updateRowObject_('Refunds', record.rowIndex, {
    Status: status,
    ReviewedBy: session.userId,
    ReviewedByName: session.fullName,
    ReviewedAt: now,
    ReviewNote: String(data.note || '').trim().substring(0, 500),
    UpdatedAt: now
  });
  logAudit_(session, 'REFUND_REVIEWED', 'REFUND', refundId, { decision: decision, status: status });

  const message = decision === 'APPROVE'
    ? 'อนุมัติคำขอแล้ว แต่ยังต้องดำเนินการคืนเงินจริงผ่านผู้ให้บริการชำระเงิน'
    : 'ปฏิเสธคำขอคืนเงินเรียบร้อยแล้ว';
  return successResponse_(message, { refundId: refundId, status: status, actualRefundProcessed: false });
}

function apiListRefunds_(session, data) {
  requirePermission_(session, 'settlement.read');
  let records = getSheetObjects_('Refunds');
  const status = String(data.status || '').toUpperCase();
  if (status) records = records.filter(function (row) { return String(row.Status).toUpperCase() === status; });
  if (session.role !== 'ADMIN') {
    records = records.filter(function (row) { return String(row.RequestedBy) === session.userId; });
  }
  records.sort(function (a, b) { return dateValue_(b.RequestedAt) - dateValue_(a.RequestedAt); });
  return successResponse_('โหลดรายการคืนเงินเรียบร้อยแล้ว', {
    refunds: records.slice(0, clampNumber_(data.limit, 1, 500, 100)).map(publicRecord_)
  });
}

/* ========================================================================== */
/*  14. DEVICE MANAGEMENT                                                      */
/* ========================================================================== */

function apiListDevices_(session) {
  requirePermission_(session, 'device.manage');
  const devices = getSheetObjects_('Devices').sort(function (a, b) {
    return String(a.DeviceName).localeCompare(String(b.DeviceName));
  });
  return successResponse_('โหลดข้อมูลอุปกรณ์เรียบร้อยแล้ว', { devices: devices.map(publicRecord_) });
}

function apiSaveDevice_(session, data) {
  requirePermission_(session, 'device.manage');
  const deviceId = String(data.deviceId || '').trim() || createId_('DEV');
  const deviceName = String(data.deviceName || '').trim();
  if (!deviceName || deviceName.length > 80) {
    return errorResponse_('กรุณาระบุชื่ออุปกรณ์', 'DEVICE_NAME_REQUIRED');
  }
  const requestedStatus = String(data.status || 'ACTIVE').toUpperCase();
  if (['ACTIVE', 'INACTIVE'].indexOf(requestedStatus) === -1) {
    return errorResponse_('สถานะอุปกรณ์ไม่ถูกต้อง', 'DEVICE_STATUS_INVALID');
  }

  const record = findRowByKey_('Devices', 'DeviceID', deviceId);
  const now = new Date();
  const values = {
    DeviceID: deviceId,
    DeviceName: deviceName,
    Status: String(data.status || (record ? record.object.Status : 'ACTIVE')).toUpperCase(),
    AssignedUserID: String(data.assignedUserId || '').trim(),
    LastSeenAt: record ? record.object.LastSeenAt : '',
    CreatedAt: record ? record.object.CreatedAt : now,
    UpdatedAt: now,
    Notes: String(data.notes || '').trim().substring(0, 250)
  };
  if (record) updateRowObject_('Devices', record.rowIndex, values);
  else appendObject_('Devices', values);

  logAudit_(session, record ? 'DEVICE_UPDATED' : 'DEVICE_CREATED', 'DEVICE', deviceId, values);
  return successResponse_('บันทึกข้อมูลอุปกรณ์เรียบร้อยแล้ว', { deviceId: deviceId });
}

function apiSetDeviceStatus_(session, data) {
  requirePermission_(session, 'device.manage');
  const deviceId = String(data.deviceId || '').trim();
  const status = String(data.status || '').toUpperCase();
  if (['ACTIVE', 'INACTIVE'].indexOf(status) === -1) return errorResponse_('สถานะอุปกรณ์ไม่ถูกต้อง', 'DEVICE_STATUS_INVALID');
  const record = findRowByKey_('Devices', 'DeviceID', deviceId);
  if (!record) return errorResponse_('ไม่พบอุปกรณ์', 'DEVICE_NOT_FOUND');
  updateRowObject_('Devices', record.rowIndex, { Status: status, UpdatedAt: new Date() });
  logAudit_(session, 'DEVICE_STATUS_CHANGED', 'DEVICE', deviceId, { status: status });
  return successResponse_('เปลี่ยนสถานะอุปกรณ์เรียบร้อยแล้ว', { deviceId: deviceId, status: status });
}

function touchDevice_(deviceId, userId) {
  const record = findRowByKey_('Devices', 'DeviceID', deviceId);
  const now = new Date();
  if (record) {
    updateRowObject_('Devices', record.rowIndex, {
      LastSeenAt: now,
      AssignedUserID: userId || record.object.AssignedUserID,
      UpdatedAt: now
    });
  } else {
    appendObject_('Devices', {
      DeviceID: deviceId,
      DeviceName: APP_CONFIG.DEFAULT_DEVICE_NAME,
      Status: 'ACTIVE',
      AssignedUserID: userId || '',
      LastSeenAt: now,
      CreatedAt: now,
      UpdatedAt: now,
      Notes: 'สร้างอัตโนมัติเมื่อมีการใช้งานครั้งแรก'
    });
  }
}

/* ========================================================================== */
/*  15. NOTIFICATIONS                                                          */
/* ========================================================================== */

function createNotification_(type, title, message, recipientId, relatedId) {
  appendObject_('Notifications', {
    NotificationID: createId_('NTF'),
    Type: type,
    Title: title,
    Message: message,
    RecipientID: recipientId || '',
    IsRead: false,
    CreatedAt: new Date(),
    ReadAt: '',
    RelatedID: relatedId || ''
  });
}

function apiGetNotifications_(session, data) {
  requirePermission_(session, 'notifications.read');
  let records = getSheetObjects_('Notifications').filter(function (row) {
    return !row.RecipientID || String(row.RecipientID) === session.userId || session.role === 'ADMIN';
  });
  if (data.unreadOnly) {
    records = records.filter(function (row) { return !toBoolean_(row.IsRead); });
  }
  records.sort(function (a, b) { return dateValue_(b.CreatedAt) - dateValue_(a.CreatedAt); });
  return successResponse_('โหลดการแจ้งเตือนเรียบร้อยแล้ว', {
    notifications: records.slice(0, clampNumber_(data.limit, 1, 200, 50)).map(publicRecord_)
  });
}

function apiMarkNotificationRead_(session, data) {
  requirePermission_(session, 'notifications.read');
  const notificationId = String(data.notificationId || '').trim();
  const record = findRowByKey_('Notifications', 'NotificationID', notificationId);
  if (!record) return errorResponse_('ไม่พบการแจ้งเตือน', 'NOTIFICATION_NOT_FOUND');
  if (record.object.RecipientID && String(record.object.RecipientID) !== session.userId && session.role !== 'ADMIN') {
    return errorResponse_('คุณไม่มีสิทธิ์ดำเนินการนี้', 'PERMISSION_DENIED');
  }
  updateRowObject_('Notifications', record.rowIndex, { IsRead: true, ReadAt: new Date() });
  return successResponse_('ทำเครื่องหมายว่าอ่านแล้ว', {});
}

/* ========================================================================== */
/*  16. AUDIT, SYSTEM LOGS AND API USAGE                                       */
/* ========================================================================== */

function logAudit_(session, action, entityType, entityId, details) {
  logAuditSystem_(session.userId, session.fullName, action, entityType, entityId, details, session.role);
}

function logAuditSystem_(userId, userName, action, entityType, entityId, details, role) {
  appendObject_('AuditLogs', {
    LogID: createId_('LOG'),
    UserID: userId || 'SYSTEM',
    UserName: userName || 'SYSTEM',
    Role: role || 'SYSTEM',
    Action: action,
    EntityType: entityType,
    EntityID: entityId,
    Details: safeStringify_(details || {}),
    CreatedAt: new Date()
  });
}

function apiListAuditLogs_(session, data) {
  requirePermission_(session, 'audit.read');
  const limit = clampNumber_(data.limit, 1, 500, 100);
  const records = getSheetObjects_('AuditLogs').sort(function (a, b) {
    return dateValue_(b.CreatedAt) - dateValue_(a.CreatedAt);
  });
  return successResponse_('โหลดประวัติกิจกรรมเรียบร้อยแล้ว', {
    logs: records.slice(0, limit).map(publicRecord_)
  });
}

function logSystem_(level, publicMessage, internalDetail, action) {
  try {
    appendObject_('SystemLogs', {
      LogID: createId_('SYS'),
      Level: level,
      PublicMessage: publicMessage,
      InternalDetail: String(internalDetail || '').substring(0, 45000),
      Action: action || '',
      CreatedAt: new Date()
    });
  } catch (ignored) {
    // Logging failure must not cause a second failure loop.
  }
}

function logSystemError_(action, error, context) {
  logSystem_('ERROR', 'เกิดข้อผิดพลาดในการดำเนินการ',
    (error && error.stack ? error.stack : String(error)) + '\nContext: ' + String(context || ''), action);
}

function apiListSystemLogs_(session, data) {
  requirePermission_(session, 'systemlogs.read');
  const limit = clampNumber_(data.limit, 1, 500, 100);
  const records = getSheetObjects_('SystemLogs').sort(function (a, b) {
    return dateValue_(b.CreatedAt) - dateValue_(a.CreatedAt);
  });
  return successResponse_('โหลดบันทึกระบบเรียบร้อยแล้ว', {
    logs: records.slice(0, limit).map(function (row) {
      return {
        logId: row.LogID,
        level: row.Level,
        publicMessage: row.PublicMessage,
        internalDetail: row.InternalDetail,
        action: row.Action,
        createdAt: toIsoString_(row.CreatedAt)
      };
    })
  });
}

function recordApiUsage_(provider, endpoint, transactionId, httpStatus, durationMs, success) {
  try {
    appendObject_('APIUsage', {
      APIUsageID: createId_('API'),
      Provider: provider,
      Endpoint: endpoint,
      TransactionID: transactionId || '',
      HttpStatus: httpStatus,
      DurationMs: durationMs,
      Success: !!success,
      CreatedAt: new Date()
    });
  } catch (ignored) {
    // API usage logging is best-effort.
  }
}

function apiListApiUsage_(session, data) {
  requireAdmin_(session);
  const rows = getSheetObjects_('APIUsage').sort(function (a, b) {
    return dateValue_(b.CreatedAt) - dateValue_(a.CreatedAt);
  });
  return successResponse_('โหลดประวัติการเรียก API เรียบร้อยแล้ว', {
    usage: rows.slice(0, clampNumber_(data.limit, 1, 500, 100)).map(publicRecord_)
  });
}

/* ========================================================================== */
/*  17. DAILY CLOSING / SETTLEMENT                                             */
/* ========================================================================== */

function apiCloseBusinessDay_(session, data) {
  requirePermission_(session, 'settlement.read');
  const targetDate = data.businessDate ? new Date(data.businessDate) : new Date();
  if (isNaN(targetDate.getTime())) return errorResponse_('วันที่ปิดยอดไม่ถูกต้อง', 'CLOSING_DATE_INVALID');
  const dateLabel = formatDate_(targetDate);
  const deviceId = String(data.deviceId || '').trim();
  const userId = String(data.userId || '').trim();

  let rows = getSheetObjects_('Transactions').filter(function (row) {
    return formatDate_(parseDate_(row.CreatedAt) || new Date(0)) === dateLabel;
  });
  if (deviceId) rows = rows.filter(function (row) { return String(row.DeviceID) === deviceId; });
  if (userId) rows = rows.filter(function (row) { return String(row.CashierID) === userId; });
  if (session.role !== 'ADMIN' && !findPermission_(session.role, 'transaction.read.all')) {
    rows = rows.filter(function (row) { return String(row.CashierID) === session.userId; });
  }

  const summary = summarizeTransactions_(rows);
  const refundRows = getSheetObjects_('Refunds').filter(function (row) {
    return String(row.Status) === 'REFUNDED' &&
      formatDate_(parseDate_(row.ProcessedAt) || new Date(0)) === dateLabel;
  });
  const refundAmount = roundMoney_(refundRows.reduce(function (sum, row) {
    return sum + Number(row.Amount || 0);
  }, 0));
  const now = new Date();
  const closing = {
    ClosingID: createId_('CLS'),
    BusinessDate: dateLabel,
    DeviceID: deviceId,
    UserID: userId,
    TransactionCount: summary.successCount,
    SuccessAmount: summary.successAmount,
    RefundAmount: refundAmount,
    NetAmount: roundMoney_(summary.successAmount - refundAmount),
    ClosedBy: session.userId,
    ClosedByName: session.fullName,
    ClosedAt: now,
    Notes: String(data.notes || '').trim().substring(0, 500)
  };
  appendObject_('DailyClosings', closing);
  logAudit_(session, 'BUSINESS_DAY_CLOSED', 'DAILY_CLOSING', closing.ClosingID, closing);

  return successResponse_('สรุปยอดประจำวันเรียบร้อยแล้ว', {
    closing: publicRecord_(closing),
    summary: summary
  });
}

/* ========================================================================== */
/*  18. DATABASE DIAGNOSTICS                                                   */
/* ========================================================================== */

function apiGetDatabaseStatus_(session) {
  requireAdmin_(session);
  const properties = PropertiesService.getScriptProperties().getProperties();
  return successResponse_('ตรวจสอบการตั้งค่าระบบเรียบร้อยแล้ว', {
    database: getDatabaseStatus_(),
    integrations: {
      spreadsheetConfigured: !!properties.SPREADSHEET_ID,
      thunderConfigured: !!properties.THUNDER_API_KEY,
      imgbbConfigured: !!properties.IMGBB_API_KEY
    },
    backendVersion: APP_CONFIG.VERSION
  });
}

/* ========================================================================== */
/*  19. GOOGLE SHEETS DATA ACCESS                                              */
/* ========================================================================== */

function getDatabaseSpreadsheet_() {
  const spreadsheetId = getScriptProperty_('SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('SPREADSHEET_ID is not set in Script Properties.');
}

function getSheet_(sheetName) {
  const spreadsheet = getDatabaseSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    if (!SHEET_SCHEMAS[sheetName]) throw new Error('Unknown sheet: ' + sheetName);
    ensureDatabaseStructure_();
    sheet = spreadsheet.getSheetByName(sheetName);
  }
  return sheet;
}

function getSheetRows_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return { headers: [], records: [] };

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(function (header) { return String(header || '').trim(); });
  const records = [];
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const rowValues = values[rowIndex];
    if (rowValues.every(function (value) { return value === '' || value === null; })) continue;
    const object = {};
    headers.forEach(function (header, columnIndex) {
      if (header) object[header] = rowValues[columnIndex];
    });
    records.push({ rowIndex: rowIndex + 1, object: object });
  }
  return { headers: headers, records: records };
}

function getSheetObjects_(sheetName) {
  return getSheetRows_(sheetName).records.map(function (record) { return record.object; });
}

function appendObject_(sheetName, object) {
  const sheet = getSheet_(sheetName);
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function (header) { return String(header || '').trim(); });
  const values = headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(object, header) ? object[header] : '';
  });
  sheet.appendRow(values);
  return sheet.getLastRow();
}

function appendObjects_(sheetName, objects) {
  if (!objects || !objects.length) return 0;
  const sheet = getSheet_(sheetName);
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function (header) { return String(header || '').trim(); });
  const values = objects.map(function (object) {
    return headers.map(function (header) {
      return Object.prototype.hasOwnProperty.call(object, header) ? object[header] : '';
    });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
  return values.length;
}

function findRowByKey_(sheetName, keyName, keyValue) {
  const result = getSheetRows_(sheetName);
  for (let i = 0; i < result.records.length; i += 1) {
    const record = result.records[i];
    if (String(record.object[keyName]) === String(keyValue)) {
      return { rowIndex: record.rowIndex, object: record.object };
    }
  }
  return null;
}

function findRowByPredicate_(sheetName, predicate) {
  const records = getSheetRows_(sheetName).records;
  for (let i = 0; i < records.length; i += 1) {
    if (predicate(records[i].object)) return records[i];
  }
  return null;
}

function updateRowObject_(sheetName, rowIndex, updates) {
  const sheet = getSheet_(sheetName);
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function (header) { return String(header || '').trim(); });
  const range = sheet.getRange(rowIndex, 1, 1, lastColumn);
  const currentValues = range.getValues()[0];
  headers.forEach(function (header, index) {
    if (Object.prototype.hasOwnProperty.call(updates, header)) {
      currentValues[index] = updates[header];
    }
  });
  range.setValues([currentValues]);
}

/* ========================================================================== */
/*  20. COMMON UTILITIES                                                       */
/* ========================================================================== */

function successResponse_(message, data) {
  return { ok: true, message: message || 'ดำเนินการเรียบร้อยแล้ว', data: data || {} };
}

function errorResponse_(message, code, data) {
  return { ok: false, message: message || 'ไม่สามารถดำเนินการได้', code: code || 'GENERAL_ERROR', data: data || {} };
}

function createAppError_(message, code, cause) {
  const error = new Error(message);
  error.name = 'EDCAppError';
  error.code = code || 'GENERAL_ERROR';
  if (cause) error.internalCause = cause;
  return error;
}

function publicMessageForError_(error) {
  if (error && error.name === 'EDCAppError') return error.message;
  return 'ไม่สามารถดำเนินการได้ กรุณาลองใหม่อีกครั้ง';
}

function getScriptProperty_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function ensureAccessCodePepper_() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('ACCESS_CODE_PEPPER')) {
    properties.setProperty('ACCESS_CODE_PEPPER', createRandomToken_());
  }
}

function hashAccessCode_(accessCode, salt) {
  const pepper = getScriptProperty_('ACCESS_CODE_PEPPER');
  if (!pepper) throw new Error('ACCESS_CODE_PEPPER is not initialized. Run setupDatabase().');
  const raw = String(salt || '') + '|' + String(accessCode || '');
  const digest = Utilities.computeHmacSha256Signature(raw, pepper, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function hashToken_(token) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(token || ''), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function constantTimeEquals_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i += 1) result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return result === 0;
}

function createRandomToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}

function generateRandomDigits_(length) {
  let digits = '';
  while (digits.length < length) {
    const hex = Utilities.getUuid().replace(/-/g, '').toLowerCase();
    for (let i = 0; i < hex.length && digits.length < length; i += 1) {
      const value = parseInt(hex.charAt(i), 16);
      if (value < 10) digits += String(value);
      else digits += String(value - 10);
    }
  }
  return digits;
}

function createId_(prefix) {
  const timestamp = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMddHHmmss');
  const random = Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
  return String(prefix || APP_CONFIG.ID_PREFIX) + '-' + timestamp + '-' + random;
}

function createTransactionId_() {
  const timestamp = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMddHHmmss');
  const random = Utilities.getUuid().replace(/-/g, '').substring(0, 6).toUpperCase();
  return 'TX' + timestamp + random;
}

function parseJsonResponse_(text) {
  try {
    return JSON.parse(text || '{}');
  } catch (error) {
    return { parseError: true, raw: String(text || '').substring(0, 2000) };
  }
}

function safeStringify_(value) {
  try {
    return JSON.stringify(value, function (key, item) {
      if (key === 'delete_url' || key === 'DeleteURL' || key === 'TokenHash' || key === 'AccessCodeHash') {
        return '[REDACTED]';
      }
      return item;
    });
  } catch (error) {
    return String(value);
  }
}

function roundMoney_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function clampNumber_(value, min, max, fallback) {
  const number = Number(value);
  if (!isFinite(number)) return fallback;
  return Math.min(max, Math.max(min, Math.floor(number)));
}

function toBoolean_(value) {
  if (typeof value === 'boolean') return value;
  const text = String(value || '').toLowerCase().trim();
  return text === 'true' || text === '1' || text === 'yes' || text === 'on';
}

function parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function dateValue_(value) {
  const date = parseDate_(value);
  return date ? date.getTime() : 0;
}

function toIsoString_(value) {
  const date = parseDate_(value);
  return date ? date.toISOString() : '';
}

function formatDate_(date) {
  const parsed = parseDate_(date);
  return parsed ? Utilities.formatDate(parsed, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd') : '';
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function endOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

function publicRecord_(record) {
  const output = {};
  Object.keys(record || {}).forEach(function (key) {
    const value = record[key];
    if (value instanceof Date) output[key] = value.toISOString();
    else if (value === undefined) output[key] = '';
    else output[key] = value;
  });
  return output;
}

/* ========================================================================== */
/*  21. MANUAL ADMIN UTILITIES                                                 */
/* ========================================================================== */

/**
 * Run after setting the Script Properties. It is safe to run more than once.
 * It will add missing sheets/columns but will not delete existing data.
 */
function runInitialSetup() {
  const result = setupDatabase();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Optional: run this from the editor after editing Settings sheet values to
 * ensure new schema fields are present. Existing data is preserved.
 */
function runSchemaCheck() {
  const result = ensureDatabaseStructure_();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Safe manual check. Does not call paid Thunder verification endpoints and
 * does not generate a payment QR. It only reports which credentials exist.
 */
function checkProjectConfiguration() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const result = {
    spreadsheetConfigured: !!properties.SPREADSHEET_ID,
    adminInitialCodeConfigured: !!properties.ADMIN_INITIAL_CODE,
    thunderApiKeyConfigured: !!properties.THUNDER_API_KEY,
    imgbbApiKeyConfigured: !!properties.IMGBB_API_KEY,
    spreadsheetAccessible: false,
    databaseStatus: []
  };

  try {
    const spreadsheet = getDatabaseSpreadsheet_();
    result.spreadsheetAccessible = true;
    result.spreadsheetName = spreadsheet.getName();
    result.databaseStatus = getDatabaseStatus_();
  } catch (error) {
    result.internalError = String(error);
  }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
