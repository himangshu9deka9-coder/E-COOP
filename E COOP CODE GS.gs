/**
 * E-COOPERATIVE — GOOGLE APPS SCRIPT BACKEND
 * ============================================================
 * Deploy this as a Web App (see instructions at the bottom of this file).
 * The E-Cooperative HTML app calls this via fetch() -- no OAuth, no API
 * keys, no per-user Google sign-in. It runs under YOUR Google account.
 *
 * Structure it creates automatically:
 *
 *   Drive:
 *     E-Cooperative Data/
 *       <District>/
 *         ARCS/  or  DRCS/
 *           <Document Type>/
 *             <uploaded files>
 *
 *   Sheets:
 *     One spreadsheet per district: "E-Cooperative — <District>"
 *       Tabs: ARCS_Society, DRCS_Society, ARCS_Office, DRCS_Office, ...
 *             (one tab per "<Office>_<Module>")
 *
 *     One separate spreadsheet for everyone's login data:
 *       "E-Cooperative Users (Login Data)"
 * ============================================================
 */

const ROOT_FOLDER_NAME = 'E-Cooperative Data';
const USERS_SHEET_NAME = 'E-Cooperative Users (Login Data)';

// ─── ENTRY POINTS ────────────────────────────────────────────────────

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const params = (e.postData && e.postData.contents) ?
      JSON.parse(e.postData.contents) : (e.parameter || {});
    const action = params.action;
    let result;

    switch (action) {
      case 'ping':
        result = { ok: true, message: 'E-Cooperative backend is alive.' };
        break;
      case 'syncLoginData':
        result = syncLoginData(params);
        break;
      case 'syncModuleRows':
        result = syncModuleRows(params);
        break;
      case 'getModuleRows':
        result = getModuleRows(params);
        break;
      case 'uploadDocument':
        result = uploadDocument(params);
        break;
      case 'deleteDocument':
        result = deleteDocument(params);
        break;
      case 'requestAdminOtp':
        result = requestAdminOtp(params);
        break;
      case 'confirmAdminOtp':
        result = confirmAdminOtp(params);
        break;
      case 'signupAccount':
        result = signupAccount(params);
        break;
      case 'loginAccount':
        result = loginAccount(params);
        break;
      case 'findAccount':
        result = findAccount(params);
        break;
      case 'requestPasswordReset':
        result = requestPasswordReset(params);
        break;
      case 'confirmPasswordReset':
        result = confirmPasswordReset(params);
        break;
      case 'updateAccountProfile':
        result = updateAccountProfile(params);
        break;
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message, stack: err.stack });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── FOLDER / SPREADSHEET RESOLUTION ─────────────────────────────────

function getOrCreateFolder(name, parentFolder) {
  const parent = parentFolder || DriveApp.getRootFolder();
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function getRootFolder() {
  return getOrCreateFolder(ROOT_FOLDER_NAME, DriveApp.getRootFolder());
}

// District -> Office -> Document type folder, created as needed.
function getDocFolder(district, officeType, docType) {
  const root = getRootFolder();
  const districtFolder = getOrCreateFolder(district, root);
  const officeFolder = getOrCreateFolder(officeType, districtFolder);
  return getOrCreateFolder(docType, officeFolder);
}

function getOrCreateDistrictSpreadsheet(district) {
  const root = getRootFolder();
  const name = 'E-Cooperative — ' + district;
  const files = root.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  const ss = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(ss.getId());
  root.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // created spreadsheets default to root -- move it
  return ss;
}

function getOrCreateUsersSpreadsheet() {
  const root = getRootFolder();
  const files = root.getFilesByName(USERS_SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  const ss = SpreadsheetApp.create(USERS_SHEET_NAME);
  const file = DriveApp.getFileById(ss.getId());
  root.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return ss;
}

function getOrCreateSheetTab(ss, tabName) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    // Remove the default "Sheet1" if this is the very first real tab added.
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  }
  return sheet;
}

// ─── ROW READ / WRITE (generic, used by every module) ────────────────

// Overwrites a tab with an array of flat row objects. Header row is the
// union of every object's keys.
function writeRows(sheet, rows) {
  sheet.clearContents();
  if (!rows || rows.length === 0) return;
  const keySet = {};
  rows.forEach(r => Object.keys(r).forEach(k => keySet[k] = true));
  const headers = Object.keys(keySet);
  const values = [headers].concat(rows.map(r => headers.map(h => {
    const v = r[h];
    if (v === undefined || v === null) return '';
    return (typeof v === 'object') ? JSON.stringify(v) : v;
  })));
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
}

function readRows(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] !== undefined ? row[i] : '');
    return obj;
  });
}

// ─── ACTION HANDLERS ──────────────────────────────────────────────────

// Writes/merges this district+office's DDO and Inspector list into the
// single shared Users spreadsheet (each district/office owns its own rows).
function syncLoginData(params) {
  const district = params.district;
  const officeType = params.officeType;
  const rows = params.rows || []; // [{Name, Phone, Role, Designation}, ...]
  if (!district || !officeType) throw new Error('district and officeType are required.');

  const ss = getOrCreateUsersSpreadsheet();
  const sheet = getOrCreateSheetTab(ss, 'Users');
  const existing = readRows(sheet);
  const filtered = existing.filter(r => !(r.District === district && r.OfficeType === officeType));
  const tagged = rows.map(r => Object.assign({}, r, { District: district, OfficeType: officeType }));
  writeRows(sheet, filtered.concat(tagged));
  return { ok: true, written: tagged.length };
}

// Generic: overwrite one "<Office>_<Module>" tab in one district's
// spreadsheet with the given rows. Used for Society, Office, Scheme, etc.
function syncModuleRows(params) {
  const district = params.district;
  const officeType = params.officeType;
  const module = params.module; // e.g. "Society", "Office", "Registration"
  const rows = params.rows || [];
  if (!district || !officeType || !module) {
    throw new Error('district, officeType, and module are required.');
  }
  const ss = getOrCreateDistrictSpreadsheet(district);
  const tabName = officeType + '_' + module;
  const sheet = getOrCreateSheetTab(ss, tabName);
  writeRows(sheet, rows);
  return { ok: true, tab: tabName, written: rows.length };
}

// Reads a module's tab back (for pulling data on a second device).
function getModuleRows(params) {
  const district = params.district;
  const officeType = params.officeType;
  const module = params.module;
  if (!district || !officeType || !module) {
    throw new Error('district, officeType, and module are required.');
  }
  const ss = getOrCreateDistrictSpreadsheet(district);
  const tabName = officeType + '_' + module;
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return { ok: true, rows: [] };
  return { ok: true, rows: readRows(sheet) };
}

// Accepts a base64 data URL (exactly what the app already stores locally
// for documents) and saves it into the right Drive folder.
function uploadDocument(params) {
  const district = params.district;
  const officeType = params.officeType;
  const docType = params.docType; // e.g. "Registration Documents", "Balance Sheets"
  const fileName = params.fileName;
  const dataUrl = params.dataUrl; // "data:<mime>;base64,<data>"
  if (!district || !officeType || !docType || !fileName || !dataUrl) {
    throw new Error('district, officeType, docType, fileName, and dataUrl are required.');
  }
  const match = dataUrl.match(/^data:([^;]+);base64,(.*)$/);
  if (!match) throw new Error('dataUrl is not a valid base64 data URL.');
  const mimeType = match[1];
  const base64Data = match[2];

  const folder = getDocFolder(district, officeType, docType);
  // Replace an existing file with the same name so re-uploads don't pile up.
  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) existing.next().setTrashed(true);

  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = folder.createFile(blob);
  return { ok: true, fileId: file.getId(), url: file.getUrl() };
}

function deleteDocument(params) {
  const district = params.district;
  const officeType = params.officeType;
  const docType = params.docType;
  const fileName = params.fileName;
  if (!district || !officeType || !docType || !fileName) {
    throw new Error('district, officeType, docType, and fileName are required.');
  }
  const folder = getDocFolder(district, officeType, docType);
  const existing = folder.getFilesByName(fileName);
  let deleted = 0;
  while (existing.hasNext()) { existing.next().setTrashed(true); deleted++; }
  return { ok: true, deleted: deleted };
}

// ─── AUTH: Admin OTP login, DDO/Inspector signup/login/password-reset ──
// Accounts are stored centrally in one "Accounts" tab (Users spreadsheet)
// so a DDO/Inspector can sign up once and log in from any device. OTP
// codes are stored in a temporary "OTP_Codes" tab with a 10-minute expiry
// and are consumed (deleted) on successful verification.

var ADMIN_EMAIL = 'himangshu9.deka9@gmail.com';

function getOtpSheet() {
  const ss = getOrCreateUsersSpreadsheet();
  return getOrCreateSheetTab(ss, 'OTP_Codes');
}

function getAccountsSheet() {
  const ss = getOrCreateUsersSpreadsheet();
  return getOrCreateSheetTab(ss, 'Accounts');
}

function generateOtp() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function storeOtp(key, code) {
  const sheet = getOtpSheet();
  const rows = readRows(sheet).filter(r => r.Key !== key);
  rows.push({ Key: key, Code: code, Expiry: String(Date.now() + 10 * 60 * 1000) });
  writeRows(sheet, rows);
}

function verifyOtpCode(key, code) {
  const sheet = getOtpSheet();
  const rows = readRows(sheet);
  const match = rows.find(r => r.Key === key && String(r.Code) === String(code));
  if (!match) return false;
  if (Number(match.Expiry) < Date.now()) return false;
  writeRows(sheet, rows.filter(r => r.Key !== key));
  return true;
}

function requestAdminOtp(params) {
  const code = generateOtp();
  storeOtp('ADMIN', code);
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: 'E-Cooperative Admin Login Code',
    body: 'Your E-Cooperative admin login code is: ' + code +
      '\n\nThis code expires in 10 minutes. If you did not request this, you can ignore this email.'
  });
  return { ok: true, sent: true };
}

function confirmAdminOtp(params) {
  const code = params.code;
  if (!code) throw new Error('code is required.');
  return { ok: true, valid: verifyOtpCode('ADMIN', code) };
}

function accountKey(district, officeType, role, phone) {
  return [district, officeType, role, phone].join('::');
}

function signupAccount(params) {
  const district = params.district, officeType = params.officeType, role = params.role,
    phone = params.phone, password = params.password, name = params.name, email = params.email || '';
  if (!district || !officeType || !role || !phone || !password || !name) {
    throw new Error('district, officeType, role, phone, password, and name are required.');
  }
  const sheet = getAccountsSheet();
  const rows = readRows(sheet);
  const key = accountKey(district, officeType, role, phone);
  if (rows.some(r => accountKey(r.District, r.OfficeType, r.Role, r.Phone) === key)) {
    return { ok: false, error: 'An account with this phone number already exists for this office.' };
  }
  rows.push({
    District: district, OfficeType: officeType, Role: role, Phone: String(phone),
    Password: password, Name: name, Email: email, DOB: '', Image: ''
  });
  writeRows(sheet, rows);
  return { ok: true };
}

function accountToClient(acct) {
  return {
    name: acct.Name, email: acct.Email, dob: acct.DOB, image: acct.Image,
    district: acct.District, officeType: acct.OfficeType, role: acct.Role, phone: acct.Phone,
    password: acct.Password
  };
}

function loginAccount(params) {
  const district = params.district, officeType = params.officeType, role = params.role,
    phone = params.phone, password = params.password;
  if (!district || !officeType || !role || !phone || !password) {
    throw new Error('district, officeType, role, phone, and password are required.');
  }
  const rows = readRows(getAccountsSheet());
  const acct = rows.find(r => r.District === district && r.OfficeType === officeType &&
    r.Role === role && String(r.Phone) === String(phone));
  if (!acct) return { ok: false, error: 'No account found.' };
  if (String(acct.Password) !== String(password)) return { ok: false, error: 'Incorrect password.' };
  return { ok: true, account: accountToClient(acct) };
}

// Looks up an account by phone + role + officeType across ALL districts, so the
// app's DDO/Inspector login (which has no district picker) can find the right
// district automatically on a new device.
function findAccount(params) {
  const role = params.role, officeType = params.officeType, phone = params.phone;
  if (!role || !officeType || !phone) {
    throw new Error('role, officeType, and phone are required.');
  }
  const rows = readRows(getAccountsSheet());
  const acct = rows.find(r => r.Role === role && r.OfficeType === officeType &&
    String(r.Phone) === String(phone));
  if (!acct) return { ok: true, account: null };
  return { ok: true, account: accountToClient(acct) };
}

function requestPasswordReset(params) {
  const district = params.district, officeType = params.officeType, role = params.role, phone = params.phone;
  if (!district || !officeType || !role || !phone) {
    throw new Error('district, officeType, role, and phone are required.');
  }
  const rows = readRows(getAccountsSheet());
  const acct = rows.find(r => r.District === district && r.OfficeType === officeType &&
    r.Role === role && String(r.Phone) === String(phone));
  if (!acct) return { ok: false, error: 'No account found.' };
  if (!acct.Email) {
    return { ok: false,
      error: 'No recovery email on file for this account. Ask your Admin/DDO to reset your password instead.' };
  }
  const code = generateOtp();
  storeOtp(accountKey(district, officeType, role, phone), code);
  MailApp.sendEmail({
    to: acct.Email,
    subject: 'E-Cooperative Password Reset Code',
    body: 'Your password reset code is: ' + code +
      '\n\nThis code expires in 10 minutes. If you did not request this, you can ignore this email.'
  });
  const email = String(acct.Email);
  const atIdx = email.indexOf('@');
  const masked = atIdx > 1 ? (email.slice(0, 2) + '***' + email.slice(atIdx)) : email;
  return { ok: true, emailHint: masked };
}

function confirmPasswordReset(params) {
  const district = params.district, officeType = params.officeType, role = params.role,
    phone = params.phone, code = params.code, newPassword = params.newPassword;
  if (!district || !officeType || !role || !phone || !code || !newPassword) {
    throw new Error('district, officeType, role, phone, code, and newPassword are required.');
  }
  const key = accountKey(district, officeType, role, phone);
  if (!verifyOtpCode(key, code)) return { ok: false, error: 'Invalid or expired code.' };
  const sheet = getAccountsSheet();
  const rows = readRows(sheet);
  const idx = rows.findIndex(r => accountKey(r.District, r.OfficeType, r.Role, r.Phone) === key);
  if (idx === -1) return { ok: false, error: 'Account not found.' };
  rows[idx].Password = newPassword;
  writeRows(sheet, rows);
  return { ok: true };
}

function updateAccountProfile(params) {
  const district = params.district, officeType = params.officeType, role = params.role, phone = params.phone;
  if (!district || !officeType || !role || !phone) {
    throw new Error('district, officeType, role, and phone are required.');
  }
  const sheet = getAccountsSheet();
  const rows = readRows(sheet);
  const idx = rows.findIndex(r => r.District === district && r.OfficeType === officeType &&
    r.Role === role && String(r.Phone) === String(phone));
  if (idx === -1) return { ok: false, error: 'Account not found.' };
  if (params.newPassword) {
    if (String(rows[idx].Password) !== String(params.currentPassword)) {
      return { ok: false, error: 'Current password is incorrect.' };
    }
    rows[idx].Password = params.newPassword;
  }
  if (params.name) rows[idx].Name = params.name;
  if (params.dob) rows[idx].DOB = params.dob;
  if (params.email) rows[idx].Email = params.email;
  if (params.image) rows[idx].Image = params.image;
  writeRows(sheet, rows);
  return { ok: true };
}

/**
 * ============================================================
 * DEPLOYMENT STEPS
 * ============================================================
 * 1. Go to script.google.com -> New project.
 * 2. Delete the default empty code, paste this entire file in.
 * 3. Save (Ctrl+S), name the project e.g. "E-Cooperative Backend".
 * 4. Click "Deploy" (top right) -> "New deployment".
 * 5. Click the gear icon next to "Select type" -> choose "Web app".
 * 6. Set:
 *      Execute as:        Me (your account)
 *      Who has access:    Anyone
 *    ("Anyone" just means anyone with the URL can call the endpoint --
 *    it does NOT give them access to your Drive/Sheets directly, only
 *    to whatever this script's functions expose. Do not add actions
 *    that return more than each district needs.)
 * 7. Click "Deploy". The first time, Google will ask you to authorize
 *    the script -- click through "Advanced" -> "Go to E-Cooperative
 *    Backend (unsafe)" if you see a warning (this is normal for your
 *    own scripts) and approve the Sheets/Drive permissions.
 * 8. Copy the "Web app URL" it gives you (ends in /exec).
 * 9. Paste that URL into the E-Cooperative app's GOOGLE_APPS_SCRIPT_URL
 *    constant (I'll wire this in next).
 *
 * Whenever you edit this script later, you must create a NEW deployment
 * (or "Manage deployments" -> edit -> new version) for changes to go live
 * -- editing the code alone does not update the already-deployed /exec URL.
 * ============================================================
 */
