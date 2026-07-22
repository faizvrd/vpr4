/**
 * Config.gs
 * Central configuration for the Vendor Progress Report (VPR) Web App.
 *
 * SETUP:
 * 1. Import database/VPR_Database.xlsx into Google Sheets
 *    (Google Sheets > File > Import > Upload > "Insert new sheet(s)").
 * 2. Copy that Google Sheet's ID (from its URL) into SPREADSHEET_ID below,
 *    OR bind this Apps Script project directly to that Sheet
 *    (Extensions > Apps Script) and leave SPREADSHEET_ID as "".
 * 3. CHANGE AUTH_SECRET below to your own random string (used to sign
 *    login sessions). Anyone who knows this value could forge a login
 *    token, so treat it like a password.
 */

var SPREADSHEET_ID = ''; // e.g. '1AbCDefGhIJKLmnopQRstuVWxyz...'  Leave blank if bound to the Sheet.

var AUTH_SECRET = 'CHANGE-ME-TO-A-LONG-RANDOM-STRING'; // used to sign session tokens

var ROLES = {
  ADMIN: 'Admin',   // full access to everything
  VENDOR: 'Vendor', // sees & edits progress only for their own registered vendor
  USER: 'User'      // view-only, everything
};

var SHEETS = {
  VENDORS: 'Vendors',
  PURCHASE_ORDERS: 'PurchaseOrders',
  MILESTONES: 'Milestones',
  SCURVE: 'SCurve',
  DOCUMENTS: 'Documents',
  SUBVENDOR: 'SubVendor',
  USERS: 'Users'
};
/**
 * Database.gs
 * Generic CRUD engine on top of Google Sheets. Every sheet's first row is
 * the header row; each header becomes an object key. All Service files
 * (VendorService.gs, MilestoneService.gs, ...) call these helpers instead
 * of touching SpreadsheetApp directly.
 *
 * PERFORMANCE: Sheets API calls (open spreadsheet, read a range, write a
 * cell) are the slow part of every request — each one is a network round
 * trip. This file caches the spreadsheet handle, header rows, and full
 * sheet reads for the lifetime of a single request (see resetCache_(),
 * called once at the top of handleRequest_ in Code.gs) so that, e.g.,
 * computing weighted progress for 20 POs reads the Milestones sheet ONCE
 * instead of 20 times. Writes go through a single batched setValues() call
 * per row instead of one setValue() call per cell, and invalidate the
 * cache for that sheet so later reads in the same request see fresh data.
 */

var _cache_ = null;

function resetCache_() {
  _cache_ = { spreadsheet: null, sheets: {}, headers: {}, rows: {} };
}

function getDb_() {
  if (!_cache_) resetCache_();
  if (!_cache_.spreadsheet) {
    _cache_.spreadsheet = SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
  }
  return _cache_.spreadsheet;
}

function getSheet_(sheetName) {
  if (!_cache_) resetCache_();
  if (!_cache_.sheets[sheetName]) {
    var sheet = getDb_().getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet not found: ' + sheetName);
    _cache_.sheets[sheetName] = sheet;
  }
  return _cache_.sheets[sheetName];
}

function getHeaders_(sheet) {
  var name = sheet.getName();
  if (!_cache_.headers[name]) {
    var lastCol = sheet.getLastColumn();
    _cache_.headers[name] = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  return _cache_.headers[name];
}

function invalidateCache_(sheetName) {
  delete _cache_.rows[sheetName];
}

/** Read every data row of a sheet as an array of plain objects. Cached per sheet per request. */
function readAll_(sheetName) {
  if (!_cache_) resetCache_();
  if (_cache_.rows[sheetName]) return _cache_.rows[sheetName];

  var sheet = getSheet_(sheetName);
  var lastRow = sheet.getLastRow();
  var headers = getHeaders_(sheet);
  var out = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      // skip fully blank rows
      if (row.every(function (v) { return v === '' || v === null; })) continue;
      var obj = { _row: r + 2 };
      for (var c = 0; c < headers.length; c++) {
        obj[headers[c]] = normalizeValue_(row[c]);
      }
      out.push(obj);
    }
  }
  _cache_.rows[sheetName] = out;
  return out;
}

function normalizeValue_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  return v;
}

/** Find one row by an id column (e.g. 'POID'). Returns object or null. */
function findById_(sheetName, idField, idValue) {
  var rows = readAll_(sheetName);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][idField]) === String(idValue)) return rows[i];
  }
  return null;
}

/** Filter rows where filterField === filterValue. */
function findWhere_(sheetName, filterField, filterValue) {
  return readAll_(sheetName).filter(function (row) {
    return String(row[filterField]) === String(filterValue);
  });
}

/** Insert a new row built from a plain object matching the header names. */
function insertRow_(sheetName, obj) {
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var row = headers.map(function (h) {
    return obj.hasOwnProperty(h) && obj[h] !== undefined ? obj[h] : '';
  });
  sheet.appendRow(row);
  invalidateCache_(sheetName);
  return obj;
}

/** Update an existing row (matched by idField/idValue) with new field values.
 *  Uses the request cache to find the row number (no extra read) and writes
 *  the whole row in a single setValues() call (no per-cell round trips). */
function updateRow_(sheetName, idField, idValue, patch) {
  var existing = findById_(sheetName, idField, idValue);
  if (!existing) return false;
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var rowValues = headers.map(function (h) {
    if (patch.hasOwnProperty(h)) return patch[h];
    return existing.hasOwnProperty(h) ? existing[h] : '';
  });
  sheet.getRange(existing._row, 1, 1, headers.length).setValues([rowValues]);
  invalidateCache_(sheetName);
  return true;
}

/** Delete a row matched by idField/idValue. Uses the request cache to find
 *  the row number directly, avoiding an extra column read. */
function deleteRow_(sheetName, idField, idValue) {
  var existing = findById_(sheetName, idField, idValue);
  if (!existing) return false;
  getSheet_(sheetName).deleteRow(existing._row);
  invalidateCache_(sheetName);
  return true;
}

/** Generate the next sequential id like "PO-0007" for a sheet/idField/prefix. */
function nextId_(sheetName, idField, prefix) {
  var rows = readAll_(sheetName);
  var max = 0;
  rows.forEach(function (row) {
    var raw = String(row[idField] || '');
    var num = parseInt(raw.replace(prefix + '-', ''), 10);
    if (!isNaN(num) && num > max) max = num;
  });
  var next = max + 1;
  var padded = ('0000' + next).slice(-4);
  return prefix + '-' + padded;
}
/**
 * AuthService.gs
 * Username/password login against the Users sheet, stateless signed
 * session tokens (no server-side session storage needed), and role-based
 * authorization helpers used by the API router in Code.gs.
 *
 * Roles:
 *  - Admin  : full access to everything (vendors, POs, all progress data, user management)
 *  - Vendor : only sees/edits progress data for POs belonging to their own VendorID
 *  - User   : view-only access to everything
 */

/* ---------------------------------------------------------------- LOGIN */

function login(username, password) {
  if (!username || !password) throw new Error('Username dan password wajib diisi');
  var user = findById_(SHEETS.USERS, 'Username', username);
  if (!user) throw new Error('Username atau password salah');
  var hash = hashPassword_(password, user.PasswordSalt);
  if (hash !== user.PasswordHash) throw new Error('Username atau password salah');

  var payload = {
    u: user.Username,
    r: user.Role,
    v: user.VendorID || '',
    n: user.Name || user.Username,
    exp: Date.now() + 12 * 60 * 60 * 1000 // 12 hour session
  };
  return {
    token: signToken_(payload),
    username: payload.u,
    role: payload.r,
    vendorId: payload.v,
    name: payload.n
  };
}

/* ---------------------------------------------------------------- TOKEN (mini signed JWT-like) */

function signToken_(payload) {
  var payloadStr = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
  return payloadStr + '.' + signPart_(payloadStr);
}

function signPart_(str) {
  var raw = Utilities.computeHmacSha256Signature(str, AUTH_SECRET);
  return Utilities.base64EncodeWebSafe(raw);
}

function verifyToken_(token) {
  if (!token) throw new Error('Silakan login terlebih dahulu');
  var parts = String(token).split('.');
  if (parts.length !== 2) throw new Error('Token tidak valid, silakan login ulang');
  if (signPart_(parts[0]) !== parts[1]) throw new Error('Token tidak valid, silakan login ulang');
  var payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  if (!payload.exp || payload.exp < Date.now()) throw new Error('Sesi berakhir, silakan login ulang');
  return payload; // { u, r, v, n, exp }
}

/* ---------------------------------------------------------------- PASSWORD HASHING */

function hashPassword_(password, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + ':' + salt);
  return raw.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

/* ---------------------------------------------------------------- AUTHORIZATION GUARDS */

function requireAdmin_(auth) {
  if (auth.r !== ROLES.ADMIN) throw new Error('Akses ditolak: fitur ini khusus Admin');
}

/** Admin & User (viewer) can view any PO; Vendor can only view their own. */
function requireCanViewPO_(auth, poId) {
  if (auth.r === ROLES.ADMIN || auth.r === ROLES.USER) return;
  if (auth.r === ROLES.VENDOR) {
    var po = findById_(SHEETS.PURCHASE_ORDERS, 'POID', poId);
    if (po && po.VendorID === auth.v) return;
    throw new Error('Akses ditolak: PO ini bukan milik vendor Anda');
  }
  throw new Error('Akses ditolak');
}

/** Admin can edit any PO's progress data; Vendor only their own; User (viewer) can never edit. */
function requireCanEditPO_(auth, poId) {
  if (auth.r === ROLES.ADMIN) return;
  if (auth.r === ROLES.VENDOR) {
    var po = findById_(SHEETS.PURCHASE_ORDERS, 'POID', poId);
    if (po && po.VendorID === auth.v) return;
    throw new Error('Akses ditolak: PO ini bukan milik vendor Anda');
  }
  throw new Error('Akses ditolak: akun Anda hanya bisa melihat (view only)');
}

/* ---------------------------------------------------------------- USER MANAGEMENT (Admin only) */

function listUsers() {
  return readAll_(SHEETS.USERS).map(function (u) {
    return { Username: u.Username, Role: u.Role, VendorID: u.VendorID, Name: u.Name, CreatedAt: u.CreatedAt };
  });
}

/** Create or update a login account. Payload: { Username, Password?, Role, VendorID?, Name }. */
function saveUserAccount(payload) {
  if (!payload.Username) throw new Error('Username wajib diisi');
  if (payload.Role === ROLES.VENDOR && !payload.VendorID) throw new Error('Akun Vendor wajib dikaitkan ke salah satu Vendor');

  var existing = findById_(SHEETS.USERS, 'Username', payload.Username);
  var salt = existing ? existing.PasswordSalt : Utilities.getUuid();
  var hash;
  if (payload.Password) {
    hash = hashPassword_(payload.Password, salt);
  } else if (existing) {
    hash = existing.PasswordHash;
  } else {
    throw new Error('Password wajib diisi untuk akun baru');
  }

  var row = {
    Username: payload.Username,
    PasswordHash: hash,
    PasswordSalt: salt,
    Role: payload.Role,
    VendorID: payload.Role === ROLES.VENDOR ? payload.VendorID : '',
    Name: payload.Name || payload.Username,
    CreatedAt: existing ? existing.CreatedAt : new Date()
  };

  if (existing) {
    updateRow_(SHEETS.USERS, 'Username', payload.Username, row);
  } else {
    insertRow_(SHEETS.USERS, row);
  }
  return payload.Username;
}

function deleteUserAccount(username) {
  return deleteRow_(SHEETS.USERS, 'Username', username);
}

/* ---------------------------------------------------------------- Lookups used by authorization guards */

function getMilestonePOID_(id) { var r = findById_(SHEETS.MILESTONES, 'MilestoneID', id); return r ? r.POID : null; }
function getSCurvePOID_(id) { var r = findById_(SHEETS.SCURVE, 'SCurveID', id); return r ? r.POID : null; }
function getDocumentPOID_(id) { var r = findById_(SHEETS.DOCUMENTS, 'DocID', id); return r ? r.POID : null; }
function getSubVendorItemPOID_(id) { var r = findById_(SHEETS.SUBVENDOR, 'ItemID', id); return r ? r.POID : null; }
/**
 * VendorService.gs
 * CRUD for Vendors and their Purchase Orders (PO). A Purchase Order is the
 * unit that all progress data (Milestones, SCurve, Documents, SubVendor)
 * hangs off of, via POID.
 */

function listVendors() {
  return readAll_(SHEETS.VENDORS);
}

function saveVendor(vendor) {
  if (vendor.VendorID) {
    updateRow_(SHEETS.VENDORS, 'VendorID', vendor.VendorID, vendor);
    return vendor.VendorID;
  }
  vendor.VendorID = nextId_(SHEETS.VENDORS, 'VendorID', 'VEN');
  vendor.CreatedAt = new Date();
  insertRow_(SHEETS.VENDORS, vendor);
  return vendor.VendorID;
}

function deleteVendor(vendorId) {
  // also cascade-delete its POs and everything under them
  var pos = findWhere_(SHEETS.PURCHASE_ORDERS, 'VendorID', vendorId);
  pos.forEach(function (po) { deletePO(po.POID); });
  return deleteRow_(SHEETS.VENDORS, 'VendorID', vendorId);
}

function listPOs(vendorId) {
  var pos = vendorId
    ? findWhere_(SHEETS.PURCHASE_ORDERS, 'VendorID', vendorId)
    : readAll_(SHEETS.PURCHASE_ORDERS);
  var vendors = listVendors();
  var vMap = {};
  vendors.forEach(function (v) { vMap[v.VendorID] = v.VendorName; });
  pos.forEach(function (po) { po.VendorName = vMap[po.VendorID] || ''; });
  return pos;
}

function getPO(poId) {
  var po = findById_(SHEETS.PURCHASE_ORDERS, 'POID', poId);
  if (po) {
    var vendor = findById_(SHEETS.VENDORS, 'VendorID', po.VendorID);
    po.VendorName = vendor ? vendor.VendorName : '';
  }
  return po;
}

function savePO(po) {
  if (po.POID) {
    updateRow_(SHEETS.PURCHASE_ORDERS, 'POID', po.POID, po);
    return po.POID;
  }
  po.POID = nextId_(SHEETS.PURCHASE_ORDERS, 'POID', 'PO');
  po.CreatedAt = new Date();
  insertRow_(SHEETS.PURCHASE_ORDERS, po);
  return po.POID;
}

function deletePO(poId) {
  findWhere_(SHEETS.MILESTONES, 'POID', poId).forEach(function (m) {
    deleteRow_(SHEETS.MILESTONES, 'MilestoneID', m.MilestoneID);
  });
  findWhere_(SHEETS.SCURVE, 'POID', poId).forEach(function (s) {
    deleteRow_(SHEETS.SCURVE, 'SCurveID', s.SCurveID);
  });
  findWhere_(SHEETS.DOCUMENTS, 'POID', poId).forEach(function (d) {
    deleteRow_(SHEETS.DOCUMENTS, 'DocID', d.DocID);
  });
  findWhere_(SHEETS.SUBVENDOR, 'POID', poId).forEach(function (i) {
    deleteRow_(SHEETS.SUBVENDOR, 'ItemID', i.ItemID);
  });
  return deleteRow_(SHEETS.PURCHASE_ORDERS, 'POID', poId);
}
/**
 * MilestoneService.gs
 * CRUD for milestone/progress rows (VPR Summary & Detail sheets), plus the
 * weighted total-progress calculation (Sum(WeightFactor * AchievedPercent / 100)).
 */

function listMilestones(poId) {
  var rows = findWhere_(SHEETS.MILESTONES, 'POID', poId);
  rows.sort(function (a, b) { return a._row - b._row; });
  return rows;
}

function saveMilestone(m) {
  m.WeightFactor = Number(m.WeightFactor) || 0;
  m.LastProgress = Number(m.LastProgress) || 0;
  m.ThisProgress = Number(m.ThisProgress) || 0;
  m.AchievedPercent = Math.min(100, Number(m.LastProgress) + Number(m.ThisProgress));

  if (m.MilestoneID) {
    updateRow_(SHEETS.MILESTONES, 'MilestoneID', m.MilestoneID, m);
    return m;
  }
  m.MilestoneID = nextId_(SHEETS.MILESTONES, 'MilestoneID', 'MS');
  return insertRow_(SHEETS.MILESTONES, m);
}

function deleteMilestone(id) {
  return deleteRow_(SHEETS.MILESTONES, 'MilestoneID', id);
}

/** Weighted overall progress (0-100) for a PO, based on Milestones rows. */
function getPOProgress(poId) {
  var rows = listMilestones(poId);
  var totalWeight = 0;
  var earned = 0;
  rows.forEach(function (m) {
    var w = Number(m.WeightFactor) || 0;
    var pct = Number(m.AchievedPercent) || 0;
    totalWeight += w;
    earned += (w * pct) / 100;
  });
  return {
    totalWeight: totalWeight,
    earnedWeight: Math.round(earned * 100) / 100,
    percent: totalWeight ? Math.round((earned / totalWeight) * 10000) / 100 : 0
  };
}
/**
 * SCurveService.gs
 * CRUD for the Cumulative Progress (S-Curve) data points: Plan % vs Actual %
 * over a sequence of periods for a given PO.
 */

function listSCurve(poId) {
  var rows = findWhere_(SHEETS.SCURVE, 'POID', poId);
  rows.sort(function (a, b) { return a._row - b._row; });
  return rows;
}

function saveSCurvePoint(s) {
  s.PlanPercent = Number(s.PlanPercent) || 0;
  s.ActualPercent = s.ActualPercent === '' || s.ActualPercent === undefined
    ? ''
    : Number(s.ActualPercent);

  if (s.SCurveID) {
    updateRow_(SHEETS.SCURVE, 'SCurveID', s.SCurveID, s);
    return s;
  }
  s.SCurveID = nextId_(SHEETS.SCURVE, 'SCurveID', 'SC');
  return insertRow_(SHEETS.SCURVE, s);
}

function deleteSCurvePoint(id) {
  return deleteRow_(SHEETS.SCURVE, 'SCurveID', id);
}
/**
 * DocumentService.gs
 * CRUD for the "Documents Status" tracking sheet (submittals / approvals).
 */

function listDocuments(poId) {
  var rows = findWhere_(SHEETS.DOCUMENTS, 'POID', poId);
  rows.sort(function (a, b) { return a._row - b._row; });
  return rows;
}

function saveDocument(d) {
  if (d.DocID) {
    updateRow_(SHEETS.DOCUMENTS, 'DocID', d.DocID, d);
    return d;
  }
  d.DocID = nextId_(SHEETS.DOCUMENTS, 'DocID', 'DOC');
  return insertRow_(SHEETS.DOCUMENTS, d);
}

function deleteDocument(id) {
  return deleteRow_(SHEETS.DOCUMENTS, 'DocID', id);
}
/**
 * SubVendorService.gs
 * CRUD for the "Sub-Vendor / Bought-out Item" material procurement sheet.
 */

function listSubVendorItems(poId) {
  var rows = findWhere_(SHEETS.SUBVENDOR, 'POID', poId);
  rows.sort(function (a, b) { return a._row - b._row; });
  return rows;
}

function saveSubVendorItem(i) {
  i.QtyRequired = Number(i.QtyRequired) || 0;
  i.ArrivalQty = Number(i.ArrivalQty) || 0;
  i.Balance = Math.max(0, i.QtyRequired - i.ArrivalQty);

  if (i.ItemID) {
    updateRow_(SHEETS.SUBVENDOR, 'ItemID', i.ItemID, i);
    return i;
  }
  i.ItemID = nextId_(SHEETS.SUBVENDOR, 'ItemID', 'ITM');
  return insertRow_(SHEETS.SUBVENDOR, i);
}

function deleteSubVendorItem(id) {
  return deleteRow_(SHEETS.SUBVENDOR, 'ItemID', id);
}
/**
 * Code.gs
 * JSON API entry point. The UI (Index.html) is hosted externally (e.g. GitHub
 * Pages) and talks to this Apps Script Web App over fetch(). This file does
 * NOT serve HTML anymore — it only returns JSON.
 *
 * Deploy: Deploy > New deployment > Web app > Execute as: Me, Access: Anyone.
 * Copy the resulting /exec URL into API_URL at the top of Index.html.
 */

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

/** Parses action/payload from either a GET query string or a POST body, routes it, and returns JSON. */
function handleRequest_(e) {
  resetCache_();
  var action, payload;
  try {
    if (e.postData && e.postData.contents) {
      var body = JSON.parse(e.postData.contents);
      action = body.action;
      payload = body.payload || {};
    } else {
      action = e.parameter.action;
      payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    }
    var data = route_(action, payload);
    return jsonOutput_({ ok: true, data: data });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Single dispatch table mapping API action names to backend functions. Every
 *  action except 'login' requires a valid session token (payload.token). */
function route_(action, payload) {
  if (action === 'login') return login(payload.username, payload.password);

  var auth = verifyToken_(payload.token);

  switch (action) {
    case 'bootstrap':            return bootstrapData(auth);
    case 'getWorkspace':         requireCanViewPO_(auth, payload.poId); return getPOWorkspace(payload.poId);

    case 'saveVendor':           requireAdmin_(auth); return saveVendor(payload);
    case 'deleteVendor':         requireAdmin_(auth); return deleteVendor(payload.vendorId);

    case 'savePO':                requireAdmin_(auth); return savePO(payload);
    case 'deletePO':              requireAdmin_(auth); return deletePO(payload.poId);

    case 'saveMilestone':         requireCanEditPO_(auth, payload.POID); return saveMilestone(payload);
    case 'deleteMilestone':       requireCanEditPO_(auth, getMilestonePOID_(payload.id)); return deleteMilestone(payload.id);

    case 'saveSCurvePoint':       requireCanEditPO_(auth, payload.POID); return saveSCurvePoint(payload);
    case 'deleteSCurvePoint':     requireCanEditPO_(auth, getSCurvePOID_(payload.id)); return deleteSCurvePoint(payload.id);

    case 'saveDocument':          requireCanEditPO_(auth, payload.POID); return saveDocument(payload);
    case 'deleteDocument':        requireCanEditPO_(auth, getDocumentPOID_(payload.id)); return deleteDocument(payload.id);

    case 'saveSubVendorItem':     requireCanEditPO_(auth, payload.POID); return saveSubVendorItem(payload);
    case 'deleteSubVendorItem':   requireCanEditPO_(auth, getSubVendorItemPOID_(payload.id)); return deleteSubVendorItem(payload.id);

    case 'listUsers':             requireAdmin_(auth); return listUsers();
    case 'saveUserAccount':       requireAdmin_(auth); return saveUserAccount(payload);
    case 'deleteUserAccount':     requireAdmin_(auth); return deleteUserAccount(payload.username);

    default: throw new Error('Unknown action: ' + action);
  }
}

/** Everything the dashboard needs on first load, scoped by role:
 *  Admin/User see every vendor & PO; Vendor sees only their own. */
function bootstrapData(auth) {
  var vendors = listVendors();
  var pos = listPOs();
  if (auth.r === ROLES.VENDOR) {
    vendors = vendors.filter(function (v) { return v.VendorID === auth.v; });
    pos = pos.filter(function (po) { return po.VendorID === auth.v; });
  }
  pos.forEach(function (po) { po.progress = getPOProgress(po.POID).percent; });
  return { vendors: vendors, pos: pos };
}

/** Everything needed to render one PO's full workspace (all 5 tabs). */
function getPOWorkspace(poId) {
  return {
    po: getPO(poId),
    milestones: listMilestones(poId),
    progress: getPOProgress(poId),
    scurve: listSCurve(poId),
    documents: listDocuments(poId),
    subvendor: listSubVendorItems(poId)
  };
}
