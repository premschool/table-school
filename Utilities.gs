var SCRIPT_PROP = PropertiesService.getScriptProperties();
// Google Sheet ID ที่ผู้ใช้ระบุล่าสุด ใช้ร่วมกันทั้ง Setup และบริการข้อมูล
var SPREADSHEET_ID = '1lXCgkCg1XAZAdgGntnlcRZdfKVb1LLFrOQdwalCyNoA';
var DRIVE_FOLDER_ID = SCRIPT_PROP.getProperty('DRIVE_FOLDER_ID') || '';

var TEACHERS_SHEET = 'Teachers';
var SUBJECTS_SHEET = 'Subjects';
var SCHEDULES_SHEET = 'Schedules';
var LEAVE_REQUESTS_SHEET = 'LeaveRequests';
var SUBSTITUTE_ASSIGNMENTS_SHEET = 'SubstituteAssignments';
var PERIODS_SHEET = 'Periods';
var SETTINGS_SHEET = 'Settings';
var LOGS_SHEET = 'Logs';
var CLASSROOMS_SHEET = 'Classrooms';

var SESSION_TTL = 6 * 3600000;

// ==================== Cache ====================
var CACHE = CacheService.getScriptCache();
var CACHE_TTL = 300;
var CHUNK_SIZE = 90000;

function _chunkedCachePut(key, data, ttl) {
  try {
    var jsonStr = JSON.stringify(data);
    if (jsonStr.length <= CHUNK_SIZE) {
      CACHE.put(key, jsonStr, ttl || CACHE_TTL);
      CACHE.put(key + '_meta', '1', ttl || CACHE_TTL);
      return;
    }
    var totalChunks = Math.ceil(jsonStr.length / CHUNK_SIZE);
    var cacheObj = {};
    for (var i = 0; i < totalChunks; i++) {
      cacheObj[key + '_c' + i] = jsonStr.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }
    cacheObj[key + '_meta'] = String(totalChunks);
    CACHE.putAll(cacheObj, ttl || CACHE_TTL);
  } catch (e) { console.warn('Cache put error: ' + e.message); }
}

function _chunkedCacheGet(key) {
  try {
    var meta = CACHE.get(key + '_meta');
    if (!meta) return null;
    var totalChunks = parseInt(meta);
    if (totalChunks === 1) {
      var single = CACHE.get(key);
      return single ? JSON.parse(single) : null;
    }
    var chunkKeys = [];
    for (var i = 0; i < totalChunks; i++) chunkKeys.push(key + '_c' + i);
    var chunks = CACHE.getAll(chunkKeys);
    var jsonStr = '';
    for (var j = 0; j < totalChunks; j++) {
      var chunk = chunks[key + '_c' + j];
      if (!chunk) return null;
      jsonStr += chunk;
    }
    return JSON.parse(jsonStr);
  } catch (e) { return null; }
}

function _registerCacheKey(prefix, key) {
  var propKey = 'CACHE_KEYS_' + prefix;
  var keys = [];
  try { var ex = SCRIPT_PROP.getProperty(propKey); if (ex) keys = JSON.parse(ex); } catch (e) {}
  if (keys.indexOf(key) < 0) { keys.push(key); SCRIPT_PROP.setProperty(propKey, JSON.stringify(keys)); }
}

function _invalidateCache(prefix) {
  try {
    var keys = SCRIPT_PROP.getProperty('CACHE_KEYS_' + prefix);
    if (keys) {
      var keyList = JSON.parse(keys);
      var rm = [];
      for (var k = 0; k < keyList.length; k++) {
        rm.push(keyList[k]); rm.push(keyList[k] + '_meta');
        for (var c = 0; c < 20; c++) rm.push(keyList[k] + '_c' + c);
      }
      CACHE.removeAll(rm);
      SCRIPT_PROP.deleteProperty('CACHE_KEYS_' + prefix);
    }
    var dr = [prefix, prefix + '_meta'];
    for (var d = 0; d < 20; d++) dr.push(prefix + '_c' + d);
    CACHE.removeAll(dr);
  } catch (e) { console.warn('Invalidate error: ' + e.message); }
}

// ==================== Sheet Access ====================
function getSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    teachers: ss.getSheetByName(TEACHERS_SHEET),
    subjects: ss.getSheetByName(SUBJECTS_SHEET),
    schedules: ss.getSheetByName(SCHEDULES_SHEET),
    leaveRequests: ss.getSheetByName(LEAVE_REQUESTS_SHEET),
    assignments: ss.getSheetByName(SUBSTITUTE_ASSIGNMENTS_SHEET),
    periods: ss.getSheetByName(PERIODS_SHEET),
    settings: ss.getSheetByName(SETTINGS_SHEET),
    logs: ss.getSheetByName(LOGS_SHEET),
    classrooms: ss.getSheetByName(CLASSROOMS_SHEET),
    ss: ss
  };
}

function getSheetData(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getDataRange().getValues();
}

function getColumnMap(sheetName) {
  var cacheKey = 'map_' + SPREADSHEET_ID + '_' + sheetName;
  var cached = _chunkedCacheGet(cacheKey);
  if (cached) return cached;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {};
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) { if (headers[i]) map[headers[i]] = i; }
    _chunkedCachePut(cacheKey, map, 3600);
    return map;
  } catch (e) { return {}; }
}

// ==================== Helpers ====================
function _buildHeaderIndex(headers) {
  var idx = {};
  for (var i = 0; i < headers.length; i++) { if (headers[i]) idx[headers[i]] = i; }
  return idx;
}

function _rowToObj(row, headers) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    var val = row[i];
    obj[headers[i]] = (val instanceof Date) ? formatDateForSheet(val) : val;
  }
  return obj;
}

function generateId(prefix) {
  return (prefix || 'id_') + new Date().getTime() + Math.floor(Math.random() * 1000);
}

function formatDateForSheet(dateObj) {
  if (!dateObj) return '';
  var d = (typeof dateObj === 'string' || typeof dateObj === 'number') ? new Date(dateObj) : dateObj;
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDateTimeForSheet(dateObj) {
  if (!dateObj) return '';
  var d = (typeof dateObj === 'string' || typeof dateObj === 'number') ? new Date(dateObj) : dateObj;
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function getDayOfWeekFromDate(dateString) {
  var d = new Date(dateString + 'T00:00:00');
  return ["อาทิตย์","จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์"][d.getDay()];
}

function findRowIndexById(sheetName, id, idColumnName) {
  var data = getSheetData(sheetName);
  var map = getColumnMap(sheetName);
  if (!map.hasOwnProperty(idColumnName)) return -1;
  var col = map[idColumnName];
  for (var i = 1; i < data.length; i++) {
    if (data[i][col] && data[i][col].toString() === id.toString()) return i + 1;
  }
  return -1;
}

// ==================== Security ====================
function _generateSalt() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

// PBKDF2-HMAC-SHA256, 6000 iterations. Sheet contains only version, rounds, salt and derived key.
var PASSWORD_KDF_ROUNDS = 6000;
function _hashPw(pw, salt) {
  return _hashPwRounds_(pw, salt || _generateSalt(), PASSWORD_KDF_ROUNDS);
}
function _hashPwRounds_(pw, salt, rounds) {
  var key = Utilities.newBlob(String(pw)).getBytes();
  var saltBlock = Utilities.newBlob(String(salt) + '\x00\x00\x00\x01').getBytes();
  var u = Utilities.computeHmacSha256Signature(saltBlock, key);
  var t = u.slice();
  for (var i = 1; i < rounds; i++) {
    u = Utilities.computeHmacSha256Signature(u, key);
    for (var j = 0; j < t.length; j++) t[j] = t[j] ^ u[j];
  }
  var hex = t.map(function(b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
  return 'pbkdf2-sha256$' + rounds + '$' + salt + '$' + hex;
}

function _verifyPassword(inputPw, storedPw) {
  if (!storedPw) return false;
  var parts = String(storedPw).split('$');
  if (parts.length === 4 && parts[0] === 'pbkdf2-sha256') {
    var rounds = Number(parts[1]);
    if (!rounds || rounds > 100000 || rounds < 1000) return false;
    var candidate = _hashPwRounds_(inputPw, parts[2], rounds).split('$')[3];
    return _constantTimeEquals_(candidate, parts[3]);
  }
  // Backward compatibility: old hash:salt credentials are upgraded after successful login.
  var legacy = String(storedPw).split(':');
  if (legacy.length !== 2) return false;
  var raw = Utilities.computeHmacSha256Signature(String(inputPw), legacy[1]);
  var oldHash = raw.map(function(b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
  return _constantTimeEquals_(oldHash, legacy[0]);
}

function _constantTimeEquals_(a, b) {
  a = String(a || ''); b = String(b || '');
  var diff = a.length ^ b.length, n = Math.max(a.length, b.length);
  for (var i = 0; i < n; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function _roleHasTeacher_(role) {
  var roles = String(role || '').split(',');
  for (var i = 0; i < roles.length; i++) if (roles[i].trim() === 'Teacher') return true;
  return false;
}

function _passwordValidationMessage_(password, role) {
  var roles = String(role || '').split(',').map(function(r) { return r.trim(); });
  var privileged = roles.indexOf('Admin') >= 0 || roles.indexOf('Director') >= 0 || roles.indexOf('Scheduler') >= 0;
  if (_roleHasTeacher_(role) && !privileged) return /^\d{6}$/.test(String(password || '')) ? '' : 'PIN ของครูต้องเป็นตัวเลข 6 หลัก';
  return String(password || '').length >= 10 ? '' : 'รหัสผ่านต้องมีอย่างน้อย 10 ตัวอักษร';
}

// จำกัดการลอง PIN/รหัสผ่านต่อชื่อผู้ใช้ 5 ครั้งในช่วง 15 นาที
function _loginAttemptCacheKey_(userId) {
  var normalized = String(userId || '').trim().toLowerCase();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalized);
  var hex = digest.map(function(b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
  return 'login_fail_' + hex;
}
function _loginIsThrottled_(userId) {
  var attempts = Number(CACHE.get(_loginAttemptCacheKey_(userId)) || 0);
  return attempts >= 5;
}
function _recordLoginFailure_(userId) {
  var key = _loginAttemptCacheKey_(userId);
  var attempts = Number(CACHE.get(key) || 0) + 1;
  CACHE.put(key, String(attempts), 900);
}
function _clearLoginFailures_(userId) {
  CACHE.remove(_loginAttemptCacheKey_(userId));
}

function _sanitize(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ==================== Logging ====================
function logAction(action, details, user) {
  try {
    // ตรวจ enable_logs จาก Settings (cached)
    var enableLogs = CACHE.get('enable_logs');
    if (enableLogs === null) {
      var settingsResult = DataService.getSettings();
      var s = settingsResult.data || {};
      enableLogs = (String(s.enable_logs || '').toLowerCase() === 'false') ? 'false' : 'true';
      CACHE.put('enable_logs', enableLogs, 300);
    }
    if (enableLogs === 'false') return;

    var sheets = getSheets();
    if (!sheets.logs) return;
    sheets.logs.appendRow([formatDateTimeForSheet(new Date()), action, user || 'system', details || '']);
  } catch (e) { console.error("Log error: " + e.toString()); }
}

// ==================== RBAC ====================
function validateRequestorRole_(requestorId, allowedRoles) {
  if (!requestorId) throw new Error('Permission Denied: No requestor ID.');
  var result = DataService.getTeachers();
  var teachers = result.data || [];
  var roleStr = null;
  for (var i = 0; i < teachers.length; i++) {
    if (teachers[i].teacher_id === requestorId) { roleStr = teachers[i].role; break; }
  }
  if (!roleStr) throw new Error('Permission Denied: User not found.');
  var userRoles = roleStr.split(',').map(function(r) { return r.trim(); });
  for (var j = 0; j < userRoles.length; j++) {
    if (allowedRoles.indexOf(userRoles[j]) >= 0) return true;
  }
  throw new Error('Permission Denied: Insufficient role.');
}

// ==================== Triggers ====================
function onSheetEdit(e) {
  if (!e) return;
  try {
    var name = e.source.getSheetName();
    var mk = 'map_' + SPREADSHEET_ID + '_' + name;
    if (name === TEACHERS_SHEET) { _invalidateCache('teachers'); CACHE.remove(mk); }
    else if (name === SCHEDULES_SHEET) { _invalidateCache('schedules'); CACHE.remove(mk); }
    else if (name === SUBJECTS_SHEET) { _invalidateCache('subjects'); CACHE.remove(mk); }
    else if (name === LEAVE_REQUESTS_SHEET) { _invalidateCache('leaves'); CACHE.remove(mk); }
    else if (name === SUBSTITUTE_ASSIGNMENTS_SHEET) { _invalidateCache('assignments'); CACHE.remove(mk); }
    else if (name === PERIODS_SHEET) { _invalidateCache('periods'); CACHE.remove(mk); }
    else if (name === SETTINGS_SHEET) { _invalidateCache('settings'); }
  } catch (err) { console.error('onSheetEdit: ' + err.toString()); }
}

function cleanupExpiredSessions() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = new Date().getTime(), deleted = 0;
  Object.keys(all).forEach(function(key) {
    if (key.indexOf('sess_') === 0) {
      try {
        var d = JSON.parse(all[key]);
        if (d.expiry && now > d.expiry) { props.deleteProperty(key); deleted++; }
      } catch (e) { props.deleteProperty(key); deleted++; }
    }
  });
  Logger.log('Cleaned ' + deleted + ' expired sessions');
}
