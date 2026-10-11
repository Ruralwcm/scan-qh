// =============================================================================
// Scan điểm QH v1.0 — Google Apps Script backend
// Sheet: 1fMmEm1GC9aei_3oeu2vcadPVsUM9ms9eVhm99d7pF8A
// Tab: KhaoSat_QH
// =============================================================================

const SHEET_ID = '1fMmEm1GC9aei_3oeu2vcadPVsUM9ms9eVhm99d7pF8A';
const USERS_TAB = '5-Users';  // reuse CTKM user tab
const SURVEY_TAB = 'KhaoSat_QH';
const TOKEN_EXPIRY_MS = 24 * 60 * 60 * 1000;

// =============================================================================
// HELPERS
// =============================================================================
function _hashPassword(pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8)
    .map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function _generateToken() {
  return Utilities.getUuid() + '_' + Date.now();
}

function _tokenStore() {
  var props = PropertiesService.getScriptProperties();
  return {
    save: function(token, data) { props.setProperty('tk_' + token, JSON.stringify(data)); },
    get: function(token) {
      var raw = props.getProperty('tk_' + token);
      if (!raw) return null;
      try {
        var data = JSON.parse(raw);
        if (Date.now() - data._ts > TOKEN_EXPIRY_MS) { props.deleteProperty('tk_' + token); return null; }
        return data;
      } catch(e) { return null; }
    },
    delete: function(token) { props.deleteProperty('tk_' + token); }
  };
}

function _getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#1F4E79').setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function _json(objOrArr) {
  return ContentService.createTextOutput(JSON.stringify(objOrArr))
    .setMimeType(ContentService.MimeType.JSON);
}

function _err(msg) {
  return _json({ ok: false, error: msg });
}

// =============================================================================
// ROUTER
// =============================================================================
function doPost(e) {
  try {
    var data = e && e.postData ? JSON.parse(e.postData.contents) : {};
    var action = data.action || '';
    var token = data.token || '';
    var user = null;

    if (action !== 'login' && action !== 'ping') {
      user = _tokenStore().get(token);
      if (!user) return _err('Hết phiên đăng nhập. Vui lòng đăng nhập lại.');
    }

    switch (action) {
      case 'ping': return _json({ ok: true, ts: new Date().toISOString() });
      case 'login': return _login(data);
      case 'saveSurvey': return _saveSurvey(user, data);
      case 'getHistory': return _getHistory(data);
      case 'getDashboard': return _getDashboard(data);
      case 'listUsers': return _listUsers(user);
      default: return _err('Unknown action: ' + action);
    }
  } catch (err) {
    return _err('Server error: ' + err.toString());
  }
}

function doGet(e) {
  return _json({ ok: true, service: 'Scan-QH-API v1.0' });
}

// =============================================================================
// AUTH — reuses "5-Users" tab from CTKM
// =============================================================================
function _login(data) {
  var username = (data.username || '').toString().trim();
  var password = (data.password || '').toString().trim();
  if (!username || !password) return _err('Vui lòng nhập tên đăng nhập và mật khẩu');

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(USERS_TAB);
  if (!sheet) return _err('Chưa có bảng users. Vui lòng tạo tab "5-Users".');

  var rows = sheet.getDataRange().getValues();
  var pwHash = _hashPassword(password);

  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().toLowerCase() === username.toLowerCase()) {
      var storedHash = (rows[i][1] || '').toString();
      if (storedHash === pwHash) {
        var token = _generateToken();
        var userData = {
          username: rows[i][0] || '',
          role: rows[i][2] || 'qlkv',
          display: rows[i][3] || rows[i][0],
          vung: rows[i][4] || '',
          tinh: rows[i][5] || '',
          _ts: Date.now()
        };
        _tokenStore().save(token, userData);
        return _json({ ok: true, token: token, user: { username: userData.username, role: userData.role, display: userData.display, vung: userData.vung, tinh: userData.tinh } });
      }
      return _err('Sai mật khẩu');
    }
  }
  return _err('Tài khoản không tồn tại');
}

// =============================================================================
// SURVEY — save to KhaoSat_QH tab
// =============================================================================
function _saveSurvey(user, data) {
  var maQH = (data.maQH || '').toString().trim();
  if (!maQH) return _err('Thiếu mã điểm QH');

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var headers = ['Mã QH', 'Ngày KS', 'Người KS', 'Username', 'Giá thuê KV (tr/th)', 'Địa chỉ nhà', 'Diện tích (m²)', 'MT (m)', 'Giá thuê (tr/th)', 'Cơ hội ghép căn', 'Ghi chú', 'Lat KS', 'Lng KS', 'Timestamp'];
  var sheet = _getOrCreateSheet(ss, SURVEY_TAB, headers);

  var now = new Date();
  var row = [
    maQH,
    data.ngayKS || Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    user.display,
    user.username,
    data.giaThueKV || '',
    data.diaChiNha || '',
    data.dienTich || '',
    data.matTien || '',
    data.giaThue || '',
    data.coHoiGhepCan || '',
    data.ghiChu || '',
    data.lat || '',
    data.lng || '',
    now.toISOString()
  ];
  sheet.appendRow(row);
  return _json({ ok: true, message: 'Đã lưu khảo sát cho ' + maQH });
}

// =============================================================================
// GET HISTORY — all surveys for a QH point
// =============================================================================
function _getHistory(data) {
  var maQH = (data.maQH || '').toString().trim();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SURVEY_TAB);
  if (!sheet) return _json({ ok: true, surveys: [] });

  var rows = sheet.getDataRange().getValues();
  var surveys = [];
  var headers = rows[0];
  for (var i = rows.length - 1; i >= 1; i--) {
    if ((rows[i][0] || '').toString().trim() === maQH) {
      surveys.push({
        maQH: rows[i][0] || '',
        ngayKS: rows[i][1] || '',
        nguoiKS: rows[i][2] || '',
        giaThueKV: rows[i][4] || '',
        diaChiNha: rows[i][5] || '',
        dienTich: rows[i][6] || '',
        matTien: rows[i][7] || '',
        giaThue: rows[i][8] || '',
        coHoiGhepCan: rows[i][9] || '',
        ghiChu: rows[i][10] || '',
        ts: rows[i][13] || ''
      });
    }
  }
  return _json({ ok: true, surveys: surveys });
}

// =============================================================================
// DASHBOARD — stats per user/GĐV
// =============================================================================
function _getDashboard(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SURVEY_TAB);
  var stats = { totalQH: 0, daKhaoSat: 0, chuaKhaoSat: 0, byGDV: {}, recentSurveys: [] };
  if (!sheet) return _json({ ok: true, stats: stats });

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var surveyed = {};    // maQH -> last survey
  for (var i = 1; i < rows.length; i++) {
    var qh = (rows[i][0] || '').toString().trim();
    if (qh && !surveyed[qh]) {
      surveyed[qh] = {
        maQH: qh,
        ngayKS: rows[i][1] || '',
        nguoiKS: rows[i][2] || '',
        diaChiNha: rows[i][5] || '',
        ts: rows[i][13] || ''
      };
    }
  }

  stats.daKhaoSat = Object.keys(surveyed).length;
  var recent = rows.slice(Math.max(1, rows.length - 20));
  for (var j = recent.length - 1; j >= 0; j--) {
    if (recent[j][0]) {
      stats.recentSurveys.push({ maQH: recent[j][0], ngayKS: recent[j][1], nguoiKS: recent[j][2], diaChiNha: recent[j][5] });
    }
  }

  return _json({ ok: true, stats: stats });
}

// =============================================================================
// LIST USERS (admin only)
// =============================================================================
function _listUsers(user) {
  if (user.role !== 'admin') return _err('Chỉ admin mới xem được danh sách users');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(USERS_TAB);
  if (!sheet) return _json({ ok: true, users: [] });
  var rows = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < rows.length; i++) {
    users.push({ username: rows[i][0] || '', role: rows[i][2] || '', display: rows[i][3] || '', vung: rows[i][4] || '', tinh: rows[i][5] || '' });
  }
  return _json({ ok: true, users: users });
}