/**
 * AUTO FACEBOOK POST — Google Apps Script (All-in-one)
 *
 * CÁCH CÀI ĐẶT:
 *   1. Extensions → Apps Script → dán file này vào → Save
 *   2. Chạy setupSpreadsheet() 1 lần để tạo tabs
 *   3. Project Settings → Script Properties → thêm SECRET_TOKEN = <chuỗi bí mật>
 *   4. Deploy → New deployment → Web App
 *      • Execute as: Me
 *      • Who has access: Anyone
 *   5. Copy URL → dán vào .env của Python (GOOGLE_WEBAPP_URL)
 *   6. Mở URL trên trình duyệt = giao diện quản lý
 */


// ═══════════════════════════════════════════════════════════════
// ENTRY POINTS
// ═══════════════════════════════════════════════════════════════

function doGet(e) {
  var action = (e.parameter && e.parameter.action) || "";

  // Không có action → giao diện quản lý
  if (!action) {
    var initialData = {};
    try { initialData = uiGetConfig(); } catch(err) { initialData = {groups:[],sources:[],dests:[],apify_keys:[],dedup_count:0}; }
    var html;
    try { html = _buildHtml(initialData); }
    catch(err2) { html = "<pre style='color:red;padding:20px'>Lỗi tạo giao diện: " + err2.toString() + "</pre>"; }
    return HtmlService
      .createHtmlOutput(html)
      .setTitle("Auto Facebook Post")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Có action → API cho Python
  if (!_auth(e)) return _err("Unauthorized");
  try { return _routeApi(action, e.parameter, null); }
  catch (err) { return _err(err.toString()); }
}

function doPost(e) {
  if (!_auth(e)) return _err("Unauthorized");
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (_) {}
  var action = body.action || (e.parameter && e.parameter.action) || "";
  try { return _routeApi(action, e.parameter, body); }
  catch (err) { return _err(err.toString()); }
}


// ═══════════════════════════════════════════════════════════════
// HÀM GỌI TỪ HTML (google.script.run — không bị CORS)
// ═══════════════════════════════════════════════════════════════

// Tải config (nhanh): groups, sources, dests, keys + đếm dedup bằng getLastRow
function uiGetConfig() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var dedupSh     = SS.getSheetByName("dedup");
  var dedup_count = dedupSh ? Math.max(0, dedupSh.getLastRow() - 1) : 0;
  return {
    groups:      _getAllRows(SS, "groups"),
    sources:     _getAllRows(SS, "source_pages"),
    dests:       _getAllRows(SS, "destination_pages"),
    apify_keys:  _getAllRows(SS, "apify_keys").map(function(r) {
      return Object.assign({}, r, {
        api_key: r.api_key ? r.api_key.substring(0, 16) + "..." : ""
      });
    }),
    dedup_count: dedup_count
  };
}

// Tải logs riêng: chỉ đọc 100 dòng mới nhất, không đọc toàn bộ bảng
function uiGetLogs() {
  var SS  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = SS.getSheetByName("logs");
  if (!sh || sh.getLastRow() <= 1) return [];
  var lastRow  = sh.getLastRow();
  var numCols  = sh.getLastColumn();
  if (numCols === 0) return [];
  var startRow = Math.max(2, lastRow - 99);
  var numRows  = lastRow - startRow + 1;
  var header   = sh.getRange(1, 1, 1, numCols).getValues()[0];
  var data     = sh.getRange(startRow, 1, numRows, numCols).getValues();
  return data.reverse().map(function(r) {
    var o = {}; header.forEach(function(k, i) { o[k] = r[i]; }); return o;
  });
}

// Giữ lại để tương thích (không dùng trong UI mới)
function uiGetAll() {
  var c = uiGetConfig();
  c.logs = uiGetLogs();
  return c;
}

// Chẩn đoán: trả về thông tin raw để debug
function uiDebug() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var info = { sheets: [], spreadsheet_name: SS.getName() };
  SS.getSheets().forEach(function(sh) {
    info.sheets.push({ name: sh.getName(), rows: sh.getLastRow(), cols: sh.getLastColumn() });
  });
  var grpSh = SS.getSheetByName("groups");
  if (grpSh) {
    var vals = grpSh.getDataRange().getValues();
    info.groups_raw = vals;
    info.groups_row_count = vals.length;
  } else {
    info.groups_raw = null;
    info.groups_row_count = 0;
  }
  // Gọi thẳng uiGetConfig để so sánh
  try {
    info.uiGetConfig_result = uiGetConfig();
  } catch(e) {
    info.uiGetConfig_error = e.toString();
  }
  return info;
}

function uiSaveRow(body) {
  var SS  = SpreadsheetApp.getActiveSpreadsheet();
  var tab = body.tab, rowData = body.row;
  var sh  = SS.getSheetByName(tab);
  if (!sh) return { ok: false, error: "Tab không tồn tại: " + tab };

  var vals   = sh.getDataRange().getValues();
  var header = vals[0];

  // Nếu có keyField + keyValue → sửa dòng hiện có
  if (body.keyField && body.keyValue !== undefined) {
    var kc = header.indexOf(body.keyField);
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][kc]) === String(body.keyValue)) {
        var updated = header.map(function(h) {
          return rowData[h] !== undefined ? rowData[h] : vals[i][header.indexOf(h)];
        });
        sh.getRange(i + 1, 1, 1, header.length).setValues([updated]);
        return { ok: true };
      }
    }
  }

  // Thêm mới
  var newRow = header.map(function(h) { return rowData[h] !== undefined ? rowData[h] : ""; });
  sh.appendRow(newRow);
  return { ok: true };
}

function uiDeleteRow(body) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sh = SS.getSheetByName(body.tab);
  if (!sh) return { ok: false, error: "Tab không tồn tại" };
  var vals   = sh.getDataRange().getValues();
  var header = vals[0];
  var kc     = header.indexOf(body.keyField);
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][kc]) === String(body.keyValue)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: "Không tìm thấy bản ghi" };
}

function uiClearLogs() {
  _clearOldLogs(SpreadsheetApp.getActiveSpreadsheet());
  return { ok: true };
}


// ═══════════════════════════════════════════════════════════════
// SETUP
// ═══════════════════════════════════════════════════════════════

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  _createTab(ss, "groups", [
    ["group_name", "is_active"],
    ["Nhóm mẫu", "TRUE"]
  ]);
  _createTab(ss, "source_pages", [
    ["group_name", "fb_page_url", "fb_page_name", "is_active", "last_scraped_at"],
    ["Nhóm mẫu", "https://www.facebook.com/thongtinchinhphu", "Thông tin Chính phủ", "TRUE", ""]
  ]);
  _createTab(ss, "destination_pages", [
    ["group_name", "fb_page_id", "fb_page_name", "fb_access_token", "is_active", "last_scheduled_at", "max_posts_per_run", "post_interval_hours"],
    ["Nhóm mẫu", "NHAP_PAGE_ID", "Tên trang của bạn", "NHAP_ACCESS_TOKEN", "TRUE", "", "4", "2"]
  ]);
  _createTab(ss, "apify_keys", [
    ["api_key", "email", "usage_count", "monthly_limit", "is_active", "last_used_at", "reset_at"],
    ["NHAP_APIFY_KEY", "your@email.com", "0", "450", "TRUE", "", _firstOfNextMonth()]
  ]);
  _createTab(ss, "dedup",    [["fb_post_id", "source_page_id", "destination_page_id", "posted_at"]]);
  _createTab(ss, "logs",     [["created_at", "fb_post_id", "destination_page_id", "result", "error_message"]]);
  _createTab(ss, "schedule", [["dest_page_id", "last_scheduled_at"]]);
  _createPendingReviewTab(ss);

  var def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  ss.getSheets().forEach(function(sh) {
    var last = Math.max(sh.getLastColumn(), 1);
    sh.getRange(1, 1, 1, last)
      .setBackground("#1a73e8").setFontColor("#fff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, last);
  });

  SpreadsheetApp.getUi().alert(
    "✅ Setup hoàn tất! 8 tabs đã được tạo.\n\n" +
    "Tiếp theo:\n" +
    "1. Deploy → New deployment → Web App\n" +
    "2. Execute as: Me | Who has access: Anyone\n" +
    "3. Copy URL → dán vào .env"
  );
}

function _createTab(ss, name, rows) {
  if (ss.getSheetByName(name)) return;
  var sh = ss.insertSheet(name);
  if (rows && rows.length) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

// Chạy hàm này 1 lần sau khi update Code.gs để thêm cột mới vào sheet cũ
function migrateSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // destination_pages: chỉ thêm header + 2 cột ở cuối hàng 1
  var dp = ss.getSheetByName("destination_pages");
  if (dp) {
    var dpHeader = dp.getRange(1, 1, 1, dp.getLastColumn()).getValues()[0];
    if (dpHeader.indexOf("max_posts_per_run") === -1)
      dp.getRange(1, dp.getLastColumn() + 1).setValue("max_posts_per_run");
    if (dpHeader.indexOf("post_interval_hours") === -1)
      dp.getRange(1, dp.getLastColumn() + 1).setValue("post_interval_hours");
  }

  // logs: chỉ thêm header ô cuối hàng 1, không đụng data
  var lg = ss.getSheetByName("logs");
  if (lg && lg.getLastColumn() > 0) {
    var lgHeader = lg.getRange(1, 1, 1, lg.getLastColumn()).getValues()[0];
    if (lgHeader.indexOf("source_page_url") === -1)
      lg.getRange(1, lg.getLastColumn() + 1).setValue("source_page_url");
  }

  Browser.msgBox("✅ Migration xong! Hãy vào sheet destination_pages điền max_posts_per_run và post_interval_hours cho từng trang đích.");
}


// ═══════════════════════════════════════════════════════════════
// API ROUTER (cho Python script)
// ═══════════════════════════════════════════════════════════════

function _routeApi(action, params, body) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  switch (action) {
    case "get_active_sources":       return _json(_getActiveSources(SS));
    case "get_destinations":         return _json(_getDestinations(SS, params.group_id));
    case "get_apify_key":            return _json(_getApifyKey(SS));
    case "is_dedup":                 return _json({ exists: _isDedup(SS, params.fb_post_id) });
    case "save_dedup":               _saveDedup(SS, body.fb_post_id, body.source_page_id, body.destination_page_id); return _json({ ok: true });
    case "get_last_scheduled":       return _json({ last_scheduled_at: _getLastScheduled(SS, params.dest_page_id) });
    case "update_last_scheduled":    _updateLastScheduled(SS, body.dest_page_id, body.last_scheduled_at); return _json({ ok: true });
    case "increment_apify_usage":    _incrementApifyUsage(SS, body.api_key, parseInt(body.count) || 1); return _json({ ok: true });
    case "update_source_scraped_at": _updateSourceScrapedAt(SS, body.page_url, body.scraped_at); return _json({ ok: true });
    case "save_log":                 _saveLog(SS, body.fb_post_id, body.destination_page_id, body.result, body.error_message || "", body.source_page_url || "", body.post_url || ""); return _json({ ok: true });
    case "save_pending_review":      _savePendingReview(SS, body.rows || []); return _json({ ok: true });
    case "get_approved_reviews":     return _json({ rows: _getApprovedReviews(SS, parseInt(params.limit || (body && body.limit) || 20)) });
    case "update_review_status":     _updateReviewStatus(SS, body.id, body.status, body.fields || {}); return _json({ ok: true });
    case "clear_dedup":              _clearSheet(SS, "dedup"); return _json({ ok: true });
    case "clear_schedule":           _clearSheet(SS, "schedule"); return _json({ ok: true });
    default:                         return _err("Unknown action: " + action);
  }
}


// ═══════════════════════════════════════════════════════════════
// PYTHON API IMPLEMENTATIONS
// ═══════════════════════════════════════════════════════════════

function _getActiveSources(SS) {
  var gmap = {};
  _getAllRows(SS, "groups").forEach(function(r) { gmap[r.group_name] = r.is_active; });
  return _getAllRows(SS, "source_pages").filter(function(r) {
    return _isTrue(r.is_active) && _isTrue(gmap[r.group_name]);
  }).map(function(r) {
    return { id: r.fb_page_url, group_id: r.group_name, fb_page_url: r.fb_page_url,
             fb_page_name: r.fb_page_name || "", page_groups: { name: r.group_name } };
  });
}

function _getDestinations(SS, groupId) {
  return _getAllRows(SS, "destination_pages").filter(function(r) {
    return r.group_name === groupId && _isTrue(r.is_active);
  }).map(function(r) {
    return { id: String(r.fb_page_id), group_id: r.group_name, fb_page_id: String(r.fb_page_id),
             fb_page_name: r.fb_page_name || "", fb_access_token: r.fb_access_token };
  });
}

function _getApifyKey(SS) {
  _resetMonthlyUsageIfNeeded(SS);
  var rows = _getAllRows(SS, "apify_keys").filter(function(r) {
    return _isTrue(r.is_active) && parseInt(r.usage_count || 0) < parseInt(r.monthly_limit || 450);
  });
  if (!rows.length) return { api_key: null };
  rows.sort(function(a, b) { return parseInt(a.usage_count || 0) - parseInt(b.usage_count || 0); });
  return { api_key: rows[0].api_key };
}

function _isDedup(SS, fbPostId) {
  return _getAllRows(SS, "dedup").some(function(r) { return String(r.fb_post_id) === String(fbPostId); });
}

function _saveDedup(SS, fbPostId, src, dest) {
  SS.getSheetByName("dedup").appendRow([String(fbPostId), String(src), String(dest), new Date().toISOString()]);
}

function _getLastScheduled(SS, destPageId) {
  var rows = _getAllRows(SS, "schedule");
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].dest_page_id) === String(destPageId)) return rows[i].last_scheduled_at || null;
  }
  return null;
}

function _updateLastScheduled(SS, destPageId, scheduledAt) {
  _upsertByKey(SS, "schedule", "dest_page_id", String(destPageId), "last_scheduled_at", scheduledAt);
}

function _incrementApifyUsage(SS, apiKey, count) {
  var sh = SS.getSheetByName("apify_keys"), vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;
  var h = vals[0], kc = h.indexOf("api_key"), uc = h.indexOf("usage_count"), ac = h.indexOf("last_used_at");
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][kc] === apiKey) {
      if (uc !== -1) sh.getRange(i + 1, uc + 1).setValue(parseInt(vals[i][uc] || 0) + count);
      if (ac !== -1) sh.getRange(i + 1, ac + 1).setValue(new Date().toISOString());
      return;
    }
  }
}

function _updateSourceScrapedAt(SS, pageUrl, scrapedAt) {
  _upsertByKey(SS, "source_pages", "fb_page_url", pageUrl, "last_scraped_at", scrapedAt);
}

function _saveLog(SS, fbPostId, destPageId, result, errMsg, sourcePageUrl, postUrl) {
  var sh = SS.getSheetByName("logs");
  if (!sh) return;
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) {
    sh.appendRow(["created_at", "fb_post_id", "destination_page_id", "result", "error_message", "source_page_url", "post_url"]);
    sh.appendRow([new Date().toISOString(), String(fbPostId), String(destPageId), result, errMsg || "", sourcePageUrl || "", postUrl || ""]);
    return;
  }
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf("source_page_url") === -1) sh.getRange(1, lastCol + 1).setValue("source_page_url");
  if (header.indexOf("post_url") === -1) sh.getRange(1, sh.getLastColumn() + 1).setValue("post_url");
  sh.appendRow([new Date().toISOString(), String(fbPostId), String(destPageId), result, errMsg || "", sourcePageUrl || "", postUrl || ""]);
}


function _pendingReviewHeaders() {
  return [
    "id", "group_id", "destination_page_id", "source_page", "source_url",
    "fb_post_id", "original_text", "ai_summary", "rewritten_caption",
    "suggested_hashtags", "suggested_cta", "media_type", "media_urls",
    "likes", "comments", "shares", "risk_level", "filter_result",
    "filter_reasons", "status", "reviewed_by", "reviewed_at", "review_note",
    "scheduled_time", "facebook_post_id", "created_at", "updated_at"
  ];
}

function _createPendingReviewTab(SS) {
  _createTab(SS, "pending_review", [_pendingReviewHeaders()]);
}

function _ensurePendingReviewTab(SS) {
  var sh = SS.getSheetByName("pending_review");
  if (!sh) {
    _createPendingReviewTab(SS);
    sh = SS.getSheetByName("pending_review");
  }
  var headers = _pendingReviewHeaders();
  var existing = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  headers.forEach(function(h) {
    if (existing.indexOf(h) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
    }
  });
  return sh;
}

function _savePendingReview(SS, rows) {
  var sh = _ensurePendingReviewTab(SS);
  if (!Array.isArray(rows) || rows.length === 0) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var existingIds = {};
  var vals = sh.getDataRange().getValues();
  var idCol = headers.indexOf("id");
  for (var i = 1; i < vals.length; i++) existingIds[String(vals[i][idCol])] = true;
  var out = [];
  rows.forEach(function(row) {
    if (existingIds[String(row.id)]) return;
    out.push(headers.map(function(h) { return row[h] !== undefined ? row[h] : ""; }));
  });
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, headers.length).setValues(out);
}

function _getApprovedReviews(SS, limit) {
  var sh = _ensurePendingReviewTab(SS);
  var vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  var headers = vals[0], statusCol = headers.indexOf("status"), out = [];
  var destRows = _getAllRows(SS, "destination_pages");
  var destMap = {};
  destRows.forEach(function(r) { destMap[String(r.fb_page_id)] = r; });
  for (var i = 1; i < vals.length && out.length < (limit || 20); i++) {
    if (String(vals[i][statusCol]).toLowerCase() === "approved") {
      var row = {};
      headers.forEach(function(h, j) { row[h] = vals[i][j]; });
      row._sheet_row = i + 1;
      var dest = destMap[String(row.destination_page_id)] || {};
      row.fb_access_token = dest.fb_access_token || "";
      row.fb_page_name = dest.fb_page_name || "";
      out.push(row);
    }
  }
  return out;
}

function _updateReviewStatus(SS, id, status, fields) {
  var sh = _ensurePendingReviewTab(SS);
  var vals = sh.getDataRange().getValues();
  var headers = vals[0], idCol = headers.indexOf("id"), statusCol = headers.indexOf("status");
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === String(id)) {
      sh.getRange(i + 1, statusCol + 1).setValue(status);
      Object.keys(fields || {}).forEach(function(k) {
        var c = headers.indexOf(k);
        if (c !== -1) sh.getRange(i + 1, c + 1).setValue(fields[k]);
      });
      return;
    }
  }
  throw new Error("Không tìm thấy pending_review id: " + id);
}

function _resetMonthlyUsageIfNeeded(SS) {
  var sh = SS.getSheetByName("apify_keys"), vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;
  var h = vals[0], uc = h.indexOf("usage_count"), rc = h.indexOf("reset_at");
  if (uc === -1 || rc === -1) return;
  var today = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd"), nm = _firstOfNextMonth();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][rc] && String(vals[i][rc]) <= today) {
      sh.getRange(i + 1, uc + 1).setValue(0);
      sh.getRange(i + 1, rc + 1).setValue(nm);
    }
  }
}

function _clearSheet(SS, tabName) {
  var sh = SS.getSheetByName(tabName);
  if (!sh || sh.getLastRow() <= 1) return;
  sh.deleteRows(2, sh.getLastRow() - 1);
}

function _clearOldLogs(SS) {
  var sh = SS.getSheetByName("logs"), vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;
  var cut = new Date(); cut.setDate(cut.getDate() - 30);
  for (var i = vals.length - 1; i >= 1; i--) {
    if (new Date(vals[i][0]) < cut) sh.deleteRow(i + 1);
  }
}


// ═══════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════

function _getAllRows(SS, tabName) {
  var sh = SS.getSheetByName(tabName);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  var h = vals[0];
  return vals.slice(1).map(function(r) {
    var o = {}; h.forEach(function(k, i) { o[k] = r[i]; }); return o;
  });
}

function _upsertByKey(SS, tab, keyField, keyValue, updateField, updateValue) {
  var sh = SS.getSheetByName(tab), vals = sh.getDataRange().getValues();
  if (vals.length <= 1) { sh.appendRow([keyValue, updateValue]); return; }
  var h = vals[0], kc = h.indexOf(keyField), vc = h.indexOf(updateField);
  if (kc === -1 || vc === -1) return;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][kc]) === String(keyValue)) {
      sh.getRange(i + 1, vc + 1).setValue(updateValue); return;
    }
  }
  var nr = new Array(h.length).fill(""); nr[kc] = keyValue; nr[vc] = updateValue;
  sh.appendRow(nr);
}

function _isTrue(v) { return v === true || String(v).toUpperCase() === "TRUE"; }
function _firstOfNextMonth() {
  var d = new Date(); d.setMonth(d.getMonth() + 1); d.setDate(1);
  return Utilities.formatDate(d, "UTC", "yyyy-MM-dd");
}
function _auth(e) {
  var s = PropertiesService.getScriptProperties().getProperty("SECRET_TOKEN") || "";
  if (!s) return true;
  return (e.parameter && e.parameter.token) === s;
}
function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function _err(msg) {
  return ContentService.createTextOutput(JSON.stringify({ error: msg })).setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════
// HTML UI
// ═══════════════════════════════════════════════════════════════

function _buildHtml(initialData) {
  var css = '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f0f2f5;display:flex}'
    + '.sidebar{width:210px;min-height:100vh;background:#1877f2;color:#fff;position:fixed;top:0;left:0;height:100%;display:flex;flex-direction:column}'
    + '.logo{padding:16px;border-bottom:1px solid rgba(255,255,255,.2)}'
    + '.logo h2{font-size:14px;font-weight:700}'
    + '.logo p{font-size:11px;opacity:.65;margin-top:2px}'
    + '.nav a{display:flex;align-items:center;gap:8px;padding:11px 16px;font-size:13px;color:rgba(255,255,255,.85);cursor:pointer;border-left:3px solid transparent;transition:.15s}'
    + '.nav a:hover{background:rgba(255,255,255,.12)}'
    + '.nav a.on{background:rgba(255,255,255,.2);border-left-color:#fff;color:#fff;font-weight:600}'
    + '.main{margin-left:210px;flex:1;padding:20px}'
    + '.topbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}'
    + '.topbar h1{font-size:19px;font-weight:700}'
    + '.card{background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:16px;overflow:hidden}'
    + '.card-h{padding:12px 16px;border-bottom:1px solid #f0f2f5;display:flex;align-items:center;justify-content:space-between}'
    + '.card-h h2{font-size:14px;font-weight:600;color:#333}'
    + 'table{width:100%;border-collapse:collapse;font-size:13px}'
    + 'th{background:#f7f8fa;padding:9px 12px;text-align:left;font-weight:600;color:#555;border-bottom:1px solid #eee;white-space:nowrap}'
    + 'td{padding:9px 12px;border-bottom:1px solid #f7f8fa;vertical-align:middle}'
    + 'tr:last-child td{border:none}'
    + 'tr:hover td{background:#fafbff}'
    + '.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}'
    + '.stat{background:#fff;border-radius:8px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08)}'
    + '.stat .n{font-size:28px;font-weight:700;color:#1877f2}'
    + '.stat .l{font-size:12px;color:#65676b;margin-top:4px}'
    + '.badge{display:inline-flex;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}'
    + '.g{background:#e8f5e9;color:#2e7d32}'
    + '.r{background:#ffebee;color:#c62828}'
    + '.b{background:#e3f2fd;color:#1565c0}'
    + '.y{background:#fff8e1;color:#7d6000}'
    + '.btn{display:inline-flex;align-items:center;gap:4px;padding:6px 12px;border-radius:6px;font-size:13px;font-weight:500;cursor:pointer;border:none;transition:.15s}'
    + '.p{background:#1877f2;color:#fff}.p:hover{background:#166fe5}'
    + '.d{background:#e53935;color:#fff}.d:hover{background:#c62828}'
    + '.o{background:#fff;color:#1877f2;border:1.5px solid #1877f2}.o:hover{background:#e7f0ff}'
    + '.sm{padding:3px 8px;font-size:12px}'
    + '.prog{height:5px;background:#eee;border-radius:3px;margin-top:3px}'
    + '.pf{height:100%;background:#1877f2;border-radius:3px}'
    + '.tab{display:none}.tab.on{display:block}'
    + '.modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:999;align-items:center;justify-content:center}'
    + '.modal.on{display:flex}'
    + '.mbox{background:#fff;border-radius:8px;width:500px;max-height:90vh;overflow-y:auto;box-shadow:0 8px 40px rgba(0,0,0,.2)}'
    + '.mh{padding:14px 18px;border-bottom:1px solid #eee;display:flex;align-items:center;justify-content:space-between}'
    + '.mh h3{font-size:15px;font-weight:700}'
    + '.mb{padding:18px}'
    + '.mf{padding:12px 18px;border-top:1px solid #eee;display:flex;gap:8px;justify-content:flex-end}'
    + '.fg{margin-bottom:12px}'
    + 'label{display:block;font-size:12px;font-weight:600;margin-bottom:4px;color:#333}'
    + '.fc{width:100%;padding:8px 10px;border:1.5px solid #ddd;border-radius:6px;font-size:13px;outline:none;background:#fff}'
    + '.fc:focus{border-color:#1877f2;box-shadow:0 0 0 3px rgba(24,119,242,.1)}'
    + '.hint{font-size:11px;color:#888;margin-top:3px}'
    + '.info{background:#e8f4fd;border:1px solid #b3d8f5;border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:13px}'
    + '.warn{background:#fff8e1;border:1px solid #ffe082;border-radius:6px;padding:8px 12px;font-size:12px;color:#7d6000;margin-top:8px}'
    + '.toast{position:fixed;bottom:20px;right:20px;padding:10px 16px;border-radius:6px;color:#fff;font-size:13px;z-index:9999;opacity:0;transform:translateY(10px);transition:.25s;pointer-events:none}'
    + '.toast.on{opacity:1;transform:none}'
    + '.ts{background:#43a047}.te{background:#e53935}'
    + '.empty{padding:28px;text-align:center;color:#888;font-size:13px}'
    + 'code{background:#f0f2f5;padding:1px 5px;border-radius:3px;font-size:12px}'
    + 'a.lnk{color:#1877f2;font-size:12px;text-decoration:none}'
    + 'a.lnk:hover{text-decoration:underline}'
    + 'ol{padding-left:16px;line-height:2;font-size:13px}'
    + '</style>';

  var html = '<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Auto Facebook Post</title>' + css + '</head><body>'

    // Sidebar
    + '<div class="sidebar">'
    + '<div class="logo"><h2>📘 Auto FB Post</h2><p>Hệ thống đăng bài tự động</p></div>'
    + '<nav class="nav">'
    + '<a class="on" data-tab="dash">📊 Tổng quan</a>'
    + '<a data-tab="groups">📁 Nhóm trang</a>'
    + '<a data-tab="sources">🔍 Trang nguồn</a>'
    + '<a data-tab="dests">📤 Trang đích</a>'
    + '<a data-tab="fb">🔐 Facebook Token</a>'
    + '<a data-tab="keys">🔑 Apify Keys</a>'
    + '<a data-tab="logs">📋 Nhật ký</a>'
    + '</nav></div>'

    // Main
    + '<div class="main">'

    // DASHBOARD
    + '<div id="tab-dash" class="tab on">'
    + '<div class="topbar"><h1>📊 Tổng quan</h1>'
    + '<button class="btn o" id="btn-reload">🔄 Làm mới</button></div>'
    + '<div class="stats">'
    + '<div class="stat"><div class="n" id="s0">—</div><div class="l">Nhóm trang</div></div>'
    + '<div class="stat"><div class="n" id="s1">—</div><div class="l">Trang nguồn</div></div>'
    + '<div class="stat"><div class="n" id="s2">—</div><div class="l">Trang đích</div></div>'
    + '<div class="stat"><div class="n" id="s3">—</div><div class="l">Bài đã đăng</div></div>'
    + '</div>'
    + '<div class="card"><div class="card-h"><h2>📋 Hoạt động gần đây</h2>'
    + '<button class="btn d sm" id="btn-clearlogs">🗑 Xóa log &gt;30 ngày</button></div>'
    + '<div id="logs-dash"><div class="empty">Đang tải...</div></div></div>'
    + '</div>'

    // GROUPS
    + '<div id="tab-groups" class="tab">'
    + '<div class="topbar"><h1>📁 Nhóm trang</h1>'
    + '<button class="btn p" data-add="groups">＋ Thêm nhóm</button>'
    + '<button class="btn" id="btn-debug" style="margin-left:8px;background:#666">🔍 Debug</button>'
    + '<button class="btn" id="btn-testrender" style="margin-left:8px;background:#2e7d32">⚡ Test Render</button></div>'
    + '<div id="debug-out" style="display:none;background:#222;color:#0f0;padding:12px;border-radius:6px;margin-bottom:12px;font-size:12px;white-space:pre-wrap;word-break:break-all"></div>'
    + '<div class="card"><table><thead><tr><th>Tên nhóm</th><th>Trạng thái</th><th>Thao tác</th></tr></thead>'
    + '<tbody id="tb-groups"><tr><td colspan="3" class="empty">Đang tải...</td></tr></tbody></table></div>'
    + '</div>'

    // SOURCES
    + '<div id="tab-sources" class="tab">'
    + '<div class="topbar"><h1>🔍 Trang nguồn (fanpage đối thủ)</h1>'
    + '<button class="btn p" data-add="source_pages">＋ Thêm trang nguồn</button></div>'
    + '<div class="card"><table><thead><tr><th>Nhóm</th><th>Tên trang</th><th>URL</th><th>Trạng thái</th><th>Scrape lần cuối</th><th></th></tr></thead>'
    + '<tbody id="tb-sources"><tr><td colspan="6" class="empty">Đang tải...</td></tr></tbody></table></div>'
    + '</div>'

    // DESTS
    + '<div id="tab-dests" class="tab">'
    + '<div class="topbar"><h1>📤 Trang đích (page của bạn)</h1>'
    + '<button class="btn p" data-add="destination_pages">＋ Thêm trang đích</button></div>'
    + '<div class="card"><table><thead><tr><th>Nhóm</th><th>Tên trang</th><th>Page ID</th><th>Trạng thái</th><th>Lịch đăng</th><th>Hẹn giờ cuối</th><th></th></tr></thead>'
    + '<tbody id="tb-dests"><tr><td colspan="7" class="empty">Đang tải...</td></tr></tbody></table></div>'
    + '</div>'

    // FACEBOOK TOKEN GUIDE
    + '<div id="tab-fb" class="tab">'
    + '<div class="topbar"><h1>🔐 Hướng dẫn lấy Facebook Access Token</h1></div>'
    + '<div class="card"><div class="card-h"><h2>Bước 1 — Tạo Facebook App</h2></div>'
    + '<div style="padding:16px"><ol>'
    + '<li>Vào <a class="lnk" href="https://developers.facebook.com/apps" target="_blank">developers.facebook.com/apps</a> → <b>Create App</b></li>'
    + '<li>Chọn loại: <b>Business</b> → điền tên app → Create</li>'
    + '<li>Vào app vừa tạo → <b>Add Product</b> → chọn <b>Facebook Login</b> → Set up</li>'
    + '</ol></div></div>'
    + '<div class="card"><div class="card-h"><h2>Bước 2 — Lấy Page Access Token</h2></div>'
    + '<div style="padding:16px"><ol>'
    + '<li>Mở <a class="lnk" href="https://developers.facebook.com/tools/explorer/" target="_blank">Graph API Explorer</a></li>'
    + '<li>Góc phải: chọn App → <b>Generate Access Token</b></li>'
    + '<li>Chọn <b>Get Page Access Token</b> → chọn Page của bạn</li>'
    + '<li>Tick quyền: <code>pages_manage_posts</code> và <code>pages_read_engagement</code></li>'
    + '<li>Nhấn <b>Generate Access Token</b> → copy token</li>'
    + '</ol>'
    + '<div class="warn">⚠️ Token vừa lấy chỉ dùng được ~2 giờ. Làm Bước 3 để lấy token 60 ngày.</div>'
    + '</div></div>'
    + '<div class="card"><div class="card-h"><h2>Bước 3 — Đổi thành Long-lived Token (60 ngày)</h2></div>'
    + '<div style="padding:16px">'
    + '<p style="font-size:13px;margin-bottom:10px">Mở trình duyệt, gọi URL sau (thay các giá trị IN HOA):</p>'
    + '<code style="display:block;background:#f0f2f5;padding:10px;border-radius:6px;line-height:1.8;word-break:break-all">'
    + 'https://graph.facebook.com/oauth/access_token?grant_type=fb_exchange_token&amp;client_id=APP_ID&amp;client_secret=APP_SECRET&amp;fb_exchange_token=TOKEN_NGAN_HAN'
    + '</code>'
    + '<p style="font-size:12px;color:#888;margin-top:8px">APP_ID và APP_SECRET: vào App Dashboard → Settings → Basic</p>'
    + '</div></div>'
    + '<div class="card"><div class="card-h"><h2>Bước 4 — Lấy Page ID</h2></div>'
    + '<div style="padding:16px">'
    + '<p style="font-size:13px;margin-bottom:8px">Trên Graph API Explorer, gõ vào ô tìm kiếm:</p>'
    + '<code>me/accounts</code>'
    + '<p style="font-size:12px;color:#888;margin-top:8px">Response trả về danh sách Pages với <b>id</b> (Page ID số) và <b>access_token</b>.</p>'
    + '</div></div>'
    + '<div class="card"><div class="card-h"><h2>Bước 5 — Nhập vào hệ thống</h2></div>'
    + '<div style="padding:16px">'
    + '<p style="font-size:13px;margin-bottom:10px">Vào tab <b>📤 Trang đích</b> → <b>＋ Thêm trang đích</b> để nhập Page ID và Access Token.</p>'
    + '<button class="btn p" data-tab="dests">→ Đến tab Trang đích</button>'
    + '</div></div>'
    + '</div>'

    // APIFY KEYS
    + '<div id="tab-keys" class="tab">'
    + '<div class="topbar"><h1>🔑 Apify Keys</h1>'
    + '<button class="btn p" data-add="apify_keys">＋ Thêm key</button></div>'
    + '<div class="card"><table><thead><tr><th>Email</th><th>API Key</th><th>Sử dụng</th><th>Trạng thái</th><th></th></tr></thead>'
    + '<tbody id="tb-keys"><tr><td colspan="5" class="empty">Đang tải...</td></tr></tbody></table></div>'
    + '</div>'

    // LOGS
    + '<div id="tab-logs" class="tab">'
    + '<div class="topbar"><h1>📋 Nhật ký đăng bài</h1>'
    + '<button class="btn d sm" id="btn-clearlogs2">🗑 Xóa log &gt;30 ngày</button></div>'
    + '<div class="card" id="logs-full"><div class="empty">Đang tải...</div></div>'
    + '</div>'

    + '</div>' // end main

    // MODAL
    + '<div class="modal" id="modal">'
    + '<div class="mbox">'
    + '<div class="mh"><h3 id="m-title"></h3>'
    + '<button id="btn-modal-close" style="background:none;border:none;font-size:22px;cursor:pointer;color:#666;line-height:1">✕</button>'
    + '</div>'
    + '<div class="mb" id="m-body"></div>'
    + '<div class="mf">'
    + '<button class="btn o" id="btn-modal-cancel">Hủy</button>'
    + '<button class="btn p" id="btn-modal-save">💾 Lưu</button>'
    + '</div></div></div>'

    + '<div class="toast" id="toast"></div>';

  var js = '<script>'
    // ── State
    + 'var D = Object.assign({groups:[],sources:[],dests:[],apify_keys:[],logs:[],dedup_count:0}, ' + JSON.stringify(initialData) + ');'
    + 'var MTAB = "", MKF = null, MKV = null;'

    // ── Tabs
    + 'function showTab(t) {'
    + '  document.querySelectorAll(".tab").forEach(function(el){el.classList.remove("on");});'
    + '  document.querySelectorAll(".nav a").forEach(function(el){el.classList.remove("on");});'
    + '  var tab = document.getElementById("tab-" + t);'
    + '  if (tab) tab.classList.add("on");'
    + '  var nav = document.querySelector(".nav a[data-tab=" + t + "]");'
    + '  if (nav) nav.classList.add("on");'
    + '}'

    // ── Load data
    + 'function reload() {'
    + '  document.getElementById("logs-dash").innerHTML = "<div class=empty>⏳ Đang tải logs...</div>";'
    + '  google.script.run'
    + '    .withSuccessHandler(function(cfg) {'
    + '      D = Object.assign(D || {}, cfg);'
    + '      document.getElementById("s0").textContent = (D.groups||[]).length;'
    + '      document.getElementById("s1").textContent = (D.sources||[]).length;'
    + '      document.getElementById("s2").textContent = (D.dests||[]).length;'
    + '      document.getElementById("s3").textContent = D.dedup_count || 0;'
    + '      renderGroups(); renderSources(); renderDests(); renderKeys();'
    + '      google.script.run'
    + '        .withSuccessHandler(function(logs) {'
    + '          D.logs = logs;'
    + '          renderLogs();'
    + '          toast("Đã cập nhật", "ts");'
    + '        })'
    + '        .withFailureHandler(function(e) { toast("Lỗi logs: " + e, "te"); })'
    + '        .uiGetLogs();'
    + '    })'
    + '    .withFailureHandler(function(e) { toast("Lỗi: " + e, "te"); })'
    + '    .uiGetConfig();'
    + '}'

    // ── Render
    + 'function render() {'
    + '  document.getElementById("s0").textContent = D.groups.length;'
    + '  document.getElementById("s1").textContent = (D.sources || []).length;'
    + '  document.getElementById("s2").textContent = (D.dests || []).length;'
    + '  document.getElementById("s3").textContent = D.dedup_count || 0;'
    + '  renderGroups(); renderSources(); renderDests(); renderKeys(); renderLogs();'
    + '}'

    + 'function badge(v) {'
    + '  var ok = v === true || String(v).toUpperCase() === "TRUE";'
    + '  return ok ? "<span class=\\"badge g\\">✓ Hoạt động</span>" : "<span class=\\"badge r\\">✗ Tắt</span>";'
    + '}'

    + 'function fmtDate(s) {'
    + '  if (!s) return "-";'
    + '  try { return new Date(s).toLocaleString("vi-VN"); } catch(e) { return s; }'
    + '}'

    + 'function mkBtn(type, tab, kf, kv) {'
    + '  var enc = encodeURIComponent(kv);'
    + '  if (type === "edit") return "<button class=\\"btn o sm\\" data-edit=\\"1\\" data-tab=\\"" + tab + "\\" data-kf=\\"" + kf + "\\" data-kv=\\"" + enc + "\\" style=\\"margin-right:4px\\">✏️</button>";'
    + '  return "<button class=\\"btn d sm\\" data-del=\\"1\\" data-tab=\\"" + tab + "\\" data-kf=\\"" + kf + "\\" data-kv=\\"" + enc + "\\">🗑</button>";'
    + '}'

    + 'function renderGroups() {'
    + '  var rows = D.groups || [];'
    + '  var html = rows.length ? rows.map(function(r) {'
    + '    return "<tr><td><b>" + r.group_name + "</b></td><td>" + badge(r.is_active) + "</td><td>"'
    + '      + mkBtn("edit","groups","group_name",r.group_name)'
    + '      + mkBtn("del","groups","group_name",r.group_name) + "</td></tr>";'
    + '  }).join("") : "<tr><td colspan=3 class=empty>Chưa có nhóm nào</td></tr>";'
    + '  document.getElementById("tb-groups").innerHTML = html;'
    + '}'

    + 'function renderSources() {'
    + '  var rows = D.sources || [];'
    + '  var html = rows.length ? rows.map(function(r) {'
    + '    var short = r.fb_page_url.replace("https://www.facebook.com/", "fb/");'
    + '    return "<tr>"'
    + '      + "<td><span class=\\"badge b\\">" + r.group_name + "</span></td>"'
    + '      + "<td><b>" + (r.fb_page_name || "") + "</b></td>"'
    + '      + "<td><a class=\\"lnk\\" href=\\"" + r.fb_page_url + "\\" target=\\"_blank\\">" + short + "</a></td>"'
    + '      + "<td>" + badge(r.is_active) + "</td>"'
    + '      + "<td style=\\"font-size:11px;color:#888\\">" + fmtDate(r.last_scraped_at) + "</td>"'
    + '      + "<td>" + mkBtn("edit","source_pages","fb_page_url",r.fb_page_url) + mkBtn("del","source_pages","fb_page_url",r.fb_page_url) + "</td>"'
    + '      + "</tr>";'
    + '  }).join("") : "<tr><td colspan=6 class=empty>Chưa có trang nguồn nào</td></tr>";'
    + '  document.getElementById("tb-sources").innerHTML = html;'
    + '}'

    + 'function renderDests() {'
    + '  var rows = D.dests || [];'
    + '  var html = rows.length ? rows.map(function(r) {'
    + '    return "<tr>"'
    + '      + "<td><span class=\\"badge b\\">" + r.group_name + "</span></td>"'
    + '      + "<td><b>" + (r.fb_page_name || "") + "</b></td>"'
    + '      + "<td style=\\"font-family:monospace;font-size:12px\\">" + r.fb_page_id + "</td>"'
    + '      + "<td>" + badge(r.is_active) + "</td>"'
    + '      + "<td style=\\"font-size:11px;color:#888\\"><b>" + (r.max_posts_per_run || 4) + "</b> bài / <b>" + (r.post_interval_hours || 2) + "</b>h</td>"'
    + '      + "<td style=\\"font-size:11px;color:#888\\">" + fmtDate(r.last_scheduled_at) + "</td>"'
    + '      + "<td>" + mkBtn("edit","destination_pages","fb_page_id",r.fb_page_id) + mkBtn("del","destination_pages","fb_page_id",r.fb_page_id) + "</td>"'
    + '      + "</tr>";'
    + '  }).join("") : "<tr><td colspan=7 class=empty>Chưa có trang đích nào</td></tr>";'
    + '  document.getElementById("tb-dests").innerHTML = html;'
    + '}'

    + 'function renderKeys() {'
    + '  var rows = D.apify_keys || [];'
    + '  var html = rows.length ? rows.map(function(r) {'
    + '    var u = parseInt(r.usage_count || 0), l = parseInt(r.monthly_limit || 450);'
    + '    var p = Math.min(100, Math.round(u / l * 100));'
    + '    var bar = "<div style=\\"font-size:12px\\">" + u + "/" + l + " (" + p + "%)</div>"'
    + '      + "<div class=prog><div class=pf style=\\"width:" + p + "%\\"></div></div>";'
    + '    return "<tr>"'
    + '      + "<td>" + (r.email || "-") + "</td>"'
    + '      + "<td style=\\"font-family:monospace;font-size:11px\\">" + (r.api_key || "") + "</td>"'
    + '      + "<td style=\\"min-width:130px\\">" + bar + "</td>"'
    + '      + "<td>" + badge(r.is_active) + "</td>"'
    + '      + "<td>" + mkBtn("del","apify_keys","email",r.email) + "</td>"'
    + '      + "</tr>";'
    + '  }).join("") : "<tr><td colspan=5 class=empty>Chưa có key nào</td></tr>";'
    + '  document.getElementById("tb-keys").innerHTML = html;'
    + '}'

    + 'function renderLogs() {'
    + '  var rows = (D.logs || []).slice(0, 100);'
    + '  var t = "<table><thead><tr><th>Thời gian</th><th>Trang nguồn</th><th>Bài gốc</th><th>Trang đích</th><th>Kết quả</th><th>Lỗi</th></tr></thead><tbody>";'
    + '  if (rows.length) {'
    + '    t += rows.map(function(r) {'
    + '      var ok = r.result === "scheduled" || r.result === "success";'
    + '      var res = ok ? "<span class=\\"badge g\\">" + r.result + "</span>" : "<span class=\\"badge r\\">" + r.result + "</span>";'
    + '      var srcUrl = r.source_page_url || ""; var srcParts = srcUrl.split("facebook.com/"); var src = srcParts.length > 1 ? srcParts[srcParts.length-1] : srcUrl;'
    + '      var pUrl = r.post_url || ""; var postLink = pUrl ? "<a href=\\"" + pUrl + "\\" target=\\"_blank\\" style=\\"color:#1a73e8;font-size:11px\\">Xem bài</a>" : (r.fb_post_id || "");'
    + '      return "<tr>"'
    + '        + "<td style=\\"white-space:nowrap;font-size:11px\\">" + fmtDate(r.created_at) + "</td>"'
    + '        + "<td style=\\"font-size:11px;max-width:120px;overflow:hidden;text-overflow:ellipsis\\">" + src + "</td>"'
    + '        + "<td style=\\"font-size:11px\\">" + postLink + "</td>"'
    + '        + "<td style=\\"font-size:11px\\">" + (r.destination_page_id || "") + "</td>"'
    + '        + "<td>" + res + "</td>"'
    + '        + "<td style=\\"font-size:11px;color:#e53935;max-width:180px\\">" + (r.error_message || "") + "</td>"'
    + '        + "</tr>";'
    + '    }).join("");'
    + '  } else {'
    + '    t += "<tr><td colspan=6 class=empty>Chưa có log</td></tr>";'
    + '  }'
    + '  t += "</tbody></table>";'
    + '  document.getElementById("logs-dash").innerHTML = t;'
    + '  document.getElementById("logs-full").innerHTML = t;'
    + '}'

    // ── Modal
    + 'var FIELDS = {'
    + '  groups: ['
    + '    {k:"group_name", l:"Tên nhóm", req:true, ph:"VD: Nhóm tin tức"},'
    + '    {k:"is_active",  l:"Hoạt động", type:"sel", opts:["TRUE","FALSE"]}'
    + '  ],'
    + '  source_pages: ['
    + '    {k:"group_name",   l:"Thuộc nhóm", type:"groups"},'
    + '    {k:"fb_page_url",  l:"URL Facebook Page", req:true, ph:"https://www.facebook.com/pagename"},'
    + '    {k:"fb_page_name", l:"Tên hiển thị", ph:"VD: VTV Tin tức"},'
    + '    {k:"is_active",    l:"Hoạt động", type:"sel", opts:["TRUE","FALSE"]}'
    + '  ],'
    + '  destination_pages: ['
    + '    {k:"group_name",      l:"Thuộc nhóm", type:"groups"},'
    + '    {k:"fb_page_id",      l:"Page ID (số)", req:true, ph:"VD: 123456789", hint:"Lấy từ me/accounts trên Graph API Explorer"},'
    + '    {k:"fb_page_name",    l:"Tên trang", ph:"Tên page của bạn"},'
    + '    {k:"fb_access_token",     l:"Access Token (Long-lived)", req:true, ph:"EAAxxxxx...", hint:"Xem tab Facebook Token để biết cách lấy"},'
    + '    {k:"is_active",           l:"Hoạt động", type:"sel", opts:["TRUE","FALSE"]},'
    + '    {k:"max_posts_per_run",   l:"Số bài tối đa mỗi lần chạy", ph:"4", hint:"Dù scrape được bao nhiêu, chỉ đăng tối đa số này (bài có tương tác cao nhất được ưu tiên)"},'
    + '    {k:"post_interval_hours", l:"Khoảng cách giữa các bài (giờ)", ph:"2", hint:"VD: 1 = cách 1 giờ, 2 = cách 2 giờ, 0.5 = cách 30 phút"}'
    + '  ],'
    + '  apify_keys: ['
    + '    {k:"api_key",       l:"Apify API Key", req:true, ph:"apify_api_xxx...", hint:"Lấy tại apify.com → Settings → Integrations → API tokens"},'
    + '    {k:"email",         l:"Email tài khoản Apify", ph:"your@email.com"},'
    + '    {k:"monthly_limit", l:"Giới hạn/tháng", ph:"450"}'
    + '  ]'
    + '};'

    + 'var TAB_NAMES = {groups:"Nhóm", source_pages:"Trang nguồn", destination_pages:"Trang đích", apify_keys:"Apify Key"};'

    + 'function openAdd(tab) {'
    + '  MTAB = tab; MKF = null; MKV = null;'
    + '  document.getElementById("m-title").textContent = "➕ Thêm " + TAB_NAMES[tab];'
    + '  document.getElementById("m-body").innerHTML = buildForm(tab, {});'
    + '  document.getElementById("modal").classList.add("on");'
    + '}'

    + 'function openEdit(tab, kf, kv) {'
    + '  MTAB = tab; MKF = kf; MKV = kv;'
    + '  document.getElementById("m-title").textContent = "✏️ Sửa " + TAB_NAMES[tab];'
    + '  var arr = tab==="groups" ? D.groups : tab==="source_pages" ? D.sources : tab==="destination_pages" ? D.dests : D.apify_keys;'
    + '  var row = (arr || []).find(function(r) { return String(r[kf]) === String(kv); }) || {};'
    + '  document.getElementById("m-body").innerHTML = buildForm(tab, row);'
    + '  document.getElementById("modal").classList.add("on");'
    + '}'

    + 'function buildForm(tab, row) {'
    + '  var fields = FIELDS[tab] || [];'
    + '  var html = "";'
    + '  if (tab === "destination_pages") {'
    + '    html += "<div class=info>📌 Cần <b>Page ID</b> và <b>Long-lived Token</b>. Xem hướng dẫn ở tab <b>🔐 Facebook Token</b></div>";'
    + '  }'
    + '  fields.forEach(function(f) {'
    + '    var v = row[f.k] !== undefined ? String(row[f.k]) : "";'
    + '    var inp;'
    + '    if (f.type === "groups") {'
    + '      var opts = (D.groups || []).map(function(g) {'
    + '        return "<option value=\\"" + g.group_name + "\\"" + (g.group_name === v ? " selected" : "") + ">" + g.group_name + "</option>";'
    + '      }).join("");'
    + '      inp = "<select class=\\"fc\\" name=\\"" + f.k + "\\">" + opts + "</select>";'
    + '    } else if (f.type === "sel") {'
    + '      var opts = (f.opts || []).map(function(o) {'
    + '        return "<option value=\\"" + o + "\\"" + (o === v ? " selected" : "") + ">" + o + "</option>";'
    + '      }).join("");'
    + '      inp = "<select class=\\"fc\\" name=\\"" + f.k + "\\">" + opts + "</select>";'
    + '    } else {'
    + '      var safeVal = v.replace(/&/g,"&amp;").replace(/"/g,"&quot;");'
    + '      inp = "<input class=\\"fc\\" name=\\"" + f.k + "\\" value=\\"" + safeVal + "\\" placeholder=\\"" + (f.ph || "") + "\\"" + (f.req ? " required" : "") + ">";'
    + '    }'
    + '    html += "<div class=fg><label>" + f.l + (f.req ? " <span style=\\"color:#e53935\\">*</span>" : "") + "</label>" + inp;'
    + '    if (f.hint) html += "<div class=hint>💡 " + f.hint + "</div>";'
    + '    html += "</div>";'
    + '  });'
    + '  return html;'
    + '}'

    + 'function closeModal() {'
    + '  document.getElementById("modal").classList.remove("on");'
    + '}'

    + 'function saveModal() {'
    + '  var inputs = document.querySelectorAll("#m-body [name]");'
    + '  var row = {}, valid = true;'
    + '  inputs.forEach(function(el) {'
    + '    if (el.required && !el.value.trim()) { el.style.borderColor = "#e53935"; valid = false; }'
    + '    else el.style.borderColor = "";'
    + '    row[el.name] = el.value.trim();'
    + '  });'
    + '  if (!valid) { toast("Vui lòng điền đủ thông tin (*)", "te"); return; }'
    + '  if (MTAB === "apify_keys") { if (!row.monthly_limit) row.monthly_limit = "450"; row.usage_count = "0"; row.is_active = "TRUE"; }'
    + '  var body = {tab: MTAB, row: row};'
    + '  if (MKF && MKV !== null) { body.keyField = MKF; body.keyValue = MKV; }'
    + '  var btn = document.getElementById("btn-modal-save");'
    + '  btn.textContent = "⏳ Đang lưu..."; btn.disabled = true;'
    + '  google.script.run'
    + '    .withSuccessHandler(function(d) {'
    + '      btn.textContent = "💾 Lưu"; btn.disabled = false;'
    + '      if (d && d.ok) { closeModal(); toast("✅ Đã lưu thành công", "ts"); reload(); }'
    + '      else toast("Lỗi: " + (d && d.error || "?"), "te");'
    + '    })'
    + '    .withFailureHandler(function(e) {'
    + '      btn.textContent = "💾 Lưu"; btn.disabled = false;'
    + '      toast("Lỗi: " + e, "te");'
    + '    })'
    + '    .uiSaveRow(body);'
    + '}'

    + 'function delRow(tab, kf, kv) {'
    + '  if (!confirm("Xóa bản ghi này?")) return;'
    + '  google.script.run'
    + '    .withSuccessHandler(function(d) { toast(d.ok ? "✅ Đã xóa" : "Lỗi xóa", d.ok ? "ts" : "te"); if (d.ok) reload(); })'
    + '    .withFailureHandler(function(e) { toast("Lỗi: " + e, "te"); })'
    + '    .uiDeleteRow({tab:tab, keyField:kf, keyValue:kv});'
    + '}'

    + 'function doClearLogs() {'
    + '  if (!confirm("Xóa tất cả log cũ hơn 30 ngày?")) return;'
    + '  google.script.run'
    + '    .withSuccessHandler(function() { toast("✅ Đã xóa log cũ", "ts"); reload(); })'
    + '    .withFailureHandler(function(e) { toast("Lỗi: " + e, "te"); })'
    + '    .uiClearLogs();'
    + '}'

    + 'function toast(msg, cls) {'
    + '  var t = document.getElementById("toast");'
    + '  t.textContent = msg; t.className = "toast " + cls;'
    + '  t.classList.add("on");'
    + '  setTimeout(function() { t.classList.remove("on"); }, 3000);'
    + '}'

    // ── Event delegation (1 listener xử lý tất cả clicks)
    + 'document.addEventListener("click", function(e) {'
    + '  var el = e.target;'

    // Nút chuyển tab (sidebar nav)
    + '  var nav = el.closest(".nav a[data-tab]");'
    + '  if (nav) { showTab(nav.dataset.tab); return; }'

    // Nút chuyển tab (nút thường có data-tab)
    + '  if (el.dataset && el.dataset.tab && !el.closest(".nav")) { showTab(el.dataset.tab); return; }'

    // Nút thêm mới (data-add)
    + '  if (el.dataset && el.dataset.add) { openAdd(el.dataset.add); return; }'

    // Nút edit (data-edit)
    + '  var editBtn = el.closest("[data-edit]");'
    + '  if (editBtn) { openEdit(editBtn.dataset.tab, editBtn.dataset.kf, decodeURIComponent(editBtn.dataset.kv)); return; }'

    // Nút delete (data-del)
    + '  var delBtn = el.closest("[data-del]");'
    + '  if (delBtn) { delRow(delBtn.dataset.tab, delBtn.dataset.kf, decodeURIComponent(delBtn.dataset.kv)); return; }'

    // Nút reload
    + '  if (el.id === "btn-reload") { reload(); return; }'

    // Nút debug
    + '  if (el.id === "btn-debug") {'
    + '    var out = document.getElementById("debug-out");'
    + '    out.style.display = "block"; out.textContent = "⏳ Đang kiểm tra...";'
    + '    google.script.run'
    + '      .withSuccessHandler(function(d) { out.textContent = JSON.stringify(d, null, 2); })'
    + '      .withFailureHandler(function(e) { out.textContent = "LỖI: " + e; })'
    + '      .uiDebug();'
    + '    return;'
    + '  }'
    + '  if (el.id === "btn-testrender") {'
    + '    var out = document.getElementById("debug-out");'
    + '    out.style.display = "block"; out.textContent = "⏳ Gọi uiGetConfig...";'
    + '    google.script.run'
    + '      .withSuccessHandler(function(cfg) {'
    + '        out.textContent = "uiGetConfig OK\\ngroups.length=" + (cfg.groups||[]).length + "\\n" + JSON.stringify(cfg.groups, null, 2);'
    + '        D = Object.assign(D||{}, cfg);'
    + '        renderGroups(); renderSources(); renderDests(); renderKeys();'
    + '        document.getElementById("s0").textContent = (D.groups||[]).length;'
    + '        document.getElementById("s1").textContent = (D.sources||[]).length;'
    + '        document.getElementById("s2").textContent = (D.dests||[]).length;'
    + '        document.getElementById("s3").textContent = D.dedup_count || 0;'
    + '      })'
    + '      .withFailureHandler(function(e) { out.textContent = "LỖI uiGetConfig: " + e; })'
    + '      .uiGetConfig();'
    + '    return;'
    + '  }'

    // Nút clear logs
    + '  if (el.id === "btn-clearlogs" || el.id === "btn-clearlogs2") { doClearLogs(); return; }'

    // Nút đóng modal
    + '  if (el.id === "btn-modal-close" || el.id === "btn-modal-cancel") { closeModal(); return; }'

    // Nút lưu modal
    + '  if (el.id === "btn-modal-save") { saveModal(); return; }'

    // Click backdrop modal
    + '  if (el.id === "modal") { closeModal(); return; }'
    + '});'

    + 'renderGroups(); renderSources(); renderDests(); renderKeys();'
    + 'document.getElementById("s0").textContent = (D.groups||[]).length;'
    + 'document.getElementById("s1").textContent = (D.sources||[]).length;'
    + 'document.getElementById("s2").textContent = (D.dests||[]).length;'
    + 'document.getElementById("s3").textContent = D.dedup_count||0;'
    + 'document.getElementById("logs-dash").innerHTML = "<div class=empty>Nhấn 🔄 để tải logs</div>";'
    + 'reload();'
    + '</script>';

  return html + js + '</body></html>';
}
