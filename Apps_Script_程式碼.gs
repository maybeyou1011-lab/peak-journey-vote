/**
 * 巔峰之旅 日期投票 —— 後端程式（貼到 Google Apps Script）
 *
 * 白話說明：
 *   這段程式幫網頁「收票」和「算票」。
 *   - 有人投票 → doPost 會把票寫進試算表（同名的人會更新，不會重複計算）
 *   - 網頁要看結果 → doGet 會算出每個日期幾票回傳
 *   你完全不用看懂程式碼，照「設定說明.txt」貼上、部署就好。
 */

// 試算表第一列的欄位標題（程式會自動建立）
var HEADERS = ['時間', '姓名', '隊別', '選的日期'];

/** 取得（或建立）存票的工作表 */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('投票');
  if (!sh) {
    sh = ss.insertSheet('投票');
    sh.appendRow(HEADERS);
  }
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

/** 有人投票時（網頁送 POST 過來） */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // 避免同時多人投票打架
  try {
    var sh = getSheet_();
    var name  = (e.parameter.name  || '').toString().trim();
    var team  = (e.parameter.team  || '').toString().trim();
    var dates = (e.parameter.dates || '').toString().trim();
    if (!name) return json_({ ok: false, msg: 'no name' });

    // 同一個名字已經投過 → 更新那一列；沒有 → 新增一列
    var data = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if ((data[i][1] || '').toString().trim() === name) { rowIndex = i + 1; break; }
    }
    var now = new Date();
    var row = [now, name, team, dates];
    if (rowIndex > 0) {
      sh.getRange(rowIndex, 1, 1, 4).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, msg: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 網頁要看票數時（送 GET 過來，用 JSONP 回傳） */
function doGet(e) {
  var sh = getSheet_();
  var data = sh.getDataRange().getValues();
  var tally = {};   // 每個日期幾票
  var total = 0;    // 投票人數
  for (var i = 1; i < data.length; i++) {
    var name = (data[i][1] || '').toString().trim();
    if (!name) continue;
    total++;
    var picks = (data[i][3] || '').toString().split(',');
    for (var j = 0; j < picks.length; j++) {
      var d = picks[j].trim();
      if (!d) continue;
      tally[d] = (tally[d] || 0) + 1;
    }
  }
  var payload = { ok: true, tally: tally, total: total };

  // JSONP：網頁用 callback 參數呼叫，避開跨網域讀取限制
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

/** 小工具：回傳 JSON */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
