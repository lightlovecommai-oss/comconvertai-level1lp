/**
 * 短影音課・課前調查的後端。收姓名（必填）＋email（選填）＋五題答案，回傳累計人數。
 *
 * ── 部署（做一次，五分鐘）────────────────────────────────
 * 1. 開一份新的 Google 試算表（名字隨便，例如「短影音課前調查」）。
 * 2. 該試算表 → 擴充功能 → Apps Script，把這整份貼進去，蓋掉原本的 myFunction。
 * 3. 部署 → 新增部署作業 → 類型選「網頁應用程式」。
 *      執行身分＝我
 *      具有存取權的使用者＝**任何人**（選錯成「只有我」就收不到任何一筆）
 * 4. 複製那條結尾是 exec 的網址，貼進 index.html 最上面的 SHEET_API。
 * 5. 改完題目或欄位要重新部署（部署 → 管理部署作業 → 鉛筆 → 版本選「新版本」）。
 *
 * 容器綁定寫法（SpreadsheetApp.getActive）＝不用填 spreadsheet ID，
 * 也就沒有一組 ID 散在 repo 裡；這頁掛在公開的 GitHub Pages，能少放一個就少放一個。
 */

var SHEET = '課前調查';
// 姓名必填、email 選填（這群人不少沒有 email）。姓名排在 email 前面＝老師讀表時先看到人。
var HEAD = ['時間', '姓名', 'email', '腳本', '拍攝', '剪輯', '已發幾支', '最想解決'];

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < HEAD.length) {
    // 加欄後重新部署時把標題列補齊，不然新欄位是空白表頭、看表的人不知道那是什麼
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function count_() {
  return Math.max(0, sheet_().getLastRow() - 1);
}

function doPost(e) {
  // 兩個人同時按送出會搶同一個 getLastRow，沒有鎖就會蓋掉彼此那一列
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var name = String(d.name || '').trim();
    var mail = String(d.email || '').trim();
    var row = [new Date(), name, mail, d.q1 || '', d.q2 || '', d.q3 || '', d.q4 || '', d.q5 || ''];
    var sh = sheet_();

    /* 同一個人重填就改掉原本那列，不新增——不然人數會被重複填答灌水。
       認人的鑰匙：有 email 就用 email（唯一），沒有就退回姓名。
       ⚠️ 退回姓名時，兩個同名的人會蓋掉彼此。三十人的班同名機率低，
       而為了防這個去要求每個人都有 email，等於把沒信箱的人擋在門外——
       那才是真正的損失。寧可事後手動拆，不要事前擋人。 */
    var key = (mail || name).toLowerCase();
    var col = mail ? 3 : 2;
    var at = -1;
    if (key && sh.getLastRow() > 1) {
      var vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        var v = String(vals[i][0]).trim().toLowerCase();
        if (v && v === key) { at = i; break; }
      }
    }
    if (at >= 0) {
      sh.getRange(at + 2, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return json_({ status: 'ok', count: count_() });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* JSONP 的 callback 名字會被原樣印進回應裡，等於把網址上的字串當程式碼送出去，
   所以只放行英數底線錢號。刻意用迴圈不用正規表示式：貼進 Apps Script 編輯器時，
   regex 字面值是最容易被貼壞、然後整支檔案編譯失敗的一段。 */
function safeCb_(s) {
  s = String(s || '');
  if (!s || s.length > 40) return '';
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    var ok = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') ||
             (c >= '0' && c <= '9') || c === '_' || c === '$';
    if (!ok) return '';
  }
  return s;
}

/** 只回人數，不回任何一筆答案或 email——這支網址是公開的。 */
function doGet(e) {
  var out = { status: 'ok', count: count_() };
  var cb = safeCb_(e && e.parameter && e.parameter.callback);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(out);
}
