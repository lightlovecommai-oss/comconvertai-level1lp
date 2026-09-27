/**
 * 10/29 Lily 沙龍會講師培訓・出席確認的後端。
 * 收姓名（必填）＋會不會到＋人數＋email（選填），回傳累計報名數。
 *
 * ── 部署（做一次，五分鐘）────────────────────────────────
 * 1. 開一份新的 Google 試算表（例如「Lily 1029 報名」）。
 *    🔴 開新的，不要沿用天麗那份——不同活動、不同生命週期，混在一起只會互相干擾。
 * 2. 該試算表 → 擴充功能 → Apps Script，把這整份貼進去，蓋掉原本的 myFunction。
 * 3. 部署 → 新增部署作業 → 類型選「網頁應用程式」。
 *      執行身分＝我
 *      具有存取權的使用者＝**任何人**
 *      🔴 選錯成「只有我」時端點回 403 登入頁，報名的人全部送不出去，
 *         **而頁面看起來一切正常**——這是整條流程最會出錯的一顆。
 * 4. 複製那條結尾是 exec 的網址，貼進 index.html 最上面的 SHEET_API。
 * 5. 改完欄位或題目要重新部署（部署 → 管理部署作業 → 鉛筆 → 版本選「新版本」）。
 *
 * 容器綁定寫法（SpreadsheetApp.getActive）＝不用填 spreadsheet ID，
 * 也就沒有一組 ID 散在 repo 裡；這頁掛在公開的 GitHub Pages，能少放一個就少放一個。
 *
 * ⚠️ 這份檔**刻意不使用任何正規表示式字面值**——貼進 Apps Script 編輯器時
 *    regex 是最容易被貼壞、然後整支編譯失敗的一段（2026-09-26 在天麗那支踩過一次）。
 */

var SHEET = '1029報名';

/* 欄位順序是刻意的：姓名（2）與 email（5）**分開站在兩端**，
   中間夾著「出席」（3）與「人數」（4）。這樣統計端點只要框第 3–4 欄，
   就能一次把名單那兩欄擋在記憶體外面——見 stats_() 的說明。 */
var HEAD = ['時間', '姓名', '出席', '人數', 'email'];

var COL_NAME = 2, COL_GOING = 3, COL_COUNT = 4, COL_MAIL = 5;

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < HEAD.length) {
    // 加欄後重新部署時補齊標題列，不然新欄位是空白表頭、看表的人不知道那是什麼
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

/* 「4 位以上」沒有精確值，一律當 4 算。
   場地抓位子寧可少估一點，現場多來一個人搬張椅子就好；
   反過來（把它當 6 算）會讓老師以為位子不夠而去換場地。 */
function seats_(v) {
  var s = String(v || '').trim();
  if (!s) return 0;
  var n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}

/* 統計：這場是出席確認，老師要的是**人頭數**不是選項分佈。
   回傳：報名筆數、會到幾組／幾人、來不了幾組。

   🔴 這支端點是公開的（沒有金鑰）。所以 getRange 只框第 3–4 欄（出席／人數），
      **姓名與 email 那兩欄根本不會被讀進記憶體**——「不會漏名單」是讀取範圍的
      結構保證，不是「我記得沒把它放進回傳值」。之後若有人要加功能，
      想漏也得先動這行 getRange，會被看見。
      要做人名對照就另開一支端點、那支才上金鑰，不要把名單塞回 stats。 */
function stats_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  var out = { status: 'ok', total: 0, goingGroups: 0, goingPeople: 0, cantCome: 0, updatedAt: Date.now() };
  if (last <= 1) return out;

  var width = COL_COUNT - COL_GOING + 1;                       // 2 ＝出席與人數相鄰兩欄
  var rows = sh.getRange(2, COL_GOING, last - 1, width).getValues();
  out.total = rows.length;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === '我會到') {
      out.goingGroups += 1;
      out.goingPeople += seats_(rows[i][1]);
    } else {
      out.cantCome += 1;
    }
  }
  return out;
}

function doPost(e) {
  // 兩個人同時按送出會搶同一個 getLastRow，沒有鎖就會蓋掉彼此那一列
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var name = String(d.name || '').trim();
    var mail = String(d.email || '').trim();
    var going = String(d.going || '').trim();
    var cnt = String(d.count || '').trim();
    var row = [new Date(), name, going, cnt, mail];
    var sh = sheet_();

    /* 同一個人改主意就改掉原本那列，不新增——不然人數會被重複報名灌水，
       而且「本來說會到、後來改成來不了」必須蓋過舊答案，不能兩列並存。
       認人的鑰匙：有 email 就用 email（唯一），沒有就退回姓名。
       ⚠️ 退回姓名時兩個同名的人會蓋掉彼此。這種規模的班同名機率低，
       而為了防這個去要求每個人都有 email，等於把沒信箱的人擋在門外——
       那才是真正的損失。寧可事後手動拆，不要事前擋人。 */
    var key = (mail || name).toLowerCase();
    var col = mail ? COL_MAIL : COL_NAME;
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

/** 預設只回報名筆數；action=stats 回人頭統計（老師抓場地位子用）。
    兩者都不回姓名或 email——GET 這條路上沒有任何一個分支碰得到那兩欄。 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var cb = safeCb_(params.callback);
  var out = String(params.action || '') === 'stats'
    ? stats_()
    : { status: 'ok', count: count_() };
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(out);
}
