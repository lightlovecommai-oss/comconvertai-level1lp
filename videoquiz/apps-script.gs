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

/* 統計後台的題目與選項──跟 index.html 的 QS 必須對齊。
   選項順序固定＝圖表 3 個切片永遠同色同順序，即使某選項還是 0 也保留切片位置。
   改題目：這裡＋index.html QS 一起改，然後重新部署。 */
var STATS_Q = [
  { key: 'q1', col: 4, tag: '腳本',     q: '你知道怎麼寫出一支「會有人看」的腳本嗎？',
    opts: ['完全沒概念', '大概知道，但自己寫不出來', '可以自己寫出來'] },
  { key: 'q2', col: 5, tag: '拍攝',     q: '你可以對著鏡頭，把一段話講完嗎？',
    opts: ['會卡住，講不下去', '講得出來，但要重錄很多次', '一兩次就過'] },
  { key: 'q3', col: 6, tag: '剪輯',     q: '你會用剪輯軟體嗎？',
    opts: ['沒裝過', '裝了，會一點點', '可以自己剪完發出去'] },
  { key: 'q4', col: 7, tag: '現況',     q: '到目前為止，你發過幾支短影音？',
    opts: ['0 支', '1 到 3 支', '4 支以上'] },
  { key: 'q5', col: 8, tag: '最想解決', q: '這三個裡面，你最想先解決哪一個？',
    opts: ['腳本寫不出來', '不敢對鏡頭', '剪輯不會用'] }
];

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

/* 統計：逐題逐選項數，回傳彙總計數。「其他」欄收沒對到選項的雜訊
   （改題但沒重新部署、手動貼進去的舊值⋯）。

   🔴 這支端點是公開的（沒有金鑰）。所以 getRange 只框答案欄（第 4–8 欄），
      **姓名與 email 那兩欄根本不會被讀進記憶體**——「不會漏名單」是讀取範圍的
      結構保證，不是「我記得沒把它放進回傳值」。之後若有人要加功能，
      想漏也得先動這行 getRange，會被看見。 */
function stats_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  var first = STATS_Q[0].col;                             // 4 ＝「腳本」欄
  var width = STATS_Q[STATS_Q.length - 1].col - first + 1; // 5 ＝五題答案連續佔五欄
  var q = STATS_Q.map(function (item) {
    var counts = {};
    item.opts.forEach(function (o) { counts[o] = 0; });
    counts.__other = 0;
    return { key: item.key, tag: item.tag, q: item.q, opts: item.opts.slice(), counts: counts };
  });
  if (last > 1) {
    // 一次讀完答案區——比逐格 getRange 快很多，且同一批視角避免中途新列造成偏誤
    var range = sh.getRange(2, first, last - 1, width).getValues();
    for (var i = 0; i < range.length; i++) {
      var row = range[i];
      for (var j = 0; j < STATS_Q.length; j++) {
        var v = String(row[STATS_Q[j].col - first] || '').trim();
        if (!v) continue;
        var bucket = q[j].counts;
        if (bucket[v] === undefined) bucket.__other += 1;
        else bucket[v] += 1;
      }
    }
  }
  return {
    status: 'ok',
    total: Math.max(0, last - 1),
    updatedAt: Date.now(),
    q: q
  };
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

/** 預設只回人數（填答頁的社會證明）；action=stats 回選項分佈（統計頁）。
    兩者都不回姓名或 email——GET 這條路上沒有任何一個分支碰得到那兩欄，
    stats_() 的 getRange 也只框答案欄。 */
/* 一次性修髒列——**手動跑，不掛在任何端點上**。
   在 Apps Script 編輯器上方的函式下拉選 cleanupLegacyRows → 執行，看執行紀錄。

   要修兩種列：
   ① 欄位左移一格：2026-09-26 加「姓名」欄之前填的那幾列，email 落在姓名欄、
      最後一題掉出表外。認法＝姓名欄裡有 "@"。把 B~G 往右搬一格、B 清空。
      ⚠️ 姓名救不回來（那時根本沒收），只能留白。
   ② 老師自己的「光頭測試」列，五題全選最差，會把統計拉向「這群人什麼都不會」。

   刻意由下往上跑：刪列會讓下面的列號整個往上移，由上往下刪會跳過列。
   重複執行是安全的——修過的列姓名欄已經沒有 "@"，測試列已經不在。
   跑完可以把這支函式整個刪掉，它不是長期資產。 */
function cleanupLegacyRows() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return;

  var shifted = 0, removed = 0;
  var rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    var rowNum = i + 2;
    var nameCell = String(rows[i][1] || '').trim();

    if (nameCell === '光頭測試') {
      sh.deleteRow(rowNum);
      removed += 1;
    } else if (nameCell.indexOf('@') >= 0) {
      // B~G（索引 1~6）往右搬一格變成 C~H，姓名欄留白
      var fixed = [rows[i][0], ''].concat(rows[i].slice(1, HEAD.length - 1));
      sh.getRange(rowNum, 1, 1, HEAD.length).setValues([fixed]);
      shifted += 1;
    }
  }

  Logger.log('修好左移 ' + shifted + ' 列、刪掉測試 ' + removed + ' 列，現在共 ' + count_() + ' 筆。');
}

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
