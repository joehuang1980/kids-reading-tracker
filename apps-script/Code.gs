/**
 * Kids Reading Tracker — Google Apps Script Web App
 * Stage 6: 在 Google Sheet 建立名為「Kids Reading Records」的試算表
 *          新增工作表「Reading_Log」，欄位如下：
 *          A: Record_ID  B: ISBN  C: Book_Title  D: Reading_Date
 *          E: Reader     F: Created_At  G: Thumbnail  H: Word_Count
 *          （G 欄可能是網址，或使用者自拍書封的 JPEG data URL；
 *            H 欄是英文書在 AR BookFind 查到的總字數，舊試算表會自動補上標題）
 *
 * Stage 7: 部署步驟
 *  1. 開啟 https://script.google.com，新增專案
 *  2. 貼上此程式碼，存檔
 *  3. 「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *     - 以以下身分執行：「我」
 *     - 誰可以存取：「所有人」
 *  4. 複製 Web App URL，填入 js/config.js 的 APPS_SCRIPT_URL
 *  5. 「專案設定」勾選「在編輯器中顯示 appsscript.json」，
 *     內容換成本資料夾的 appsscript.json（明確列出需要的權限）
 *  6. 在編輯器上方選擇函式「testLookup」並按「執行」，依提示授權
 *     （需允許「連線至外部服務」，中文書查詢才能連到國圖與博客來；
 *       未授權時 debug=1 會顯示「沒有呼叫 UrlFetchApp.fetch 的權限」）
 *     英文書字數查詢可在編輯器執行「testArLookup」測試。
 *
 *  7.（選用）Google Books API 金鑰：「專案設定」→「指令碼屬性」新增
 *     屬性 GOOGLE_BOOKS_API_KEY，值填金鑰。金鑰只存在這裡，不要寫進 js/config.js
 *     （網站程式是公開的，任何人都能看到並盜用額度）。
 *
 *  ※ 更新程式碼後須「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」，
 *    前端才會用到新功能（例如中文書查詢 action=lookup）。
 *
 * Stage 10: Excel 匯出 — 在 Google Sheet 選「檔案」→「下載」→「Microsoft Excel (.xlsx)」
 */

var SHEET_NAME = 'Reading_Log';
var HEADERS    = ['Record_ID', 'ISBN', 'Book_Title', 'Reading_Date', 'Reader', 'Created_At', 'Thumbnail', 'Word_Count'];
var NUM_COLS   = HEADERS.length;
var MAX_CELL_CHARS = 49000;   // Google 試算表單一儲存格上限約 50,000 字元

function getSheet() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (!sheet.getRange(1, NUM_COLS).getValue()) {
    sheet.getRange(1, NUM_COLS).setValue(HEADERS[NUM_COLS - 1]);  // 舊試算表補上新欄位標題
  }
  return sheet;
}

// Stage 8: Receive reading record from frontend
function doPost(e) {
  try {
    var sheet = getSheet();
    var data  = JSON.parse(e.postData.contents);

    // Validation
    if (!data.title || !data.readingDate || !data.reader) {
      return _json({ success: false, message: 'Missing required fields' });
    }

    // Stage 9: Duplicate check — same ISBN + reader + date within 5 minutes
    if (data.isbn) {
      var lr = sheet.getLastRow();
      if (lr > 1) {
        var checkRows = Math.min(20, lr - 1);
        var rows = sheet.getRange(Math.max(2, lr - checkRows + 1), 1, checkRows, 7).getValues();
        var now  = new Date();
        for (var i = 0; i < rows.length; i++) {
          var r       = rows[i];
          var rowTime = r[5] instanceof Date ? r[5] : new Date(r[5]);
          // r[3] may be a Date object; normalise to yyyy-MM-dd string for comparison
          var rowDate = r[3] instanceof Date
            ? Utilities.formatDate(r[3], 'Asia/Taipei', 'yyyy-MM-dd') : String(r[3]);
          if (String(r[1]) === String(data.isbn) && rowDate === data.readingDate && r[4] === data.reader) {
            if ((now - rowTime) < 300000) {  // within 5 minutes
              return _json({ success: false, duplicate: true, message: '重複紀錄' });
            }
          }
        }
      }
    }

    var recordId = 'R-' +
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd') + '-' +
      Math.random().toString(36).substr(2, 4).toUpperCase();

    var thumbnail = String(data.thumbnail || '');
    if (thumbnail.length > MAX_CELL_CHARS) thumbnail = '';  // 超過儲存格上限就不存封面，避免整筆失敗
    var wordCount = parseInt(data.wordCount, 10);

    sheet.appendRow([
      recordId,
      data.isbn      || '',
      data.title,
      data.readingDate,
      data.reader,
      new Date(),
      thumbnail,
      wordCount > 0 ? wordCount : '',
    ]);

    return _json({ success: true, recordId: recordId });
  } catch (err) {
    return _json({ success: false, message: String(err) });
  }
}

// Stage 11–12: Return all records for history / dashboard pages
function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.action === 'lookup') return _json(lookupBook(params.isbn, params.debug === '1'));
  if (params.action === 'ar')     return _json(lookupArBook(params.isbn, params.debug === '1'));
  try {
    var sheet = getSheet();
    var lr    = sheet.getLastRow();
    if (lr < 2) return _json({ records: [] });

    var rows    = sheet.getRange(2, 1, lr - 1, NUM_COLS).getValues();
    var records = rows
      .filter(function(r) { return r[2]; })  // skip rows without title
      .map(function(r) {
        return {
          id:          r[0],
          isbn:        String(r[1] || ''),  // Sheets stores ISBN as number
          title:       r[2],
          readingDate: r[3] instanceof Date
            ? Utilities.formatDate(r[3], 'Asia/Taipei', 'yyyy-MM-dd') : String(r[3]),
          reader:      r[4],
          createdAt:   r[5] instanceof Date ? r[5].toISOString() : String(r[5]),
          thumbnail:   r[6] || '',
          wordCount:   Number(r[7]) || null,
        };
      });

    return _json({ records: records });
  } catch (err) {
    return _json({ success: false, message: String(err) });
  }
}

// ─── 中文書查詢（全國新書資訊網 → 博客來）──────────────────────────────────
// 瀏覽器無法直接跨網域讀這些網站，因此由 Apps Script 代為查詢。
// 回傳 { book: { title, authors, publisher, thumbnail, source } } 或 { book: null }
// 診斷：在瀏覽器開啟 <Web App URL>?action=lookup&isbn=978...&debug=1
//       會略過快取，並在 trace 中列出每個網址的回應狀態與內容片段。
var LOOKUP_VERSION = 4;
var _trace = null;

function lookupBook(isbn, debug) {
  isbn = String(isbn || '').replace(/[^0-9Xx]/g, '').toUpperCase();
  if (!/^(\d{13}|\d{9}[0-9X])$/.test(isbn)) return { book: null, message: 'invalid isbn' };

  _trace = debug ? [] : null;
  var cache = CacheService.getScriptCache();
  var hit   = debug ? null : cache.get('isbn:' + isbn);
  if (hit) return { book: JSON.parse(hit), v: LOOKUP_VERSION };

  var book = null;
  try { book = _lookupNcl(isbn); } catch (err) { _note('ncl error: ' + err); }
  if (!book || !book.title) {
    try { book = _lookupGoogleBooks(isbn); } catch (err) { _note('google books error: ' + err); }
  }
  if (!book || !book.title) {
    try { book = _lookupBooksComTw(isbn); } catch (err) { _note('books.com.tw error: ' + err); }
  }

  var out = { book: (book && book.title) ? book : null, v: LOOKUP_VERSION };
  if (out.book) cache.put('isbn:' + isbn, JSON.stringify(book), 21600);  // 6 小時
  if (_trace) out.trace = _trace;
  return out;
}

// 在 Apps Script 編輯器中選此函式按「執行」，可於執行記錄看到診斷結果
function testLookup() {
  Logger.log(JSON.stringify(lookupBook('9786264179645', true), null, 2));
}

function _note(msg) {
  Logger.log(msg);
  if (_trace) _trace.push(msg);
}

// UrlFetchApp.fetch 包裝：診斷模式下記錄每次請求
function _fetch(url, opts) {
  var resp = UrlFetchApp.fetch(url, opts);
  if (_trace) {
    var text = '';
    try { text = resp.getContentText('UTF-8'); } catch (e) {}
    var h = resp.getAllHeaders();
    _trace.push({
      url:      url.replace(/([?&]key=)[^&]+/, '$1***'),  // 不在診斷結果洩漏金鑰
      code:     resp.getResponseCode(),
      bytes:    text.length,
      location: h['Location'] || h['location'] || undefined,
      snippet:  text.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ')
                    .replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 400),
    });
  }
  return resp;
}

var _FETCH_OPTS = {
  muteHttpExceptions: true,
  followRedirects: true,
  headers: {
    'User-Agent': 'Mozilla/5.0 (iPad; CPU OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1',
    'Accept-Language': 'zh-TW,zh;q=0.9',
  },
};

// ─── 英文書字數（AR BookFind）─────────────────────────────────────────────
// 流程：選擇使用者身分（usertype.aspx）→ 首頁快速搜尋 ISBN → 書籍詳細頁讀取字數。
// 欄位 ID 參考 github.com/anthonator/ar-book-finder（如 lblWordCount、txtKeyWords、btnDoIt）。
// 回傳 { ar: { wordCount, title, author, atosLevel, arPoints, site } } 或 { ar: null }
// 診斷：<Web App URL>?action=ar&isbn=978...&debug=1
var AR_SITES = ['https://www.arbookfind.com/', 'https://www.arbookfind.co.uk/'];

function lookupArBook(isbn, debug) {
  isbn = String(isbn || '').replace(/[^0-9Xx]/g, '').toUpperCase();
  if (!/^(\d{13}|\d{9}[0-9X])$/.test(isbn)) return { ar: null, message: 'invalid isbn' };

  _trace = debug ? [] : null;
  var cache = CacheService.getScriptCache();
  var hit   = debug ? null : cache.get('ar:' + isbn);
  if (hit) return { ar: JSON.parse(hit), v: LOOKUP_VERSION };

  // AR 可能只登錄 ISBN-10 或 ISBN-13，兩種都試
  var queries = [isbn];
  var alt = isbn.length === 13 ? _isbn13to10(isbn) : _isbn10to13(isbn);
  if (alt) queries.push(alt);

  var result = null;
  for (var i = 0; i < AR_SITES.length && !result; i++) {
    for (var j = 0; j < queries.length && !result; j++) {
      try { result = _arSearch(AR_SITES[i], queries[j]); }
      catch (err) { _note('ar ' + AR_SITES[i] + ' error: ' + err); }
    }
  }

  // 查到快取 6 小時；查不到快取 1 小時，避免重複打擾網站
  cache.put('ar:' + isbn, JSON.stringify(result), result ? 21600 : 3600);
  var out = { ar: result, v: LOOKUP_VERSION };
  if (_trace) out.trace = _trace;
  return out;
}

// 在 Apps Script 編輯器中選此函式按「執行」，可於執行記錄看到診斷結果
function testArLookup() {
  var r = lookupArBook('9780439708180', true);  // Harry Potter and the Sorcerer's Stone
  // 精簡的記錄，避免執行記錄被截斷；完整內容可用網址 ?action=ar&isbn=…&debug=1 查看
  Logger.log('結果：' + JSON.stringify(r.ar));
  (r.trace || []).forEach(function(t) {
    if (typeof t === 'string') Logger.log('・' + t);
    else if (t.url) Logger.log(t.code + ' ' + t.url + (t.location ? ' → ' + t.location : '') + (t.code === 200 ? '  「' + t.snippet.slice(0, 80) + '」' : ''));
    else if (t.postFields) Logger.log('   送出：' + Object.keys(t.postFields).join(', '));
  });
}

// 網站偶爾會回「An error has occurred. Please try your request again.」
// （同樣的表單內容再送一次就會成功），所以選身分、搜尋這兩步各最多試 AR_TRIES 次。
var AR_TRIES = 3;

function _arIsError(s) {
  return /bookfindererror/i.test(s.lastUrl);
}

// 選擇使用者身分，成功回傳之後拿到的頁面 HTML；失敗回傳 null
function _arChooseUserType(s, base) {
  for (var t = 0; t < AR_TRIES; t++) {
    if (t > 0) { _note('ar: 選擇身分失敗，重試第 ' + t + ' 次'); Utilities.sleep(500); }
    var html = s.get(base + 'usertype.aspx');
    if (t === 0) _traceInputs('usertype', html);
    if (!html || _arIsError(s)) continue;
    if (!/type=["']?radio/i.test(html)) return html;   // 已經選過身分

    var form  = _aspForm(html);
    var radio = _pickRadio(html, [/parent/i, /student/i]);
    if (radio) form.fields[radio.name] = radio.value;
    _addSubmit(html, form.fields, null);
    html = s.post(_absUrl(form.action || 'usertype.aspx', s.lastUrl), form.fields);
    if (html && !_arIsError(s)) return html;
  }
  return null;
}

// 快速搜尋 ISBN，成功回傳搜尋結果（或詳細頁）HTML；失敗回傳 null
function _arQuickSearch(s, base, isbn, html) {
  for (var t = 0; t < AR_TRIES; t++) {
    if (t > 0) { _note('ar: 搜尋失敗，重試第 ' + t + ' 次'); Utilities.sleep(500); }
    if (!html || html.indexOf('txtKeyWords') < 0) html = s.get(base + 'default.aspx');
    if (!html || html.indexOf('txtKeyWords') < 0) html = _arChooseUserType(s, base);  // 身分失效就重選
    if (!html || html.indexOf('txtKeyWords') < 0) { html = null; continue; }
    if (t === 0) _traceInputs('search', html);

    var form = _aspForm(html);
    var kw   = _inputById(html, 'txtKeyWords');
    form.fields[kw ? kw.name : 'ctl00$ContentPlaceHolder1$txtKeyWords'] = isbn;
    _addSubmit(html, form.fields, 'btnDoIt');
    var result = s.post(_absUrl(form.action || 'default.aspx', s.lastUrl), form.fields);
    if (result && !_arIsError(s)) return result;
    html = null;   // 下一次重新取得搜尋表單
  }
  return null;
}

function _arSearch(base, isbn) {
  var s = _session('en-US,en;q=0.9');

  // 1. 選擇使用者身分（第一次進站時網站會要求）
  var html = _arChooseUserType(s, base);
  if (!html) { _note('ar: ' + base + ' 選擇身分一直失敗'); return null; }

  // 2. 快速搜尋
  html = _arQuickSearch(s, base, isbn, html);
  if (!html) { _note('ar: ' + base + ' 搜尋一直失敗'); return null; }

  // 3. 書籍詳細頁（只有一筆結果時，網站可能直接顯示詳細頁）
  if (!/lblWordCount/.test(html)) {
    var link = html.match(/href=["']([^"']*bookdetail\.aspx[^"']*)["']/i);
    if (!link) { _note('ar: ' + isbn + ' 搜尋結果沒有書籍'); return null; }
    html = s.get(_absUrl(_decode(link[1]), s.lastUrl));
    if (!html) { _note('ar: 詳細頁沒有回應'); return null; }
  }

  var wordCount = parseInt(_textById(html, 'lblWordCount').replace(/[^\d]/g, ''), 10) || null;
  var title     = _arTitle(html);
  if (!wordCount && !title) { _note('ar: 詳細頁找不到字數與書名'); return null; }

  return {
    wordCount: wordCount,
    title:     title,
    author:    _textById(html, 'lblAuthor').replace(/^by\s+/i, ''),
    atosLevel: _textById(html, 'lblBookLevel'),
    arPoints:  _textById(html, 'lblPoints'),
    site:      base,
  };
}

// 書名：先找 id 含 Title 的元素（如 lblBookTitle），再找 detailTable 內第一個 <strong>
function _arTitle(html) {
  var clean = function(t) { return _decode(t.replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim(); };
  var re = /<(span|a|td|div|h\d|strong)\b[^>]*id=["'][^"']*title[^"']*["'][^>]*>([\s\S]*?)<\/\1>/gi, m;
  while ((m = re.exec(html))) {
    var t = clean(m[2]);
    if (t) return t;
  }
  var i = html.search(/id=["'][^"']*detailTable["']/i);
  if (i < 0) return '';
  m = html.slice(i).match(/<strong[^>]*>([\s\S]*?)<\/strong>/i);
  return m ? clean(m[1]) : '';
}

// ─── 診斷用：列出頁面上的表單欄位、送出的欄位（長字串只顯示長度）───
function _summarizeFields(fields) {
  var out = {};
  Object.keys(fields).forEach(function(k) {
    var v = String(fields[k]);
    out[k] = v.length > 40 ? '(' + v.length + ' chars)' : v;
  });
  return out;
}

function _traceInputs(label, html) {
  if (!_trace || !html) return;
  var forms = (html.match(/<form\b[^>]*>/gi) || []).map(function(t) {
    return { id: _attr(t, 'id'), action: _attr(t, 'action') };
  });
  var tags = (html.match(/<(input|select|button)\b[^>]*>/gi) || []).map(function(t) {
    var v = _attr(t, 'value') || '';
    return [(t.match(/^<(\w+)/) || [])[1], _attr(t, 'type'), _attr(t, 'name'), _attr(t, 'id'),
            v.length > 30 ? '(' + v.length + ' chars)' : v].join(' | ');
  });
  var postbacks = (html.match(/__doPostBack\([^)]*\)/g) || []).slice(0, 10);
  _trace.push({ page: label, forms: forms, inputs: tags, postbacks: postbacks });
}

// ─── ASP.NET 表單小工具 ───
function _attr(tag, name) {
  var m = tag.match(new RegExp('\\s' + name + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\'|([^\\s>]+))', 'i'));
  return m ? _decode(m[1] !== undefined ? m[1] : (m[2] !== undefined ? m[2] : m[3])) : null;
}

function _inputs(html, type) {
  var tags = html.match(/<input\b[^>]*>/gi) || [];
  return tags.filter(function(t) { return (_attr(t, 'type') || 'text').toLowerCase() === type; })
             .map(function(t) { return { name: _attr(t, 'name'), value: _attr(t, 'value') || '', id: _attr(t, 'id') || '', tag: t }; })
             .filter(function(x) { return x.name; });
}

// 取得第一個 <form> 的 action 與所有 hidden 欄位（__VIEWSTATE、__EVENTVALIDATION 等）
function _aspForm(html) {
  var formTag = (html.match(/<form\b[^>]*>/i) || [''])[0];
  var fields  = {};
  _inputs(html, 'hidden').forEach(function(x) { fields[x.name] = x.value; });
  return { action: formTag ? _attr(formTag, 'action') : '', fields: fields };
}

function _inputById(html, idPart) {
  var tags = html.match(/<input\b[^>]*>/gi) || [];
  for (var i = 0; i < tags.length; i++) {
    var id = _attr(tags[i], 'id') || '';
    if (id.indexOf(idPart) >= 0) return { name: _attr(tags[i], 'name'), value: _attr(tags[i], 'value') || '' };
  }
  return null;
}

function _pickRadio(html, prefs) {
  var radios = _inputs(html, 'radio');
  for (var i = 0; i < prefs.length; i++) {
    for (var j = 0; j < radios.length; j++) {
      if (prefs[i].test(radios[j].id) || prefs[i].test(radios[j].value)) return radios[j];
    }
  }
  return radios[0] || null;
}

// 加上送出按鈕；沒有 submit 按鈕時改用 __doPostBack 的目標
function _addSubmit(html, fields, idPart) {
  var btn = idPart ? _inputById(html, idPart) : null;
  if (!btn) {
    var subs = _inputs(html, 'submit').concat(_inputs(html, 'image'));
    // 優先選名稱或文字像「送出」的按鈕，避免按到語言切換等其他按鈕
    btn = subs.filter(function(x) { return /submit|continue|go\b|ok\b/i.test(x.name + ' ' + x.id + ' ' + x.value); })[0]
       || subs[0] || null;
  }
  if (btn && btn.name) { fields[btn.name] = btn.value || 'Submit'; return; }
  var pb = _decode(html).match(/__doPostBack\(\s*['"]([^'"]+)['"]/);
  if (pb) { fields.__EVENTTARGET = pb[1]; fields.__EVENTARGUMENT = ''; }
}

function _textById(html, idPart) {
  var m = html.match(new RegExp('<(span|td|div|a)\\b[^>]*id=["\'][^"\']*' + idPart + '["\'][^>]*>([\\s\\S]*?)</\\1>', 'i'));
  return m ? _decode(m[2].replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim() : '';
}

function _isbn13to10(isbn) {
  if (!/^978\d{10}$/.test(isbn)) return null;
  var core = isbn.slice(3, 12), sum = 0;
  for (var i = 0; i < 9; i++) sum += (10 - i) * +core[i];
  var c = (11 - sum % 11) % 11;
  return core + (c === 10 ? 'X' : String(c));
}

function _isbn10to13(isbn) {
  if (!/^\d{9}[0-9X]$/.test(isbn)) return null;
  var core = '978' + isbn.slice(0, 9), sum = 0;
  for (var i = 0; i < 12; i++) sum += (i % 2 ? 3 : 1) * +core[i];
  return core + String((10 - sum % 10) % 10);
}

// Google Books（帶金鑰）：前端不帶金鑰的查詢額度用完時的備援。
// 金鑰存在指令碼屬性 GOOGLE_BOOKS_API_KEY；沒設定就略過。
function _lookupGoogleBooks(isbn) {
  var key = PropertiesService.getScriptProperties().getProperty('GOOGLE_BOOKS_API_KEY');
  if (!key) { _note('google books: 未設定 GOOGLE_BOOKS_API_KEY，略過'); return null; }

  var resp = _fetch('https://www.googleapis.com/books/v1/volumes?q=isbn:' + isbn + '&country=TW' +
                    '&key=' + encodeURIComponent(key), { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;
  var item = (JSON.parse(resp.getContentText()).items || [])[0];
  var info = item && item.volumeInfo;
  if (!info || !info.title) return null;

  var links = info.imageLinks || {};
  return {
    title:     String(info.title).trim(),
    authors:   (info.authors || []).join('、'),
    publisher: info.publisher || '',
    thumbnail: (links.thumbnail || links.smallThumbnail || '').replace(/^http:/, 'https:'),
    source:    'Google Books',
  };
}

// 國家圖書館「全國新書資訊網」：查詢會經過多次轉址並依賴 session cookie，
// 所以手動跟隨轉址並帶上 cookie。結果表格欄位以 td[data-th="書名"] 等標示。
var NCL_BASE = 'https://isbn.ncl.edu.tw/NEW_ISBNNet/';

function _lookupNcl(isbn) {
  var html = _fetchWithCookies(NCL_BASE + 'H30_SearchBooks.php?Pact=Search&Pval=' + encodeURIComponent(isbn));
  if (!html) return null;

  var count = html.match(/找到\s*(\d+)\s*筆/);
  if (count && +count[1] === 0) { _note('ncl: 找到 0 筆'); return null; }

  // 取結果表格中第一筆含「書名」欄位的資料列
  var tableStart = html.search(/class=["'][^"']*table-searchbooks/);
  if (tableStart < 0) { _note('ncl: 找不到結果表格 table-searchbooks'); return null; }
  var rows = html.slice(tableStart).split(/<tr[\s>]/i);
  var row  = null;
  for (var i = 0; i < rows.length; i++) {
    if (/data-th=["']書名["']/.test(rows[i])) { row = rows[i]; break; }
  }
  if (!row) { _note('ncl: 表格中沒有「書名」欄位'); return null; }

  // 國圖書名會在「=」後附上原文書名（例如「原始智慧: …= Primal intelligence: …」），只保留中文部分
  var title = _cell(row, '書名').replace(/\s*=\s*[^=]*$/, '').trim();
  if (!title) return null;

  var thumbnail = '';
  var coverCell = _cellHtml(row, '封面圖');
  var img = coverCell.match(/<img[^>]+src=["']([^"']+)["']/i);
  if (img && img[1].indexOf('no_cover') < 0) thumbnail = _absUrl(_decode(img[1]), NCL_BASE);

  return {
    title:     title,
    authors:   _cell(row, '作者'),
    publisher: _cell(row, '出版者'),
    thumbnail: thumbnail.replace(/^http:/, 'https:'),
    source:    '全國新書資訊網',
  };
}

function _cellHtml(row, label) {
  var m = row.match(new RegExp('<td[^>]*data-th=["\']' + label + '["\'][^>]*>([\\s\\S]*?)</td>', 'i'));
  return m ? m[1] : '';
}

function _cell(row, label) {
  var h = _cellHtml(row, label);
  var a = h.match(/<a[^>]*>([\s\S]*?)<\/a>/i);  // 書名通常包在連結裡
  var text = (a ? a[1] : h).replace(/<[^>]+>/g, ' ');
  return _decode(text).replace(/\s+/g, ' ').trim();
}

// 帶 cookie 的連線：手動跟隨轉址（最多 6 次），每一步都帶上累積的 cookie。
// get(url) / post(url, 欄位物件) 成功回傳 HTML 文字，失敗回傳 null；lastUrl 是最後到達的網址。
function _session(acceptLanguage) {
  var cookies = {};
  var s = { lastUrl: '' };

  function request(url, method, payload) {
    for (var hop = 0; hop < 6; hop++) {
      var header = Object.keys(cookies).map(function(k) { return k + '=' + cookies[k]; }).join('; ');
      var opts = {
        method: method,
        muteHttpExceptions: true,
        followRedirects: false,
        headers: {
          'User-Agent':      _FETCH_OPTS.headers['User-Agent'],
          'Accept-Language': acceptLanguage || 'zh-TW,zh;q=0.9',
        },
      };
      if (header) opts.headers.Cookie = header;
      if (payload) opts.payload = payload;   // 物件會以 application/x-www-form-urlencoded 送出
      if (payload && _trace) _trace.push({ postFields: _summarizeFields(payload) });

      var resp = _fetch(url, opts);
      var all  = resp.getAllHeaders();
      var set  = all['Set-Cookie'] || all['set-cookie'];
      if (set) {
        (Array.isArray(set) ? set : [set]).forEach(function(c) {
          var kv = String(c).split(';')[0];
          var eq = kv.indexOf('=');
          if (eq > 0) cookies[kv.slice(0, eq).trim()] = kv.slice(eq + 1).trim();
        });
      }

      s.lastUrl = url;
      var code = resp.getResponseCode();
      if (code === 200) return resp.getContentText('UTF-8');
      if (code >= 300 && code < 400) {
        var loc = all['Location'] || all['location'];
        if (!loc) return null;
        url = _absUrl(Array.isArray(loc) ? loc[0] : loc, url);
        method = 'get';      // 轉址後一律改用 GET
        payload = null;
        continue;
      }
      return null;
    }
    return null;
  }

  s.get  = function(url) { return request(url, 'get', null); };
  s.post = function(url, fields) { return request(url, 'post', fields); };
  return s;
}

function _fetchWithCookies(url) {
  return _session().get(url);
}

function _absUrl(href, base) {
  if (/^https?:\/\//i.test(href)) return href;
  if (href.indexOf('//') === 0) return 'https:' + href;
  var origin = base.match(/^https?:\/\/[^\/]+/)[0];
  if (href.charAt(0) === '/') return origin + href;
  var path = base.replace(/[?#].*$/, '');
  // 只有查詢參數（例如「?&Pact=…」）時，沿用原本的檔名，只換掉 ? 之後的部分
  if (href.charAt(0) === '?') return path + href;
  if (href.charAt(0) === '#') return base.replace(/#.*$/, '') + href;
  return path.replace(/[^\/]*$/, '') + href.replace(/^\.\//, '');
}

function _lookupBooksComTw(isbn) {
  var search = _fetch(
    'https://search.books.com.tw/search/query/key/' + isbn + '/cat/all', _FETCH_OPTS);
  if (search.getResponseCode() !== 200) return null;
  var html = search.getContentText('UTF-8');

  var m = html.match(/books\.com\.tw\/products\/([0-9A-Z]{10})/);
  if (!m) { _note('books.com.tw: 搜尋結果沒有商品'); return null; }
  var productId = m[1];

  var page = _fetch('https://www.books.com.tw/products/' + productId, _FETCH_OPTS);
  if (page.getResponseCode() !== 200) return null;
  var p = page.getContentText('UTF-8');

  var title = _meta(p, 'og:title') || _tag(p, 'h1');
  title = title.replace(/\s*[-|｜]\s*博客來.*$/, '').trim();
  if (!title) return null;

  // description 格式例如：「書名：XXX，語言：繁體中文，ISBN：…，出版社：YYY，作者：ZZZ，…」
  var desc      = _meta(p, 'description');
  var authors   = _field(desc, '作者') || _authorLink(p);
  var publisher = _field(desc, '出版社');
  var thumbnail = _meta(p, 'og:image').replace(/^http:/, 'https:');

  return {
    title:     title,
    authors:   authors,
    publisher: publisher,
    thumbnail: thumbnail,
    source:    '博客來',
  };
}

function _meta(html, name) {
  var re = new RegExp('<meta[^>]+(?:property|name)=["\']' + name + '["\'][^>]*content=["\']([^"\']*)', 'i');
  var m  = html.match(re);
  if (!m) {
    re = new RegExp('<meta[^>]+content=["\']([^"\']*)["\'][^>]*(?:property|name)=["\']' + name + '["\']', 'i');
    m  = html.match(re);
  }
  return m ? _decode(m[1]).trim() : '';
}

function _tag(html, tag) {
  var m = html.match(new RegExp('<' + tag + '[^>]*>([\\s\\S]*?)</' + tag + '>', 'i'));
  return m ? _decode(m[1].replace(/<[^>]+>/g, '')).trim() : '';
}

function _field(desc, label) {
  var m = desc.match(new RegExp(label + '[：:]\\s*([^，,]+)'));
  return m ? m[1].trim() : '';
}

function _authorLink(html) {
  var m = html.match(/作者[：:][\s\S]{0,200}?<a[^>]*>([^<]+)<\/a>/);
  return m ? _decode(m[1]).trim() : '';
}

function _decode(s) {
  return String(s)
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, function(_, n) { return String.fromCharCode(+n); });
}

// One-time helper: run manually in Apps Script editor to clear all data rows
function clearAllRecords() {
  var sheet = getSheet();
  var lr = sheet.getLastRow();
  if (lr > 1) sheet.deleteRows(2, lr - 1);
  Logger.log('Cleared ' + (lr - 1) + ' records.');
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
