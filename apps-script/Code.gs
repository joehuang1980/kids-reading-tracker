/**
 * Kids Reading Tracker — Google Apps Script Web App
 * Stage 6: 在 Google Sheet 建立名為「Kids Reading Records」的試算表
 *          新增工作表「Reading_Log」，欄位如下：
 *          A: Record_ID  B: ISBN  C: Book_Title  D: Reading_Date
 *          E: Reader     F: Created_At  G: Thumbnail
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
 *
 *  ※ 更新程式碼後須「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」，
 *    前端才會用到新功能（例如中文書查詢 action=lookup）。
 *
 * Stage 10: Excel 匯出 — 在 Google Sheet 選「檔案」→「下載」→「Microsoft Excel (.xlsx)」
 */

var SHEET_NAME = 'Reading_Log';

function getSheet() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Record_ID', 'ISBN', 'Book_Title', 'Reading_Date', 'Reader', 'Created_At', 'Thumbnail']);
    sheet.setFrozenRows(1);
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

    sheet.appendRow([
      recordId,
      data.isbn      || '',
      data.title,
      data.readingDate,
      data.reader,
      new Date(),
      data.thumbnail || '',
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
  try {
    var sheet = getSheet();
    var lr    = sheet.getLastRow();
    if (lr < 2) return _json({ records: [] });

    var rows    = sheet.getRange(2, 1, lr - 1, 7).getValues();
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
var LOOKUP_VERSION = 3;
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
      url:      url,
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

  var title = _cell(row, '書名');
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

// 手動跟隨轉址（最多 5 次），每一步都帶上累積的 cookie
function _fetchWithCookies(url) {
  var cookies = {};
  for (var hop = 0; hop < 5; hop++) {
    var header = Object.keys(cookies).map(function(k) { return k + '=' + cookies[k]; }).join('; ');
    var opts = {
      muteHttpExceptions: true,
      followRedirects: false,
      headers: {
        'User-Agent':      _FETCH_OPTS.headers['User-Agent'],
        'Accept-Language': 'zh-TW,zh;q=0.9',
      },
    };
    if (header) opts.headers.Cookie = header;

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

    var code = resp.getResponseCode();
    if (code === 200) return resp.getContentText('UTF-8');
    if (code >= 300 && code < 400) {
      var loc = all['Location'] || all['location'];
      if (!loc) return null;
      url = _absUrl(Array.isArray(loc) ? loc[0] : loc, url);
      continue;
    }
    return null;
  }
  return null;
}

function _absUrl(href, base) {
  if (/^https?:\/\//i.test(href)) return href;
  if (href.indexOf('//') === 0) return 'https:' + href;
  var origin = base.match(/^https?:\/\/[^\/]+/)[0];
  if (href.charAt(0) === '/') return origin + href;
  return base.replace(/[?#].*$/, '').replace(/[^\/]*$/, '') + href;
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
