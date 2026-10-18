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
  if (params.action === 'lookup') return _json(lookupBook(params.isbn));
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
function lookupBook(isbn) {
  isbn = String(isbn || '').replace(/[^0-9Xx]/g, '').toUpperCase();
  if (!/^(\d{13}|\d{9}[0-9X])$/.test(isbn)) return { book: null, message: 'invalid isbn' };

  var cache = CacheService.getScriptCache();
  var hit   = cache.get('isbn:' + isbn);
  if (hit) return { book: JSON.parse(hit) };

  var book = null;
  try { book = _lookupNcl(isbn); } catch (err) { Logger.log('ncl: ' + err); }
  if (!book || !book.title) {
    try { book = _lookupBooksComTw(isbn); } catch (err) { Logger.log('books.com.tw: ' + err); }
  }

  if (book && book.title) {
    cache.put('isbn:' + isbn, JSON.stringify(book), 21600);  // 6 小時
    return { book: book };
  }
  return { book: null };
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
  if (count && +count[1] === 0) return null;

  // 取結果表格中第一筆含「書名」欄位的資料列
  var tableStart = html.search(/class=["'][^"']*table-searchbooks/);
  if (tableStart < 0) return null;
  var rows = html.slice(tableStart).split(/<tr[\s>]/i);
  var row  = null;
  for (var i = 0; i < rows.length; i++) {
    if (/data-th=["']書名["']/.test(rows[i])) { row = rows[i]; break; }
  }
  if (!row) return null;

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

    var resp = UrlFetchApp.fetch(url, opts);
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
  var search = UrlFetchApp.fetch(
    'https://search.books.com.tw/search/query/key/' + isbn + '/cat/all', _FETCH_OPTS);
  if (search.getResponseCode() !== 200) return null;
  var html = search.getContentText('UTF-8');

  var m = html.match(/books\.com\.tw\/products\/([0-9A-Z]{10})/);
  if (!m) return null;
  var productId = m[1];

  var page = UrlFetchApp.fetch('https://www.books.com.tw/products/' + productId, _FETCH_OPTS);
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
