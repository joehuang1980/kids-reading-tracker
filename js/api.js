// Book lookup (Google Books / Open Library / 全國新書資訊網 / 博客來) + Google Apps Script integration (Stage 7–8, 13–14)

// Stage 13–14: fetch book metadata including cover, author, publisher
// 依序查詢多個來源，中文書在 Google Books 常查不到，因此加上：
//   1. Google Books
//   2. Open Library（部分中文書）
//   3. 國家圖書館全國新書資訊網 → 博客來（透過 Apps Script 代查，避開瀏覽器跨網域限制）
// 任一來源找到書名就回傳；全部失敗且都是網路錯誤時拋出錯誤，讓畫面顯示「無法連線」。
async function fetchBookByIsbn(isbn) {
  const sources = [fetchFromGoogleBooks, fetchFromOpenLibrary, fetchFromTaiwanSources];
  let errors = 0;
  for (const source of sources) {
    try {
      const book = await source(isbn);
      if (book?.title) return book;
    } catch (err) {
      errors++;
      console.warn(`[book lookup] ${source.name} failed:`, err);
    }
  }
  if (errors === sources.length) throw new Error('All book sources failed');
  return null;
}

async function fetchJson(url, timeoutMs = 8000) {
  const ctrl  = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const resp = await fetch(url, { signal: ctrl.signal });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    return await resp.json();
  } finally {
    clearTimeout(timer);
  }
}

async function fetchFromGoogleBooks(isbn) {
  const base = `https://www.googleapis.com/books/v1/volumes?q=isbn:${encodeURIComponent(isbn)}`;
  let data;
  try {
    data = await fetchJson(CONFIG.GOOGLE_BOOKS_API_KEY
      ? `${base}&key=${encodeURIComponent(CONFIG.GOOGLE_BOOKS_API_KEY)}` : base);
  } catch (err) {
    // 金鑰額度用完（429）或金鑰被拒（403）時，改用不帶金鑰的方式再查一次
    if (!CONFIG.GOOGLE_BOOKS_API_KEY || !/HTTP (429|403)/.test(err.message)) throw err;
    data = await fetchJson(base);
  }
  const info = data.items?.[0]?.volumeInfo;
  if (!info) return null;

  return {
    title:     info.title?.trim() ?? '',
    authors:   (info.authors ?? []).join('、'),
    publisher: info.publisher ?? '',
    thumbnail: (info.imageLinks?.thumbnail ?? info.imageLinks?.smallThumbnail ?? '')
               .replace(/^http:/, 'https:'),
    source:    'Google Books',
  };
}

async function fetchFromOpenLibrary(isbn) {
  const keyName = `ISBN:${isbn}`;
  const data = await fetchJson(
    `https://openlibrary.org/api/books?bibkeys=${encodeURIComponent(keyName)}&format=json&jscmd=data`);
  const info = data[keyName];
  if (!info) return null;

  return {
    title:     [info.title, info.subtitle].filter(Boolean).join('：').trim(),
    authors:   (info.authors ?? []).map(a => a.name).filter(Boolean).join('、'),
    publisher: (info.publishers ?? []).map(p => p.name).filter(Boolean).join('、'),
    thumbnail: (info.cover?.medium ?? info.cover?.small ?? '').replace(/^http:/, 'https:'),
    source:    'Open Library',
  };
}

// 全國新書資訊網 / 博客來：需在 Apps Script 部署含 lookupBook() 的新版本
async function fetchFromTaiwanSources(isbn) {
  if (!CONFIG.APPS_SCRIPT_URL) return null;
  const data = await fetchJson(
    `${CONFIG.APPS_SCRIPT_URL}?action=lookup&isbn=${encodeURIComponent(isbn)}`, 15000);
  const b = data.book;
  if (!b?.title) return null;

  return {
    title:     String(b.title).trim(),
    authors:   b.authors ?? '',
    publisher: b.publisher ?? '',
    thumbnail: (b.thumbnail ?? '').replace(/^http:/, 'https:'),
    source:    b.source ?? '台灣書目',
  };
}

// Stage 8: POST to Apps Script — no Content-Type header avoids CORS preflight
async function submitRecord(record) {
  if (!CONFIG.APPS_SCRIPT_URL) {
    throw new Error('尚未設定 APPS_SCRIPT_URL（請參閱 apps-script/Code.gs 說明）');
  }
  const resp = await fetch(CONFIG.APPS_SCRIPT_URL, {
    method: 'POST',
    body: JSON.stringify(record),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  return resp.json();
}

// Stage 11–12: GET all records from Apps Script
async function fetchRecords() {
  if (!CONFIG.APPS_SCRIPT_URL) return [];
  const resp = await fetch(`${CONFIG.APPS_SCRIPT_URL}?action=list`);
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const data = await resp.json();
  return data.records ?? [];
}
