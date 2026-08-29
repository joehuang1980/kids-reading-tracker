// Google Books API + Google Apps Script integration (Stage 7–8, 13–14)

// Stage 13–14: fetch book metadata including cover, author, publisher
async function fetchBookByIsbn(isbn) {
  const key = CONFIG.GOOGLE_BOOKS_API_KEY
    ? `&key=${encodeURIComponent(CONFIG.GOOGLE_BOOKS_API_KEY)}`
    : '';
  const url = `https://www.googleapis.com/books/v1/volumes?q=isbn:${encodeURIComponent(isbn)}${key}`;

  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

  const data = await resp.json();
  const info = data.items?.[0]?.volumeInfo;
  if (!info) return null;

  return {
    title:     info.title?.trim() ?? '',
    authors:   (info.authors ?? []).join('、'),
    publisher: info.publisher ?? '',
    thumbnail: (info.imageLinks?.thumbnail ?? info.imageLinks?.smallThumbnail ?? '')
               .replace(/^http:/, 'https:'),
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
