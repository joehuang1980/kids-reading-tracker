// Main form logic: ISBN lookup, scanner, submit, offline queue (Stage 3–9, 16)

// ─── Offline queue (Stage 16) ──────────────────────────────────────────────
function _qGet() {
  try { return JSON.parse(localStorage.getItem(CONFIG.OFFLINE_QUEUE_KEY) ?? '[]'); }
  catch { return []; }
}
function _qSave(q) { localStorage.setItem(CONFIG.OFFLINE_QUEUE_KEY, JSON.stringify(q)); }
function _qAdd(record) { const q = _qGet(); q.push(record); _qSave(q); _badgeUpdate(); }

async function _qFlush() {
  if (!CONFIG.APPS_SCRIPT_URL) return;
  const q = _qGet();
  if (!q.length) return;
  const failed = [];
  for (const r of q) {
    try { await submitRecord(r); }
    catch { failed.push(r); }
  }
  _qSave(failed);
  _badgeUpdate();
}

function _badgeUpdate() {
  const badge = document.getElementById('syncBadge');
  if (!badge) return;
  const n = _qGet().length;
  if (n > 0) {
    badge.textContent = `⏳ ${n} 筆紀錄等待同步，點擊重試`;
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
}

window.addEventListener('online', _qFlush);

// ─── State ─────────────────────────────────────────────────────────────────
let _book     = null;  // { title, authors, publisher, thumbnail }
let _scanning = false;

// ─── UI helpers ────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

function showStatus(msg, type = '') {
  const el = $('status');
  el.textContent = msg;
  el.className = `status ${type}`.trim();
  el.classList.remove('hidden');
}
function hideStatus() { $('status').classList.add('hidden'); }

function resetBookResult() {
  _book = null;
  $('result').classList.add('hidden');
  $('manualBox').classList.add('hidden');
  $('formBox').classList.add('hidden');
  $('bookTitle').textContent = '';
  $('bookMeta').textContent  = '';
  const img = $('bookCover');
  img.src = '';
  img.classList.add('hidden');
  $('coverPh').classList.remove('hidden');
  $('manualTitle').value = '';
}

// ─── Book lookup (Stage 3, 13–14) ──────────────────────────────────────────
async function lookupBook() {
  resetBookResult();
  hideStatus();

  const isbn = normalizeIsbn($('isbn').value);
  $('isbn').value = isbn;

  if (!isValidIsbnShape(isbn)) {
    showStatus('請輸入 10 碼或 13 碼的有效 ISBN 格式。', 'error');
    $('manualBox').classList.remove('hidden');
    $('formBox').classList.remove('hidden');
    return;
  }

  $('lookupBtn').disabled = true;
  $('lookupBtn').textContent = '查詢中…';
  showStatus('正在查詢書籍資料…');

  try {
    const book = await fetchBookByIsbn(isbn);
    if (book?.title) {
      _book = book;
      $('bookTitle').textContent = book.title;
      $('bookMeta').textContent  = [book.authors, book.publisher].filter(Boolean).join('　·　');
      if (book.thumbnail) {
        const img = $('bookCover');
        img.src = book.thumbnail;
        img.classList.remove('hidden');
        $('coverPh').classList.add('hidden');
      }
      $('result').classList.remove('hidden');
      showStatus('已找到書籍資料。', 'success');
    } else {
      showStatus('查不到這本書，請手動輸入書名。', 'error');
      $('manualBox').classList.remove('hidden');
    }
  } catch {
    showStatus('目前無法連線查詢，請手動輸入書名。', 'error');
    $('manualBox').classList.remove('hidden');
  } finally {
    $('lookupBtn').disabled = false;
    $('lookupBtn').textContent = '查詢書籍';
    $('formBox').classList.remove('hidden');
  }
}

// ─── Submit record (Stage 5, 8, 9, 16) ────────────────────────────────────
async function handleSubmit() {
  const isbn   = normalizeIsbn($('isbn').value);
  const title  = (_book?.title || $('manualTitle').value).trim();
  const date   = $('readingDate').value;
  const reader = $('readerSelect').value;

  // Stage 5: Validation
  if (!title)  { showStatus('請輸入或確認書名。', 'error'); return; }
  if (!date)   { showStatus('請選擇閱讀日期。', 'error'); return; }
  if (!reader) { showStatus('請選擇閱讀者。', 'error'); return; }

  const record = {
    isbn,
    title,
    readingDate: date,
    reader,
    authors:   _book?.authors   ?? '',
    publisher: _book?.publisher ?? '',
    thumbnail: _book?.thumbnail ?? '',
    submittedAt: new Date().toISOString(),
  };

  // Stage 9: Disable button immediately to prevent duplicate submit
  $('submitBtn').disabled = true;
  showStatus('正在儲存閱讀紀錄…');

  try {
    if (!navigator.onLine) throw new Error('offline');

    const result = await submitRecord(record);

    // Stage 9: Backend duplicate detection
    if (result?.duplicate) {
      showStatus('⚠️ 偵測到重複紀錄，請確認是否真的需要再次儲存。', 'error');
      $('submitBtn').disabled = false;
      return;
    }

    showStatus(`✅ 閱讀紀錄已儲存！（${title}・${reader}）`, 'success');
    _afterSubmit(reader, date);
  } catch (err) {
    if (err.message === 'offline' || !navigator.onLine) {
      // Stage 16: Offline queue
      _qAdd(record);
      showStatus('📵 目前無網路，紀錄已暫存，恢復連線後自動同步。', 'offline');
      _afterSubmit(reader, date);
    } else {
      console.error(err);
      showStatus(`儲存失敗，請稍後再試。（${err.message}）`, 'error');
      $('submitBtn').disabled = false;
    }
  }
}

// Stage 9: Keep reader & date; clear book fields for next entry
function _afterSubmit(reader, date) {
  $('isbn').value = '';
  resetBookResult();
  $('readerSelect').value = reader;
  $('readingDate').value = date;
  $('submitBtn').disabled = false;
}

// ─── Camera scanner (Stage 4) ──────────────────────────────────────────────
async function handleScan() {
  if (_scanning) {
    await stopScanner();
    _scanning = false;
    $('scanBtn').textContent = '📷';
    return;
  }
  _scanning = true;
  $('scanBtn').textContent = '⏹';
  await startScanner(
    (isbn) => {
      _scanning = false;
      $('scanBtn').textContent = '📷';
      $('isbn').value = isbn;
      lookupBook();
    },
    (err) => {
      _scanning = false;
      $('scanBtn').textContent = '📷';
      console.warn('Scanner error:', err.message);
      showStatus('無法開啟相機，請手動輸入 ISBN。', 'error');
      $('manualBox').classList.remove('hidden');
      $('formBox').classList.remove('hidden');
    },
  );
}

// ─── Init ───────────────────────────────────────────────────────────────────
function init() {
  $('readingDate').value = new Date().toISOString().slice(0, 10);
  $('readerSelect').innerHTML  = CONFIG.READERS
    .map(r => `<option value="${r}">${r}</option>`).join('');

  _badgeUpdate();
  if (navigator.onLine) _qFlush();

  $('lookupBtn').addEventListener('click', lookupBook);
  $('isbn').addEventListener('keydown', e => { if (e.key === 'Enter') lookupBook(); });
  $('scanBtn').addEventListener('click', handleScan);
  $('submitBtn').addEventListener('click', handleSubmit);
  $('syncBadge').addEventListener('click', _qFlush);
  $('syncBadge').addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); _qFlush(); }
  });
}

init();
