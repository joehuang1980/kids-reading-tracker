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
let _book     = null;  // { title, authors, publisher, thumbnail, language, source }
let _cover    = null;  // 自行拍攝的書封（JPEG data URL）
let _ar       = null;  // 英文書 AR 字數查詢：{ isbn, promise, result }
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
  img.onerror = null;               // 清空圖片也會觸發 error，先取消處理
  img.removeAttribute('src');
  img.classList.add('hidden');
  $('coverPh').classList.remove('hidden');
  $('manualTitle').value = '';

  _cover = null;
  _ar    = null;
  $('coverCaptureBox').classList.add('hidden');
  $('coverPreview').src = '';
  $('coverPreview').classList.add('hidden');
  $('coverBtn').textContent = '📷 拍攝書封';
  $('coverHint').textContent = '查不到封面，可以自己拍一張書封留作紀錄。';
  $('coverInput').value = '';
  $('arBox').classList.add('hidden');
  $('arBox').textContent = '';
}

// ─── 自拍書封 ──────────────────────────────────────────────────────────────
function _offerCoverCapture() {
  $('coverCaptureBox').classList.remove('hidden');
}

function _showCover(src) {
  const img = $('bookCover');
  img.onerror = () => {             // 封面網址失效時，改讓使用者自己拍
    if (img.getAttribute('src') !== src) return;
    img.classList.add('hidden');
    $('coverPh').classList.remove('hidden');
    if (_book) _book.thumbnail = '';
    _offerCoverCapture();
  };
  img.src = src;
  img.classList.remove('hidden');
  $('coverPh').classList.add('hidden');
}

async function handleCoverFile(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  showStatus('正在處理書封照片…');
  try {
    _cover = await makeCoverThumbnail(file);
    if (!$('result').classList.contains('hidden')) {
      _showCover(_cover);             // 書籍資料區已有封面位置，就顯示在那裡
    } else {
      $('coverPreview').src = _cover;
      $('coverPreview').classList.remove('hidden');
    }
    $('coverHint').textContent = '已加入你拍的書封，不滿意可以重拍。';
    $('coverBtn').textContent = '📷 重拍書封';
    showStatus('已加入書封照片。', 'success');
  } catch (err) {
    console.warn('cover photo error:', err);
    showStatus('照片處理失敗，請再拍一次。', 'error');
  } finally {
    $('coverInput').value = '';
  }
}

// ─── 英文書 AR 字數 ────────────────────────────────────────────────────────
function _startArLookup(isbn) {
  const box = $('arBox');
  box.textContent = '📏 正在查詢 AR BookFind 字數…';
  box.classList.remove('hidden');

  const state = { isbn, result: undefined };
  state.promise = fetchArInfo(isbn)
    .catch(err => { console.warn('[AR] lookup failed:', err); return null; })
    .then(ar => {
      state.result = ar;
      if (_ar === state) _renderAr(ar);
      return ar;
    });
  _ar = state;
}

function _renderAr(ar) {
  const box = $('arBox');
  if (!ar) {
    box.textContent = '📏 AR BookFind 查不到這本書，字數不會記錄。';
    return;
  }
  // 其他來源都查不到書名時，用 AR 的書名
  if (!_book && ar.title && !$('manualTitle').value.trim()) $('manualTitle').value = ar.title;

  const parts = [];
  parts.push(ar.wordCount ? `📏 AR 字數：${ar.wordCount.toLocaleString()} 字` : '📏 AR BookFind 沒有這本書的字數');
  if (ar.atosLevel) parts.push(`ATOS ${ar.atosLevel}`);
  if (ar.arPoints)  parts.push(`${ar.arPoints} 點`);
  box.textContent = parts.join('　·　');
}

// 送出時取得字數：查詢還沒好就最多再等 10 秒
async function _wordCountFor(isbn) {
  if (!_ar || _ar.isbn !== isbn) return null;
  if (_ar.result !== undefined) return _ar.result?.wordCount ?? null;
  showStatus('正在等待 AR 字數查詢…');
  const ar = await Promise.race([
    _ar.promise,
    new Promise(resolve => setTimeout(() => resolve(null), 10000)),
  ]);
  return ar?.wordCount ?? null;
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
      if (book.thumbnail) _showCover(book.thumbnail);
      else _offerCoverCapture();
      $('result').classList.remove('hidden');
      showStatus(`已找到書籍資料（來源：${book.source ?? '—'}）。`, 'success');
      if (isEnglishBook(isbn, book)) _startArLookup(isbn);
    } else {
      showStatus('查不到這本書，請手動輸入書名。', 'error');
      $('manualBox').classList.remove('hidden');
      _offerCoverCapture();
      if (isEnglishIsbn(isbn)) _startArLookup(isbn);   // AR 也許有這本書的書名與字數
    }
  } catch {
    showStatus('目前無法連線查詢，請手動輸入書名。', 'error');
    $('manualBox').classList.remove('hidden');
    _offerCoverCapture();
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

  $('submitBtn').disabled = true;
  const wordCount = await _wordCountFor(isbn);

  const record = {
    isbn,
    title,
    readingDate: date,
    reader,
    authors:   _book?.authors   ?? '',
    publisher: _book?.publisher ?? '',
    thumbnail: _cover || _book?.thumbnail || '',
    wordCount,
    submittedAt: new Date().toISOString(),
  };

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
      console.warn('Scanner error:', err);
      showStatus(`${describeScannerError(err)}也可以手動輸入 ISBN。`, 'error');
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
  $('coverBtn').addEventListener('click', () => $('coverInput').click());
  $('coverInput').addEventListener('change', handleCoverFile);
  $('syncBadge').addEventListener('click', _qFlush);
  $('syncBadge').addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); _qFlush(); }
  });
}

init();
