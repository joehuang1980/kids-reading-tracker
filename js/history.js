// Reading history page logic (Stage 11)

const $ = id => document.getElementById(id);
const MONTHS_TW = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];

let _allRecords = [];

// Load offline queue records to show pending items too
function _pendingRecords() {
  try {
    return JSON.parse(localStorage.getItem(CONFIG.OFFLINE_QUEUE_KEY) ?? '[]')
      .map(r => ({ ...r, _pending: true }));
  } catch { return []; }
}

function escHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

async function loadHistory() {
  $('list').innerHTML = '<div class="empty"><div class="empty-icon">⏳</div>載入中…</div>';

  const pending = _pendingRecords();

  if (!CONFIG.APPS_SCRIPT_URL) {
    _allRecords = pending;
    _populateFilters();
    renderList();
    if (pending.length === 0) {
      $('list').innerHTML = '<div class="empty"><div class="empty-icon">⚙️</div>尚未設定後台連線（APPS_SCRIPT_URL），請參閱 apps-script/Code.gs 說明。</div>';
    }
    return;
  }

  try {
    const synced = await fetchRecords();
    synced.sort((a, b) => (b.readingDate ?? '').localeCompare(a.readingDate ?? ''));
    // Merge: pending first, then synced
    _allRecords = [...pending, ...synced];
    _populateFilters();
    renderList();
  } catch (err) {
    _allRecords = pending;
    _populateFilters();
    renderList();
    $('errorMsg').textContent = `⚠️ 無法載入同步紀錄：${err.message}`;
    $('errorMsg').classList.remove('hidden');
  }
}

function _populateFilters() {
  const readers = [...new Set(_allRecords.map(r => r.reader).filter(Boolean))].sort();
  $('filterReader').innerHTML = '<option value="">全部閱讀者</option>' +
    readers.map(r => `<option value="${r}">${escHtml(r)}</option>`).join('');

  const years = [...new Set(_allRecords.map(r => r.readingDate?.slice(0, 4)).filter(Boolean))].sort().reverse();
  $('filterYear').innerHTML = '<option value="">全部年份</option>' +
    years.map(y => `<option value="${y}">${y} 年</option>`).join('');

  $('filterMonth').innerHTML = '<option value="">全部月份</option>' +
    MONTHS_TW.map((m, i) => `<option value="${String(i + 1).padStart(2, '0')}">${m}</option>`).join('');
}

function renderList() {
  const reader = $('filterReader').value;
  const year   = $('filterYear').value;
  const month  = $('filterMonth').value;

  const filtered = _allRecords.filter(r => {
    if (reader && r.reader !== reader) return false;
    if (year   && r.readingDate?.slice(0, 4) !== year) return false;
    if (month  && r.readingDate?.slice(5, 7) !== month) return false;
    return true;
  });

  $('summary').textContent = `共 ${filtered.length} 筆`;

  if (!filtered.length) {
    $('list').innerHTML = '<div class="empty"><div class="empty-icon">📭</div>沒有符合的紀錄</div>';
    return;
  }

  $('list').innerHTML = filtered.map(r => `
    <div class="record-card${r._pending ? ' pending' : ''}">
      ${r.thumbnail
        ? `<img class="record-cover" src="${escHtml(r.thumbnail)}" alt="" loading="lazy">`
        : `<div class="record-cover-ph">📖</div>`}
      <div class="record-info">
        <div class="record-title">${escHtml(r.title)}</div>
        <div class="record-meta">
          ${r.reader ? `👤 ${escHtml(r.reader)}　` : ''}📅 ${escHtml(r.readingDate ?? '')}${r.authors ? `<br>${escHtml(r.authors)}` : ''}${r.isbn ? `<br><span style="font-size:12px">ISBN: ${escHtml(r.isbn)}</span>` : ''}
        </div>
        ${r._pending ? '<div class="record-pending-badge">⏳ 等待同步</div>' : ''}
      </div>
    </div>
  `).join('');
}

$('filterReader').addEventListener('change', renderList);
$('filterYear').addEventListener('change', renderList);
$('filterMonth').addEventListener('change', renderList);

loadHistory();
