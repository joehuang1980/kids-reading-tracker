// Dashboard statistics logic (Stage 12)

const $ = id => document.getElementById(id);
const MONTHS_TW = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];

function escHtml(s) {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

async function loadDashboard() {
  if (!CONFIG.APPS_SCRIPT_URL) {
    const msg = '<div class="empty"><div class="empty-icon">⚙️</div>尚未設定後台連線，請先完成 Stage 6–7 設定。</div>';
    $('readerRows').innerHTML = msg;
    $('barChart').innerHTML = '';
    return;
  }

  try {
    const records = await fetchRecords();
    renderStats(records);
  } catch (err) {
    $('readerRows').innerHTML = `<div class="empty">載入失敗：${escHtml(err.message)}</div>`;
  }
}

function renderStats(records) {
  const now       = new Date();
  const thisYear  = String(now.getFullYear());
  const thisMonth = `${thisYear}-${String(now.getMonth() + 1).padStart(2, '0')}`;

  // Stat cards
  $('statMonth').textContent = records.filter(r => r.readingDate?.startsWith(thisMonth)).length;
  $('statYear').textContent  = records.filter(r => r.readingDate?.startsWith(thisYear)).length;
  $('statTotal').textContent = records.length;

  // Per reader
  const byReader = {};
  for (const r of records) {
    if (r.reader) byReader[r.reader] = (byReader[r.reader] ?? 0) + 1;
  }
  const sortedReaders = Object.entries(byReader).sort((a, b) => b[1] - a[1]);
  $('readerRows').innerHTML = sortedReaders.length
    ? sortedReaders.map(([name, count]) => `
        <div class="reader-row">
          <span>${escHtml(name)}</span>
          <strong>${count} 本</strong>
        </div>
      `).join('')
    : '<div class="empty" style="padding:20px">尚無紀錄</div>';

  // Monthly bar chart — last 12 months
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const d   = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    months.push({ label: MONTHS_TW[d.getMonth()], count: records.filter(r => r.readingDate?.startsWith(key)).length });
  }
  const maxCount = Math.max(...months.map(m => m.count), 1);
  $('barChart').innerHTML = months.map(m => `
    <div class="bar-row">
      <span>${m.label}</span>
      <div class="bar-track">
        <div class="bar-fill" style="width:${(m.count / maxCount * 100).toFixed(1)}%"></div>
      </div>
      <span class="bar-count">${m.count}</span>
    </div>
  `).join('');
}

loadDashboard();
