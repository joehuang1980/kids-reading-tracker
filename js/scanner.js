// Camera barcode scanner using html5-qrcode (Stage 4)
// Supports EAN-13, EAN-8, UPC-A on iOS Safari ≥15.1 and Android

let _scanner = null;

async function startScanner(onDetected, onError) {
  const container = document.getElementById('scanner-container');
  if (!container) { onError(new Error('No scanner container')); return; }
  container.classList.remove('hidden');

  if (typeof Html5Qrcode === 'undefined') {
    container.classList.add('hidden');
    onError(new Error('SCANNER_LIB_MISSING'));
    return;
  }
  if (!window.isSecureContext || !navigator.mediaDevices?.getUserMedia) {
    container.classList.add('hidden');
    onError(new Error('CAMERA_UNSUPPORTED'));
    return;
  }

  await _destroyScanner();

  _scanner = new Html5Qrcode('reader');
  try {
    await _scanner.start(
      { facingMode: 'environment' },
      {
        fps: 10,
        qrbox: { width: 260, height: 130 },
        formatsToSupport: [
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.UPC_A,
        ],
      },
      (text) => {
        stopScanner();
        onDetected(text.replace(/[^0-9X]/gi, ''));
      },
    );
  } catch (err) {
    await _destroyScanner();
    container.classList.add('hidden');
    onError(err);
  }
}

// 把相機錯誤轉成看得懂的中文說明
function describeScannerError(err) {
  const msg = String(err?.name || '') + ' ' + String(err?.message || err || '');
  if (/SCANNER_LIB_MISSING/.test(msg))       return '掃描元件沒有載入，請重新整理頁面。';
  if (/CAMERA_UNSUPPORTED/.test(msg))        return '這個瀏覽器不支援相機（需使用 https 網址開啟）。';
  if (/NotAllowed|Permission/i.test(msg))    return '相機權限被拒絕。請到「設定 → Safari → 相機」改成「允許」，或在網址列左側的「ㄚA」選單中允許相機，再重新整理。';
  if (/NotFound|Overconstrained/i.test(msg)) return '找不到可用的相機。';
  if (/NotReadable|TrackStart|in use/i.test(msg)) return '相機正被其他 App 使用中，請關閉其他使用相機的 App 後再試。';
  return `無法開啟相機（${msg.trim() || '未知錯誤'}）。`;
}

async function stopScanner() {
  await _destroyScanner();
  const container = document.getElementById('scanner-container');
  if (container) container.classList.add('hidden');
}

async function _destroyScanner() {
  if (!_scanner) return;
  try { await _scanner.stop(); } catch {}
  try { await _scanner.clear(); } catch {}
  _scanner = null;
}
