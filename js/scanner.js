// Camera barcode scanner using html5-qrcode (Stage 4)
// Supports EAN-13, EAN-8, UPC-A on iOS Safari ≥15.1 and Android

let _scanner = null;

async function startScanner(onDetected, onError) {
  const container = document.getElementById('scanner-container');
  if (!container) { onError(new Error('No scanner container')); return; }
  container.classList.remove('hidden');

  if (typeof Html5Qrcode === 'undefined') {
    container.classList.add('hidden');
    onError(new Error('Scanner library not available — check network'));
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
