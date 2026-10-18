// Camera barcode scanner (Stage 4)
// 使用 zxing-wasm（ZXing C++ 的 WebAssembly 版）辨識 EAN-13 書籍條碼，
// 對書背上細小、略模糊的一維條碼比 JavaScript 版的 ZXing 準確很多。
// 相機要求較高解析度並連續對焦，只接受 978/979 開頭的 ISBN 條碼
// （忽略書上的店內碼、價格碼等其他條碼）。

const SCAN_INTERVAL_MS = 120;
const SCAN_MAX_WIDTH   = 1280;   // 送去辨識的畫面最大寬度（太大會變慢）

let _stream     = null;
let _video      = null;
let _timer      = null;
let _canvas     = null;
let _zxingReady = false;

function _prepareZXing() {
  if (_zxingReady) return;
  const base = new URL('js/vendor/', location.href).href;
  ZXingWASM.prepareZXingModule({
    overrides: {
      locateFile: (path, prefix) => path.endsWith('.wasm') ? base + path : prefix + path,
    },
  });
  _zxingReady = true;
}

async function startScanner(onDetected, onError) {
  const container = document.getElementById('scanner-container');
  const reader    = document.getElementById('reader');
  if (!container || !reader) { onError(new Error('No scanner container')); return; }

  if (typeof ZXingWASM === 'undefined') { onError(new Error('SCANNER_LIB_MISSING')); return; }
  if (!window.isSecureContext || !navigator.mediaDevices?.getUserMedia) {
    onError(new Error('CAMERA_UNSUPPORTED'));
    return;
  }

  await _destroyScanner();
  container.classList.remove('hidden');

  try {
    _prepareZXing();
    _stream = await navigator.mediaDevices.getUserMedia({
      audio: false,
      video: {
        facingMode: { ideal: 'environment' },
        width:  { ideal: 1920 },
        height: { ideal: 1080 },
      },
    });

    // 盡量開啟連續自動對焦（不支援的裝置會忽略）
    const track = _stream.getVideoTracks()[0];
    try { await track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }); } catch {}

    _video = document.createElement('video');
    _video.setAttribute('playsinline', '');   // iOS 必須，否則會跳全螢幕
    _video.muted = true;
    _video.autoplay = true;
    _video.srcObject = _stream;
    reader.innerHTML = '<div class="scan-guide" aria-hidden="true"></div>';
    reader.prepend(_video);
    await _video.play();
  } catch (err) {
    await _destroyScanner();
    container.classList.add('hidden');
    onError(err);
    return;
  }

  let busy = false;
  _timer = setInterval(async () => {
    if (busy || !_video || _video.readyState < 2) return;
    busy = true;
    try {
      const isbn = await _scanFrame();
      if (isbn && _timer) {
        if (navigator.vibrate) navigator.vibrate(80);
        await stopScanner();
        onDetected(isbn);
      }
    } catch (err) {
      console.warn('[scanner] decode error:', err);
    } finally {
      busy = false;
    }
  }, SCAN_INTERVAL_MS);
}

// 擷取畫面上「看得到」的區域（video 用 object-fit: cover 會裁切），送去辨識
async function _scanFrame() {
  const vw = _video.videoWidth, vh = _video.videoHeight;
  const ew = _video.clientWidth, eh = _video.clientHeight;
  if (!vw || !vh || !ew || !eh) return null;

  const scale = Math.max(ew / vw, eh / vh);   // object-fit: cover 的縮放比例
  const sw = Math.min(vw, ew / scale);
  const sh = Math.min(vh, eh / scale);
  const sx = (vw - sw) / 2;
  const sy = (vh - sh) / 2;

  const outW = Math.round(Math.min(sw, SCAN_MAX_WIDTH));
  const outH = Math.round(sh * (outW / sw));

  _canvas ??= document.createElement('canvas');
  _canvas.width  = outW;
  _canvas.height = outH;
  const ctx = _canvas.getContext('2d', { willReadFrequently: true });
  ctx.drawImage(_video, sx, sy, sw, sh, 0, 0, outW, outH);
  const imageData = ctx.getImageData(0, 0, outW, outH);

  const results = await ZXingWASM.readBarcodes(imageData, {
    formats: ['EAN13'],
    tryHarder: true,
    tryRotate: true,
    maxNumberOfSymbols: 3,
  });
  const hit = results.find(r => r.isValid && /^97[89]\d{10}$/.test(r.text));
  return hit ? hit.text : null;
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
  if (_timer) { clearInterval(_timer); _timer = null; }
  if (_stream) { _stream.getTracks().forEach(t => t.stop()); _stream = null; }
  if (_video) { _video.srcObject = null; _video.remove(); _video = null; }
  const reader = document.getElementById('reader');
  if (reader) reader.innerHTML = '';
}
