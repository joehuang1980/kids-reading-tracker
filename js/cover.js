// 書封拍照：把相機拍的照片縮成小張 JPEG（data URL），直接存進試算表的 Thumbnail 欄位。
// 試算表單一儲存格上限約 50,000 字元，所以壓到 MAX_COVER_CHARS 以內。

const MAX_COVER_CHARS = 40000;

async function makeCoverThumbnail(file, maxW = 300, maxH = 400) {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((resolve, reject) => {
      const i = new Image();
      i.onload  = () => resolve(i);
      i.onerror = () => reject(new Error('無法讀取照片'));
      i.src = url;
    });

    let scale = Math.min(1, maxW / img.naturalWidth, maxH / img.naturalHeight);
    const canvas = document.createElement('canvas');
    for (let attempt = 0; attempt < 6; attempt++) {
      canvas.width  = Math.max(1, Math.round(img.naturalWidth  * scale));
      canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      for (const q of [0.8, 0.7, 0.6, 0.5]) {
        const dataUrl = canvas.toDataURL('image/jpeg', q);
        if (dataUrl.length <= MAX_COVER_CHARS) return dataUrl;
      }
      scale *= 0.8;   // 還是太大就再縮小
    }
    throw new Error('照片太大');
  } finally {
    URL.revokeObjectURL(url);
  }
}
