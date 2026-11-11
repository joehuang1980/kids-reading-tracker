// ISBN normalization and validation (Stage 3)

function normalizeIsbn(raw) {
  return raw.toUpperCase().replace(/[^0-9X]/g, '');
}

function isValidIsbnShape(isbn) {
  return /^\d{13}$/.test(isbn) || /^\d{9}[0-9X]$/.test(isbn);
}

// 英語出版品的 ISBN 群組：978-0、978-1、979-8（ISBN-10 開頭 0 或 1）
function isEnglishIsbn(isbn) {
  return /^97(8[01]|98)\d+$/.test(isbn) || /^[01]\d{8}[\dX]$/.test(isbn);
}

// 有書目語言就以語言判斷，否則依 ISBN 群組判斷
function isEnglishBook(isbn, book) {
  if (book?.language) return /^en/i.test(book.language);
  return isEnglishIsbn(isbn);
}
