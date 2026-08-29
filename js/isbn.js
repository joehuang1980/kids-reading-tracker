// ISBN normalization and validation (Stage 3)

function normalizeIsbn(raw) {
  return raw.toUpperCase().replace(/[^0-9X]/g, '');
}

function isValidIsbnShape(isbn) {
  return /^\d{13}$/.test(isbn) || /^\d{9}[0-9X]$/.test(isbn);
}
