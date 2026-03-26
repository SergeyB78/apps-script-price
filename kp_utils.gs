/*************** kp_utils.gs ***************/
/**
 * Общие утилиты (используются и PDF‑модулем, и модулем Журнала)
 */

/**
 * Ищем строку по лейблу в A (A:C merged) и берём значение из D
 */
function findValueByLabelInColD_(sh, label) {
  const target = String(label || '').trim();
  if (!target) return '';

  const maxScan = Math.min(250, sh.getLastRow());
  for (let r = 1; r <= maxScan; r++) {
    const a = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (a === target) return sh.getRange(r, 4).getValue(); // D
  }
  return '';
}

/**
 * В блоке условий: слева текст в A (A:I merged), справа значение в J (J:K merged)
 */
function findTermsValueRight_(sh, leftText) {
  const target = String(leftText || '').trim();
  if (!target) return '';

  const maxScan = Math.min(300, sh.getLastRow());
  for (let r = 1; r <= maxScan; r++) {
    const a = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (a === target) return sh.getRange(r, 10).getValue(); // J
  }
  return '';
}

/* ===================== NUM HELPERS ===================== */

function toNumber_(v) {
  if (typeof v === 'number') return v;
  const s = String(v ?? '').trim();
  if (!s) return 0;
  const norm = s.replace(/\s+/g, '').replace(',', '.');
  const n = Number(norm);
  return Number.isFinite(n) ? n : 0;
}

/**
 * Приводим проценты к доле 0..1.
 * Поддержка: 0.7 / 70 / "70%"
 */
function normalizePercent_(v) {
  if (typeof v === 'number') {
    if (v > 1.5) return v / 100;
    if (v < 0) return 0;
    return v;
  }
  const s = String(v || '').trim().replace('%', '');
  const n = toNumber_(s);
  if (!Number.isFinite(n)) return 0;
  if (n > 1.5) return n / 100;
  return n;
}

function isFiniteNumber_(n) {
  return typeof n === 'number' && Number.isFinite(n);
}
