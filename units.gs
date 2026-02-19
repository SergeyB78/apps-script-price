/**
 * utils.gs — общий набор утилит (безопасно: само по себе ничего не меняет).
 * Дальше будем постепенно использовать эти функции вместо копипаста.
 */

/** Простая проверка с понятной ошибкой */
function assert_(condition, message) {
  if (!condition) throw new Error(message || 'Assertion failed');
}

/** null/undefined/пустая строка */
function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '';
}

/** Нормализатор строк */
function s_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

/** Преобразование "1 234,56" / "1234.56" в число */
function toNumber_(v, def = 0) {
  if (v === null || v === undefined || v === '') return def;
  if (typeof v === 'number') return isNaN(v) ? def : v;

  const str = String(v).trim().replace(/\s+/g, '').replace(',', '.');
  const num = Number(str);
  return isNaN(num) ? def : num;
}

/** Получить лист по имени (с понятной ошибкой, если не найден) */
function getSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  assert_(sh, `Не найден лист: "${sheetName}"`);
  return sh;
}

/**
 * Получить карту заголовков (1-я строка): { "Имя колонки": индекс }
 * Индекс — 0-based, т.е. как в массиве values[0][idx]
 */
function getHeaderMap_(sheet, headerRow = 1) {
  const lastCol = sheet.getLastColumn();
  assert_(lastCol > 0, `Лист "${sheet.getName()}" пустой (нет колонок).`);

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = s_(h);
    if (key) map[key] = i;
  });
  return map;
}

/**
 * Найти индекс колонки по одному из возможных заголовков.
 * Пример: findColIdx_(map, [CFG.DB_HEADERS.STATUS, 'Статус'])
 */
function findColIdx_(headerMap, candidates) {
  for (const name of candidates) {
    if (headerMap.hasOwnProperty(name)) return headerMap[name];
  }
  return null;
}

/** Проверка "активности" по значению из ячейки (true / TRUE / "Активная позиция") */
function isActiveValue_(v) {
  const t = String(v === null || v === undefined ? '' : v).trim().toLowerCase();
  if (t === 'true') return true;
  if (v === true) return true;

  // если CFG есть (будет в проекте) — сравниваем с CFG.DB_VALUES.ACTIVE_STATUS
  try {
    const active = String(CFG?.DB_VALUES?.ACTIVE_STATUS || '').trim().toLowerCase();
    if (active && t === active) return true;
  } catch (e) {}

  return false;
}

/**
 * Обертка на LockService (чтобы случайно 2 запуска не конфликтовали).
 * Использование потом:
 * withLock_(() => buildKP(), 30000);
 */
function withLock_(fn, timeoutMs = 30000) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(timeoutMs);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
