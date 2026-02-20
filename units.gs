/**
 * utils.gs — общий набор утилит (пока опционально).
 */

function assert_(condition, message) {
  if (!condition) throw new Error(message || 'Assertion failed');
}

function getSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  assert_(sh, `Не найден лист: "${sheetName}"`);
  return sh;
}

function withLock_(fn, timeoutMs = 30000) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(timeoutMs);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
