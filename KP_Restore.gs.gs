/**
 * KP_Restore.gs
 *
 * Восстановление КП из выбранной строки листа "Журнал КП".
 *
 * Логика:
 * 1) берём активную строку в "Журнал КП";
 * 2) запрещаем восстановление, если КП уже ушло в производство;
 * 3) очищаем количества на листе "Прайс" и проставляем их по UID из JSON;
 * 4) очищаем рабочую область листа "КП";
 * 5) запускаем существующую buildKP({ skipConfirm: true, mode: 'reset' });
 * 6) ПОСЛЕ buildKP() записываем шапку и ручные настройки из журнала;
 * 7) ПОСЛЕ buildKP() возвращаем примечания и индивидуальные скидки по UID.
 *
 * ВАЖНО:
 * - buildKP() должна уже существовать в проекте.
 * - onOpen() не трогаем здесь.
 * - Для меню просто добавьте в ваш текущий onOpen/main:
 *     .addItem('Восстановить КП из журнала', 'restoreKPFromJournal')
 */

const KP_RESTORE_CFG = {
  SHEETS: {
    JOURNAL_KP: 'Журнал КП',
    JOURNAL_PROD: 'Журнал заявок на производство',
    KP: 'КП',
    PRICE: 'Прайс'
  },

  KP_CELLS: {
    CUSTOMER: 'D10',
    ADDRESS: 'D11',
    CONTRACT_NO: 'D12',
    MANAGER: 'D13',
    // D14 не трогаем, там телефон подтягивается формулой от менеджера
    DATE_KP: 'D15',
    KP_NO: 'D16',
    DISCOUNT_GLOBAL: 'D19',
    INSTALL_PCT: 'D20',
    DELIVERY: 'D21',
    PREPAY_PCT: 'J24',
    TERM_MAIN: 'J25',
    TERM_ECO: 'J26',
    VALID_DAYS: 'J27'
  },

  KP_TABLE: {
    HEADER_ROW: 30,
    START_ROW: 31,
    FIRST_COL: 1,      // A
    COLS_COUNT: 14,    // A:N
    NOTE_COL: 12,      // L
    DISCOUNT_COL: 13,  // M
    UID_COL: 14        // N
  },

  PRICE: {
    UID_COL: 1,        // A
    QTY_COL: 9,        // I
    FIRST_DATA_ROW: 2
  },

  JOURNAL_HEADERS: {
    KP_NO: 'КП №',
    DATE_KP: 'Дата КП',
    MANAGER: 'Менеджер',
    PHONE: 'Телефон',
    CUSTOMER: 'Заказчик',
    ADDRESS: 'Адрес заказчика',
    CONTRACT_NO: '№ договора',
    DISCOUNT_GLOBAL: 'Скидка (-) / Наценка (+), %',
    INSTALL_PCT: 'Размер монтажа от стоимости оборудования, %',
    DELIVERY: 'Доставка, руб',
    PREPAY_PCT: 'Предоплата, %',
    TERM_MAIN: 'Срок (Основное)',
    TERM_ECO: 'Срок (ЭКО)',
    VALID_DAYS: 'КП действительно, дней',
    POSITIONS_JSON: 'Позиции (JSON)',
    DRIVE_FILE_ID: 'Drive File ID',
    STATUS: 'Статус'
  },

  PROD_HEADERS: {
    KP_DRIVE_FILE_ID: 'КП Drive File ID'
  },

  BLOCKED_STATUS: 'Отправлена в производство'
};


/**
 * Главная функция восстановления.
 * Привязать к пункту меню.
 */
function restoreKPFromJournal() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  try {
    const activeSheet = ss.getActiveSheet();
    if (!activeSheet || activeSheet.getName() !== KP_RESTORE_CFG.SHEETS.JOURNAL_KP) {
      ui.alert('Перейдите на лист "Журнал КП", выберите нужную строку и повторите команду.');
      return;
    }

    const activeRange = activeSheet.getActiveRange();
    if (!activeRange) {
      ui.alert('Не выбрана строка в "Журнал КП".');
      return;
    }

    const row = activeRange.getRow();
    if (row <= 1) {
      ui.alert('Выберите строку с данными, а не строку заголовков.');
      return;
    }

    const journalRecord = getJournalRecordFromActiveRow_(activeSheet, row);
    validateJournalRecord_(journalRecord);

    const prodSheet = ss.getSheetByName(KP_RESTORE_CFG.SHEETS.JOURNAL_PROD);
    const blockReason = getRestoreBlockReason_(journalRecord, prodSheet);
    if (blockReason) {
      ui.alert(blockReason);
      return;
    }

    const positions = parseAndNormalizePositionsJson_(
      journalRecord[KP_RESTORE_CFG.JOURNAL_HEADERS.POSITIONS_JSON]
    );

    if (!positions.length) {
      ui.alert('В выбранной строке не найдено ни одной позиции для восстановления.');
      return;
    }

    const confirm = ui.alert(
      'Восстановление КП',
      'Текущее содержимое листов "КП" и "Прайс" будет заменено данными из выбранной строки журнала. Продолжить?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    const priceSheet = ss.getSheetByName(KP_RESTORE_CFG.SHEETS.PRICE);
    const kpSheet = ss.getSheetByName(KP_RESTORE_CFG.SHEETS.KP);

    if (!priceSheet) throw new Error('Не найден лист "Прайс".');
    if (!kpSheet) throw new Error('Не найден лист "КП".');
    if (typeof buildKP !== 'function') {
      throw new Error('В проекте не найдена функция buildKP().');
    }

    // 1. Проставляем количества в Прайс
    const missingInPrice = restorePriceQuantitiesByUid_(priceSheet, positions);

    // 2. Очищаем КП
    clearKPForRestore_(kpSheet);
    SpreadsheetApp.flush();

    // 3. Формируем КП через существующую функцию
    buildKP({ skipConfirm: true, mode: 'reset' });
    SpreadsheetApp.flush();

    // 4. ПОСЛЕ buildKP() возвращаем ручные данные шапки и настроек
    applyJournalHeaderToKP_(kpSheet, journalRecord);
    SpreadsheetApp.flush();

    // 5. ПОСЛЕ buildKP() возвращаем примечания и построчные скидки
    const missingInKp = applyLineOverridesToKPByUid_(kpSheet, positions);
    SpreadsheetApp.flush();

    ss.setActiveSheet(kpSheet);
    kpSheet.setActiveSelection(KP_RESTORE_CFG.KP_CELLS.CUSTOMER);

    const problems = [];
    if (missingInPrice.length) {
      problems.push('Не найдены в "Прайс" UID: ' + missingInPrice.join(', '));
    }
    if (missingInKp.length) {
      problems.push('Не найдены в строках "КП" UID: ' + missingInKp.join(', '));
    }

    let msg = 'КП восстановлено из выбранной строки журнала.';
    if (problems.length) {
      msg += '\n\n' + problems.join('\n');
    }

    ui.alert('Готово', msg, ui.ButtonSet.OK);

  } catch (err) {
    Logger.log('restoreKPFromJournal error: ' + (err && err.stack ? err.stack : err));
    ui.alert('Ошибка восстановления', String(err && err.message ? err.message : err), ui.ButtonSet.OK);
  }
}


/**
 * Чтение активной строки журнала в объект по заголовкам.
 */
function getJournalRecordFromActiveRow_(journalSheet, row) {
  const lastCol = journalSheet.getLastColumn();
  const headers = journalSheet
    .getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(v => String(v).trim());

  const values = journalSheet.getRange(row, 1, 1, lastCol).getValues()[0];
  return objectFromHeadersAndValues_(headers, values);
}


/**
 * Проверка обязательных полей строки журнала.
 */
function validateJournalRecord_(record) {
  const h = KP_RESTORE_CFG.JOURNAL_HEADERS;

  const required = [
    h.KP_NO,
    h.CUSTOMER,
    h.POSITIONS_JSON
  ];

  const missing = required.filter(key => {
    const v = record[key];
    return v === '' || v === null || typeof v === 'undefined';
  });

  if (missing.length) {
    throw new Error('В строке журнала отсутствуют обязательные поля: ' + missing.join(', '));
  }
}


/**
 * Причина блокировки восстановления.
 */
function getRestoreBlockReason_(journalRecord, prodSheet) {
  const h = KP_RESTORE_CFG.JOURNAL_HEADERS;

  const status = normalizeString_(journalRecord[h.STATUS]);
  if (status === normalizeString_(KP_RESTORE_CFG.BLOCKED_STATUS)) {
    return 'Восстановление невозможно: данное КП уже имеет статус "Отправлена в производство".';
  }

  const driveFileId = normalizeString_(journalRecord[h.DRIVE_FILE_ID]);
  if (driveFileId && prodSheet && isDriveFileIdAlreadyInProduction_(prodSheet, driveFileId)) {
    return 'Восстановление невозможно: по данному КП уже создана/передана заявка в производство.';
  }

  return '';
}


/**
 * Проверка наличия Drive File ID КП в журнале производства.
 */
function isDriveFileIdAlreadyInProduction_(prodSheet, driveFileId) {
  const lastRow = prodSheet.getLastRow();
  const lastCol = prodSheet.getLastColumn();
  if (lastRow < 2) return false;

  const headers = prodSheet
    .getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(v => String(v).trim());

  const colIndex = headers.indexOf(KP_RESTORE_CFG.PROD_HEADERS.KP_DRIVE_FILE_ID);
  if (colIndex === -1) {
    throw new Error('В листе "Журнал заявок на производство" не найдена колонка "КП Drive File ID".');
  }

  const values = prodSheet
    .getRange(2, colIndex + 1, lastRow - 1, 1)
    .getDisplayValues()
    .flat();

  const target = normalizeString_(driveFileId);
  return values.some(v => normalizeString_(v) === target);
}


/**
 * Разбор JSON позиций.
 * На выходе массив объектов:
 * [
 *   { uid, qty, note, discountPct }
 * ]
 */
function parseAndNormalizePositionsJson_(jsonText) {
  let parsed;
  try {
    parsed = JSON.parse(String(jsonText || '[]'));
  } catch (e) {
    throw new Error('Не удалось разобрать поле "Позиции (JSON)".');
  }

  if (!Array.isArray(parsed)) {
    throw new Error('"Позиции (JSON)" не является массивом.');
  }

  const map = {};

  parsed.forEach(item => {
    const uid = normalizeUid_(item && item.uid);
    if (!uid) return;

    const qty = toNumber_(item.qty);
    const note = item && typeof item.note !== 'undefined' ? String(item.note) : '';
    const hasDiscount = item && typeof item.discountPct !== 'undefined' && item.discountPct !== '';
    const discountPct = hasDiscount ? toNumber_(item.discountPct) : null;

    if (!map[uid]) {
      map[uid] = {
        uid: uid,
        qty: 0,
        note: '',
        discountPct: null
      };
    }

    map[uid].qty += qty;

    // Берём последнее непустое примечание
    if (note !== '') {
      map[uid].note = note;
    }

    // Берём последнее явно сохранённое значение скидки
    if (hasDiscount) {
      map[uid].discountPct = discountPct;
    }
  });

  return Object.keys(map)
    .map(uid => map[uid])
    .filter(item => item.qty > 0);
}


/**
 * Очистка количеств в Прайсе и установка новых по UID.
 */
function restorePriceQuantitiesByUid_(priceSheet, positions) {
  const cfg = KP_RESTORE_CFG.PRICE;
  const lastRow = priceSheet.getLastRow();
  if (lastRow < cfg.FIRST_DATA_ROW) return [];

  const rowCount = lastRow - cfg.FIRST_DATA_ROW + 1;

  const uidValues = priceSheet
    .getRange(cfg.FIRST_DATA_ROW, cfg.UID_COL, rowCount, 1)
    .getDisplayValues()
    .flat()
    .map(normalizeUid_);

  const uidToRow = {};
  const qtyClearRanges = [];

  uidValues.forEach((uid, idx) => {
    const row = cfg.FIRST_DATA_ROW + idx;
    if (uid) {
      uidToRow[uid] = row;
      qtyClearRanges.push('I' + row);
    }
  });

  if (qtyClearRanges.length) {
    priceSheet.getRangeList(qtyClearRanges).clearContent();
  }

  const missing = [];

  positions.forEach(pos => {
    const row = uidToRow[pos.uid];
    if (!row) {
      missing.push(pos.uid);
      return;
    }
    priceSheet.getRange(row, cfg.QTY_COL).setValue(pos.qty);
  });

  return uniqueStrings_(missing);
}


/**
 * Очистка рабочего листа КП перед восстановлением.
 * D14 не трогаем, там формула телефона.
 */
function clearKPForRestore_(kpSheet) {
  const c = KP_RESTORE_CFG.KP_CELLS;
  const t = KP_RESTORE_CFG.KP_TABLE;

  kpSheet.getRangeList([
    c.CUSTOMER,
    c.ADDRESS,
    c.CONTRACT_NO,
    c.MANAGER,
    c.DATE_KP,
    c.KP_NO,
    c.DISCOUNT_GLOBAL,
    c.INSTALL_PCT,
    c.DELIVERY,
    c.PREPAY_PCT,
    c.TERM_MAIN,
    c.TERM_ECO,
    c.VALID_DAYS
  ]).clearContent();

  const maxRows = kpSheet.getMaxRows();
  const rowsToClear = maxRows - t.START_ROW + 1;
  if (rowsToClear > 0) {
    kpSheet
      .getRange(t.START_ROW, t.FIRST_COL, rowsToClear, t.COLS_COUNT)
      .clearContent();
  }
}


/**
 * Заполнение шапки и ручных настроек в КП.
 * ВАЖНО: вызывать после buildKP().
 */
function applyJournalHeaderToKP_(kpSheet, record) {
  const c = KP_RESTORE_CFG.KP_CELLS;
  const h = KP_RESTORE_CFG.JOURNAL_HEADERS;

  kpSheet.getRange(c.CUSTOMER).setValue(record[h.CUSTOMER] || '');
  kpSheet.getRange(c.ADDRESS).setValue(record[h.ADDRESS] || '');
  kpSheet.getRange(c.CONTRACT_NO).setValue(record[h.CONTRACT_NO] || '');
  kpSheet.getRange(c.MANAGER).setValue(record[h.MANAGER] || '');
  kpSheet.getRange(c.DATE_KP).setValue(record[h.DATE_KP] || '');
  kpSheet.getRange(c.KP_NO).setValue(record[h.KP_NO] || '');

  kpSheet.getRange(c.DISCOUNT_GLOBAL).setValue(toNumberOrBlank_(record[h.DISCOUNT_GLOBAL]));
  kpSheet.getRange(c.INSTALL_PCT).setValue(toNumberOrBlank_(record[h.INSTALL_PCT]));
  kpSheet.getRange(c.DELIVERY).setValue(toNumberOrBlank_(record[h.DELIVERY]));

  let prepay = toNumberOrBlank_(record[h.PREPAY_PCT]);
  if (prepay !== '' && Math.abs(prepay) > 1) {
    prepay = prepay / 100;
  }
  kpSheet.getRange(c.PREPAY_PCT).setValue(prepay);

  kpSheet.getRange(c.TERM_MAIN).setValue(record[h.TERM_MAIN] || '');
  kpSheet.getRange(c.TERM_ECO).setValue(record[h.TERM_ECO] || '');
  kpSheet.getRange(c.VALID_DAYS).setValue(toNumberOrBlank_(record[h.VALID_DAYS]));
}


/**
 * После buildKP() возвращаем примечания и индивидуальные скидки по UID.
 *
 * Важно:
 * - если discountPct в JSON отсутствует, колонку M НЕ трогаем,
 *   чтобы осталась формула/наследование от общей скидки.
 */
function applyLineOverridesToKPByUid_(kpSheet, positions) {
  const cfg = KP_RESTORE_CFG.KP_TABLE;
  const lastRow = kpSheet.getLastRow();
  if (lastRow < cfg.START_ROW) return positions.map(p => p.uid);

  const uidValues = kpSheet
    .getRange(cfg.START_ROW, cfg.UID_COL, lastRow - cfg.START_ROW + 1, 1)
    .getDisplayValues()
    .flat()
    .map(normalizeUid_);

  const uidToRow = {};
  uidValues.forEach((uid, idx) => {
    if (uid) uidToRow[uid] = cfg.START_ROW + idx;
  });

  const missing = [];

  positions.forEach(pos => {
    const row = uidToRow[pos.uid];
    if (!row) {
      missing.push(pos.uid);
      return;
    }

    // Примечание возвращаем всегда
    kpSheet.getRange(row, cfg.NOTE_COL).setValue(pos.note || '');

    // Скидку строки ставим только если она явно есть в JSON
    if (pos.discountPct !== null && pos.discountPct !== '' && typeof pos.discountPct !== 'undefined') {
      kpSheet.getRange(row, cfg.DISCOUNT_COL).setValue(toNumber_(pos.discountPct));
    }
  });

  return uniqueStrings_(missing);
}


/**
 * Собирает объект по массивам заголовков и значений.
 */
function objectFromHeadersAndValues_(headers, values) {
  const obj = {};
  headers.forEach((header, i) => {
    obj[header] = values[i];
  });
  return obj;
}


/**
 * Нормализация UID.
 */
function normalizeUid_(value) {
  return String(value == null ? '' : value).trim().toUpperCase();
}


/**
 * Нормализация обычной строки.
 */
function normalizeString_(value) {
  return String(value == null ? '' : value)
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}


/**
 * Преобразование в число.
 */
function toNumber_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return 0;
  if (typeof value === 'number') return value;

  const cleaned = String(value)
    .replace(/\s/g, '')
    .replace(',', '.');

  const n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}


/**
 * Преобразование в число или пусто.
 */
function toNumberOrBlank_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return '';
  if (typeof value === 'number') return value;

  const cleaned = String(value)
    .replace(/\s/g, '')
    .replace(',', '.');

  const n = Number(cleaned);
  return isNaN(n) ? '' : n;
}


/**
 * Уникальные непустые строки.
 */
function uniqueStrings_(arr) {
  return [...new Set((arr || []).filter(Boolean))];
}