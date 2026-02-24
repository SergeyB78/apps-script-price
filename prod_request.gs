/** prod_request.gs
 * Заявка на производство по выбранной строке "Журнал КП"
 *
 * ВАЖНО:
 * - НЕ создаёт и НЕ дорабатывает "Журнал КП" (без дублирования существующей логики)
 * - Создаёт/доращивает только "Журнал заявок на производство"
 * - Работает с 1 колонкой статуса в "Журнал КП":
 *   "Статус" (предпочтительно) или "Статус КП" (fallback)
 */

var PROD_REQ_CFG = PROD_REQ_CFG || {
  // Листы (через централизованные схемы, если доступны)
  KP_LOG_SHEET:
    (typeof getSheetSchemaName_ === 'function')
      ? getSheetSchemaName_('KP_LOG')
      : 'Журнал КП',

  PROD_LOG_SHEET:
    (typeof getSheetSchemaName_ === 'function')
      ? getSheetSchemaName_('PROD_REQUEST_LOG')
      : 'Журнал заявок на производство',

  // Статусы (одна колонка в Журнале КП)
  STATUS_VALUES: ['Новая', 'Аннулирована', 'Отправлена в производство'],
  STATUS_NEW: 'Новая',
  STATUS_CANCELED: 'Аннулирована',
  STATUS_SENT: 'Отправлена в производство',

  // Возможные названия колонки статуса
  STATUS_HEADER_ALIASES: ['Статус', 'Статус КП'],

  // Заголовки журнала заявок на производство (из схемы, если доступна)
  PROD_LOG_HEADERS:
    (typeof getSheetSchemaHeaders_ === 'function')
      ? getSheetSchemaHeaders_('PROD_REQUEST_LOG')
      : [
          'Дата/время создания',
          '№ заявки на производство',
          'КП №',
          'Дата КП',
          'Заказчик',
          'Адрес Заказчика',
          'Менеджер',
          'Телефон',
          'Drive File ID КП',
          'PDF URL (Drive)',
          'Кол-во позиций',
          'Сумма количеств',
          'UID_list',
          'Позиции (JSON)',
          'Статус заявки',
          'Комментарий'
        ]
};

/* ========================================================================== */
/* ПУБЛИЧНЫЕ ФУНКЦИИ                                                            */
/* ========================================================================== */

/**
 * Основная команда меню:
 * создать заявку на производство по активной строке листа "Журнал КП"
 */
function createProductionOrderFromSelectedKp() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    ui.alert('Не удалось получить блокировку документа. Повторите через несколько секунд.');
    return;
  }

  try {
    // Создаём/доращиваем только журнал заявок на производство
    var prodLog = ensureProductionLogSheet_();

    var sh = ss.getActiveSheet();
    if (!sh || sh.getName() !== PROD_REQ_CFG.KP_LOG_SHEET) {
      ui.alert(
        'Откройте лист "' + PROD_REQ_CFG.KP_LOG_SHEET + '" и выберите строку КП, ' +
        'по которой нужно создать заявку на производство.'
      );
      return;
    }

    var activeRange = sh.getActiveRange();
    if (!activeRange) {
      ui.alert('Выберите строку в листе "' + PROD_REQ_CFG.KP_LOG_SHEET + '".');
      return;
    }

    var row = activeRange.getRow();
    if (row < 2) {
      ui.alert('Выберите строку данных (не заголовок).');
      return;
    }

    var kpMap = getHeaderMap_(sh);
    if (!Object.keys(kpMap).length) {
      ui.alert('Не удалось прочитать заголовки листа "' + PROD_REQ_CFG.KP_LOG_SHEET + '".');
      return;
    }

    var col = resolveKpColumns_(kpMap);
    validateRequiredKpColumns_(col);

    var rowObj = readRowObject_(sh, row);

    var kpNo = strv_(rowObj[col.kpNo]);
    var customer = strv_(rowObj[col.customer]);
    var fileId = strv_(rowObj[col.driveFileId]);
    var positionsJsonRaw = rowObj[col.positionsJson];
    var currentStatus = strv_(rowObj[col.status]) || PROD_REQ_CFG.STATUS_NEW;

    // Проверка обязательных значений в выбранной строке
    var missingValues = [];
    if (!kpNo) missingValues.push('КП №');
    if (!customer) missingValues.push('Заказчик');
    if (!fileId) missingValues.push('Drive File ID');
    if (!strv_(positionsJsonRaw)) missingValues.push('Позиции (JSON)');

    if (missingValues.length) {
      ui.alert(
        'Нельзя создать заявку на производство.\n\n' +
        'В выбранной строке не заполнены поля:\n- ' + missingValues.join('\n- ')
      );
      return;
    }

    // Проверка статуса выбранной строки
    if (currentStatus === PROD_REQ_CFG.STATUS_CANCELED) {
      ui.alert('У выбранной записи статус "' + PROD_REQ_CFG.STATUS_CANCELED + '". Создание заявки запрещено.');
      return;
    }
    if (currentStatus === PROD_REQ_CFG.STATUS_SENT) {
      ui.alert('По выбранной записи уже создана заявка на производство (статус "' + PROD_REQ_CFG.STATUS_SENT + '").');
      return;
    }

    // Парсим "Позиции (JSON)"
    var parsed = parsePositionsJsonSafe_(positionsJsonRaw);
    if (!parsed.ok) {
      ui.alert(
        'Невозможно создать заявку: поле "Позиции (JSON)" заполнено некорректно.\n\n' +
        parsed.error
      );
      return;
    }

    var stats = summarizePositionsForProduction_(parsed.items);
    if (stats.itemsCount <= 0) {
      ui.alert('Невозможно создать заявку: в "Позиции (JSON)" не найдено позиций для производства.');
      return;
    }

    // Находим дубли по (КП № + Заказчик)
    var duplicates = findDuplicateRowsByKpNoAndCustomer_(sh, kpMap, col, kpNo, customer);

    // Если среди дублей уже есть запись "Отправлена в производство" (не текущая) — блокируем
    var existingSent = null;
    for (var i = 0; i < duplicates.length; i++) {
      var d = duplicates[i];
      if (d.row === row) continue;
      if (d.status === PROD_REQ_CFG.STATUS_SENT) {
        existingSent = d;
        break;
      }
    }

    if (existingSent) {
      ui.alert(
        'По этой группе КП (Клиент + № КП) уже есть заявка в производство.\n\n' +
        'Строка: ' + existingSent.row + '\n' +
        'Статус: ' + existingSent.status + '\n\n' +
        'Повторное создание запрещено.'
      );
      return;
    }

    // Ввод номера заявки
    var suggestedNo = suggestProductionRequestNo_(kpNo);
    var promptRes = ui.prompt(
      'Заявка на производство',
      'Введите номер заявки на производство.\n\n' +
      'КП №: ' + kpNo + '\n' +
      'Заказчик: ' + customer + '\n' +
      'Позиции: ' + stats.itemsCount + '\n' +
      'Сумма количеств: ' + stats.qtySum + '\n\n' +
      'Если оставить пусто — будет подставлен номер:\n' + suggestedNo,
      ui.ButtonSet.OK_CANCEL
    );

    if (promptRes.getSelectedButton() !== ui.Button.OK) return;

    var reqNo = strv_(promptRes.getResponseText());
    if (!reqNo) reqNo = suggestedNo;

    // Финальное подтверждение
    var confirmText =
      'Будет создана заявка на производство:\n\n' +
      '№ заявки: ' + reqNo + '\n' +
      'КП №: ' + kpNo + '\n' +
      'Заказчик: ' + customer + '\n' +
      'Позиции: ' + stats.itemsCount + '\n' +
      'Сумма количеств: ' + stats.qtySum + '\n' +
      'UID: ' + (stats.uidList || '-') + '\n\n' +
      'После подтверждения:\n' +
      '• выбранная запись → "' + PROD_REQ_CFG.STATUS_SENT + '"\n' +
      '• дубли (тот же клиент + тот же № КП) → "' + PROD_REQ_CFG.STATUS_CANCELED + '"';

    var ok = ui.alert('Подтверждение', confirmText, ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;

    // Запись в "Журнал заявок на производство"
    var prodMap = getHeaderMap_(prodLog);

    var prodRowObj = {};
    prodRowObj['Дата/время создания'] = new Date();
    prodRowObj['№ заявки на производство'] = reqNo;
    prodRowObj['КП №'] = kpNo;
    prodRowObj['Дата КП'] = col.kpDate ? (rowObj[col.kpDate] || '') : '';
    prodRowObj['Заказчик'] = customer;
    prodRowObj['Адрес Заказчика'] = col.customerAddr ? (rowObj[col.customerAddr] || '') : '';
    prodRowObj['Менеджер'] = col.manager ? (rowObj[col.manager] || '') : '';
    prodRowObj['Телефон'] = col.phone ? (rowObj[col.phone] || '') : '';
    prodRowObj['Drive File ID КП'] = fileId;
    prodRowObj['PDF URL (Drive)'] = col.pdfUrl ? (rowObj[col.pdfUrl] || '') : '';
    prodRowObj['Кол-во позиций'] = stats.itemsCount;
    prodRowObj['Сумма количеств'] = stats.qtySum;
    prodRowObj['UID_list'] = stats.uidList;
    prodRowObj['Позиции (JSON)'] = parsed.rawString;
    prodRowObj['Статус заявки'] = 'Создана';
    prodRowObj['Комментарий'] = '';

    var prodRow = appendObjectByHeaders_(prodLog, prodRowObj, prodMap);
    formatProductionLogRow_(prodLog, prodRow);

    // Обновляем статусы в "Журнал КП"
    setCellValue_(sh, row, col.status, PROD_REQ_CFG.STATUS_SENT);

    // Все дубли (кроме выбранной строки) -> Аннулирована
    for (var j = 0; j < duplicates.length; j++) {
      var dup = duplicates[j];
      if (dup.row === row) continue;
      setCellValue_(sh, dup.row, col.status, PROD_REQ_CFG.STATUS_CANCELED);
    }

    SpreadsheetApp.flush();

    ui.alert(
      'Готово.\n\n' +
      'Заявка на производство создана.\n' +
      'Строка в "' + PROD_REQ_CFG.PROD_LOG_SHEET + '": ' + prodRow + '\n' +
      'Статусы в "' + PROD_REQ_CFG.KP_LOG_SHEET + '" обновлены.'
    );

  } catch (err) {
    SpreadsheetApp.getUi().alert('Ошибка: ' + (err && err.message ? err.message : err));
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Можно запускать вручную один раз, чтобы создать/обновить "Журнал заявок на производство"
 * (Журнал КП не трогает)
 */
function ensureProductionLogSheetManual_() {
  ensureProductionLogSheet_();
  SpreadsheetApp.getUi().alert('Лист "' + PROD_REQ_CFG.PROD_LOG_SHEET + '" проверен/подготовлен.');
}

/* ========================================================================== */
/* ЖУРНАЛ ЗАЯВОК НА ПРОИЗВОДСТВО                                                */
/* ========================================================================== */

function ensureProductionLogSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PROD_REQ_CFG.PROD_LOG_SHEET);

  if (!sh) {
    sh = ss.insertSheet(PROD_REQ_CFG.PROD_LOG_SHEET);
    sh.getRange(1, 1, 1, PROD_REQ_CFG.PROD_LOG_HEADERS.length).setValues([PROD_REQ_CFG.PROD_LOG_HEADERS]);
    sh.setFrozenRows(1);
    safeAutoResize_(sh, 1, PROD_REQ_CFG.PROD_LOG_HEADERS.length);
    return sh;
  }

  // Если лист есть, но пустой
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, PROD_REQ_CFG.PROD_LOG_HEADERS.length).setValues([PROD_REQ_CFG.PROD_LOG_HEADERS]);
    sh.setFrozenRows(1);
    safeAutoResize_(sh, 1, PROD_REQ_CFG.PROD_LOG_HEADERS.length);
    return sh;
  }

  // Доращиваем отсутствующие колонки по заголовкам
  var map = getHeaderMap_(sh);
  var changed = false;

  for (var i = 0; i < PROD_REQ_CFG.PROD_LOG_HEADERS.length; i++) {
    var h = PROD_REQ_CFG.PROD_LOG_HEADERS[i];
    if (!map[h]) {
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(h);
      changed = true;
    }
  }

  if (changed) {
    sh.setFrozenRows(1);
    safeAutoResize_(sh, 1, sh.getLastColumn());
  }

  return sh;
}

/* ========================================================================== */
/* ПОИСК И ВАЛИДАЦИЯ КОЛОНОК В "ЖУРНАЛ КП"                                       */
/* ========================================================================== */

function resolveKpColumns_(headerMap) {
  return {
    // обязательные
    kpNo: findHeaderKeyByAliases_(headerMap, ['КП №', '№ КП']),
    customer: findHeaderKeyByAliases_(headerMap, ['Заказчик', 'Наименование Заказчика']),
    driveFileId: findHeaderKeyByAliases_(headerMap, ['Drive File ID КП', 'Drive File ID']),
    positionsJson: findHeaderKeyByAliases_(headerMap, ['Позиции (JSON)']),
    status: findHeaderKeyByAliases_(headerMap, PROD_REQ_CFG.STATUS_HEADER_ALIASES),

    // необязательные
    kpDate: findHeaderKeyByAliases_(headerMap, ['Дата КП']),
    customerAddr: findHeaderKeyByAliases_(headerMap, ['Адрес Заказчика', 'Адрес заказчика']),
    manager: findHeaderKeyByAliases_(headerMap, ['Менеджер']),
    phone: findHeaderKeyByAliases_(headerMap, ['Телефон']),
    pdfUrl: findHeaderKeyByAliases_(headerMap, ['PDF URL (Drive)', 'PDF URL', 'Ссылка на PDF'])
  };
}

function validateRequiredKpColumns_(col) {
  var missing = [];

  if (!col.kpNo) missing.push('КП № / № КП');
  if (!col.customer) missing.push('Заказчик');
  if (!col.driveFileId) missing.push('Drive File ID (или Drive File ID КП)');
  if (!col.positionsJson) missing.push('Позиции (JSON)');
  if (!col.status) missing.push('Статус (или Статус КП)');

  if (missing.length) {
    throw new Error(
      'В листе "Журнал КП" не найдены обязательные колонки:\n- ' +
      missing.join('\n- ') +
      '\n\nПроверьте названия заголовков.'
    );
  }
}

/* ========================================================================== */
/* ДУБЛИКАТЫ: (КП № + ЗАКАЗЧИК)                                                  */
/* ========================================================================== */

function findDuplicateRowsByKpNoAndCustomer_(sh, headerMap, col, kpNoValue, customerValue) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = sh.getLastColumn();
  var data = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  var kpNoIdx = headerMap[col.kpNo] - 1;
  var customerIdx = headerMap[col.customer] - 1;
  var statusIdx = headerMap[col.status] - 1;

  var kpNoKey = normText_(kpNoValue);
  var customerKey = normText_(customerValue);

  var out = [];

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var rowVals = data[i];

    var rowKpNo = normText_(rowVals[kpNoIdx]);
    var rowCustomer = normText_(rowVals[customerIdx]);

    if (!rowKpNo || !rowCustomer) continue;
    if (rowKpNo !== kpNoKey) continue;
    if (rowCustomer !== customerKey) continue;

    out.push({
      row: rowNum,
      status: strv_(rowVals[statusIdx]) || PROD_REQ_CFG.STATUS_NEW
    });
  }

  return out;
}

/* ========================================================================== */
/* ПАРСИНГ И АГРЕГАЦИЯ "Позиции (JSON)"                                           */
/* ========================================================================== */

function parsePositionsJsonSafe_(raw) {
  try {
    var s = typeof raw === 'string' ? raw : String(raw || '');
    s = s.trim();
    if (!s) return { ok: false, error: 'Пустое значение JSON.' };

    var parsed = JSON.parse(s);

    // Поддержка разных форматов:
    // 1) [ {...}, {...} ]
    // 2) { items: [ ... ] }
    var items = [];
    if (Array.isArray(parsed)) {
      items = parsed;
    } else if (parsed && Array.isArray(parsed.items)) {
      items = parsed.items;
    } else {
      return { ok: false, error: 'Ожидался массив позиций или объект с полем items.' };
    }

    return {
      ok: true,
      items: items,
      rawString: s
    };
  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : String(e)
    };
  }
}

function summarizePositionsForProduction_(items) {
  var itemsCount = 0;
  var qtySum = 0;
  var uidArr = [];

  if (!Array.isArray(items)) {
    return { itemsCount: 0, qtySum: 0, uidList: '' };
  }

  for (var i = 0; i < items.length; i++) {
    var it = items[i] || {};

    // Кол-во (ищем по типичным ключам)
    var qty = pickNumberByKeys_(it, [
      'qty', 'quantity', 'Кол-во', 'Количество',
      'count', 'qnty'
    ]);

    if (qty > 0) {
      itemsCount++;
      qtySum += qty;
    } else {
      // fallback: если qty отсутствует, но позиция есть — считаем как 1 позицию
      var hasArt = !!strv_(pickAny_(it, ['art', 'Артикул', 'sku']));
      var hasName = !!strv_(pickAny_(it, ['name', 'Наименование']));
      if (hasArt || hasName) {
        itemsCount++;
      }
    }

    // UID (ищем по типичным ключам)
    var uid = strv_(pickAny_(it, [
      'uid', 'UID', 'priceUid', 'price_uid', 'uid_price', 'UID_Прайс'
    ]));

    if (uid && uidArr.indexOf(uid) === -1) {
      uidArr.push(uid);
    }
  }

  return {
    itemsCount: itemsCount,
    qtySum: qtySum,
    uidList: uidArr.join(';')
  };
}

/* ========================================================================== */
/* ЗАПИСЬ В ЖУРНАЛ ЗАЯВОК ПО ЗАГОЛОВКАМ                                           */
/* ========================================================================== */

function appendObjectByHeaders_(sheet, obj, headerMap) {
  var sh = sheet;
  var map = headerMap || getHeaderMap_(sh);

  var lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error('Лист "' + sh.getName() + '" не содержит заголовков.');

  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var rowOut = [];

  for (var i = 0; i < headers.length; i++) {
    var h = strv_(headers[i]);
    rowOut.push(Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : '');
  }

  var nextRow = Math.max(2, sh.getLastRow() + 1);
  sh.getRange(nextRow, 1, 1, rowOut.length).setValues([rowOut]);

  return nextRow;
}

function formatProductionLogRow_(sh, row) {
  var map = getHeaderMap_(sh);

  var cTime = map['Дата/время создания'];
  if (cTime) sh.getRange(row, cTime).setNumberFormat('dd.MM.yyyy HH:mm:ss');

  var cJson = map['Позиции (JSON)'];
  if (cJson) sh.setColumnWidth(cJson, Math.max(sh.getColumnWidth(cJson), 380));

  var cUrl = map['PDF URL (Drive)'];
  if (cUrl) sh.setColumnWidth(cUrl, Math.max(sh.getColumnWidth(cUrl), 260));
}

/* ========================================================================== */
/* ОБЩИЕ HELPERS                                                                 */
/* ========================================================================== */

function getHeaderMap_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};

  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};

  for (var i = 0; i < headers.length; i++) {
    var h = strv_(headers[i]);
    if (h && !map[h]) map[h] = i + 1; // 1-based
  }

  return map;
}

function readRowObject_(sh, row) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];

  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    var h = strv_(headers[i]);
    if (!h) continue;
    obj[h] = vals[i];
  }

  return obj;
}

function findHeaderKeyByAliases_(headerMap, aliases) {
  if (!headerMap) return '';

  // Точное совпадение
  for (var i = 0; i < aliases.length; i++) {
    if (headerMap[aliases[i]]) return aliases[i];
  }

  // Нормализованный поиск
  var targetNorms = {};
  for (var j = 0; j < aliases.length; j++) {
    targetNorms[normHeader_(aliases[j])] = true;
  }

  var keys = Object.keys(headerMap);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    if (targetNorms[normHeader_(key)]) return key;
  }

  return '';
}

function setCellValue_(sh, row, headerName, value) {
  var map = getHeaderMap_(sh);
  var col = map[headerName];
  if (!col) throw new Error('Не найдена колонка "' + headerName + '" в листе "' + sh.getName() + '".');
  sh.getRange(row, col).setValue(value);
}

function suggestProductionRequestNo_(kpNo) {
  var tz = Session.getScriptTimeZone() || 'Asia/Yerevan';
  var ts = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmm');
  var kp = strv_(kpNo);

  if (kp) return 'ЗП-' + kp + '-' + ts;
  return 'ЗП-' + ts;
}

function safeAutoResize_(sh, col, numCols) {
  try {
    sh.autoResizeColumns(col, numCols);
  } catch (e) {}
}

function strv_(v) {
  return String(v == null ? '' : v).trim();
}

function normText_(v) {
  return strv_(v).replace(/\s+/g, ' ').toLowerCase();
}

function normHeader_(v) {
  return normText_(v).replace(/\n+/g, ' ');
}

function pickAny_(obj, keys) {
  if (!obj || typeof obj !== 'object') return '';
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (Object.prototype.hasOwnProperty.call(obj, k) && obj[k] != null && obj[k] !== '') {
      return obj[k];
    }
  }
  return '';
}

function pickNumberByKeys_(obj, keys) {
  var v = pickAny_(obj, keys);
  return toNumberLoosePR_(v);
}

function toNumberLoosePR_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;

  var s = String(v).replace(/\s+/g, '').replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}