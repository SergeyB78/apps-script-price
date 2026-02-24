/** 
 * kp_log.gs — сбор данных для журнала КП и запись в лист "Журнал КП"
 *
 * КЛЮЧЕВОЕ:
 * - appendToLog_() записывает строку ПО ЗАГОЛОВКАМ листа, чтобы значения не "съезжали"
 *   при удалении/перестановке/скрытии столбцов (например Drive File ID).
 * - Автоформат денежных колонок: "#,##0.00"
 * - Автостатус для новой записи в журнале КП: "Новая" (колонка "Статус")
 */

/* ========================= Helpers / Safe wrappers ========================= */

function _normHeader_(s) {
  return String(s || '')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _toNumberLoose_(v) {
  if (typeof toNumber_ === 'function') return toNumber_(v);

  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;

  var s = String(v).replace(/\s+/g, '').replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function _normalizePercentLoose_(v) {
  if (typeof normalizePercent_ === 'function') return normalizePercent_(v);

  var n = _toNumberLoose_(v);
  if (!n) return 0;
  return n > 1 ? n / 100 : n; // 70 -> 0.7
}

function _safeFindValueByLabelInColD_(sh, label) {
  if (typeof findValueByLabelInColD_ === 'function') return findValueByLabelInColD_(sh, label);

  // fallback: ищем подпись в A, значение в D
  var maxScan = Math.min(400, sh.getLastRow());
  for (var r = 1; r <= maxScan; r++) {
    var a = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (a === label) return sh.getRange(r, 4).getValue();
  }
  return '';
}

function _safeFindTermsValueRight_(sh, label) {
  if (typeof findTermsValueRight_ === 'function') return findTermsValueRight_(sh, label);

  // fallback: ищем подпись в A, значение в J
  var maxScan = Math.min(700, sh.getLastRow());
  for (var r = 1; r <= maxScan; r++) {
    var a = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (a === label) return sh.getRange(r, 10).getValue(); // J
  }
  return '';
}

/* ========================= API used by controller ========================= */

/**
 * Мета-данные для журнала КП.
 */
function extractMetaForLog_(sh) {
  var kpNo = String(_safeFindValueByLabelInColD_(sh, 'Коммерческое предложение №') || '').trim();
  var kpDate = _safeFindValueByLabelInColD_(sh, 'Дата КП');
  var manager = String(_safeFindValueByLabelInColD_(sh, 'Менеджер') || '').trim();
  var phone = String(_safeFindValueByLabelInColD_(sh, 'Телефон') || '').trim();
  var customer = String(_safeFindValueByLabelInColD_(sh, 'Наименование Заказчика') || '').trim();
  var customerAddr = String(_safeFindValueByLabelInColD_(sh, 'Адрес Заказчика') || '').trim();
  var contractNo = String(_safeFindValueByLabelInColD_(sh, '№ Договора') || '').trim();

  var discountPct = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Скидка (-) / Наценка (+), %'));
  var installPct = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Размер монтажа от стоимости оборудования, %'));

  var equipTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого за оборудование, руб'));
  if (!equipTotal) equipTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого оборудование, руб'));

  var installTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Монтаж, руб'));
  var delivery = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Доставка, руб'));
  var toPay = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого к оплате, руб'));

  var vat = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'НДС 22%, руб'));
  if (!vat) vat = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'В том числе НДС 22%, руб'));

  var prepayPctNorm = _normalizePercentLoose_(
    _safeFindTermsValueRight_(sh, 'Предоплата за оборудование составляет:')
  );

  var prepaySum = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Сумма предоплаты, руб'));
  if (!prepaySum && equipTotal && prepayPctNorm) prepaySum = equipTotal * prepayPctNorm;

  var mainLead =
    String(
      _safeFindTermsValueRight_(
        sh,
        'Срок поставки Основное производство исчисляется с момента поступления предоплаты на р/счет и составляет:'
      ) || ''
    ).trim() ||
    String(_safeFindValueByLabelInColD_(sh, 'Срок (Основное)') || '').trim();

  var ecoLead =
    String(
      _safeFindTermsValueRight_(
        sh,
        'Срок поставки ЭКО-серия исчисляется с момента поступления предоплаты на р/счет и составляет:'
      ) || ''
    ).trim() ||
    String(_safeFindValueByLabelInColD_(sh, 'Срок (ЭКО)') || '').trim();

  var validDays =
    _toNumberLoose_(_safeFindTermsValueRight_(sh, 'Данное КП действительно в течение:')) ||
    _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'КП действительно, дней'));

  return {
    timestamp: new Date(),
    kpNo: kpNo,
    kpDate: kpDate,
    manager: manager,
    phone: phone,
    customer: customer,
    customerAddr: customerAddr,
    contractNo: contractNo,
    discountPct: discountPct,
    installPct: installPct,
    equipTotal: equipTotal,
    installTotal: installTotal,
    delivery: delivery,
    toPay: toPay,
    vat: vat,
    prepayPctNorm: prepayPctNorm,
    prepaySum: prepaySum,
    mainLead: mainLead,
    ecoLead: ecoLead,
    validDays: validDays
  };
}

/**
 * Корзина → JSON (включая Примечание и индивидуальную скидку, если такие колонки есть).
 */
function extractCartAsJson_(sh) {
  var headerRow = _findCartHeaderRow_(sh, 'Артикул');
  if (!headerRow) return { items: [], jsonString: '[]' };

  var maxCol = Math.min(40, sh.getLastColumn());
  var hdr = sh.getRange(headerRow, 1, 1, maxCol).getDisplayValues()[0].map(_normHeader_);

  var colArt = _colByName_(hdr, ['артикул']) || 1;
  var colName = _colByName_(hdr, ['наименование']) || 4;
  var colUnit = _colByName_(hdr, ['ед. изм.', 'ед.изм.', 'ед.']) || 5;
  var colPrice = _colByName_(hdr, ['стоимость оборудования', 'цена', 'стоимость']) || 6;
  var colQty = _colByName_(hdr, ['кол-во', 'количество']) || 7;
  var colSumEquip = _colByName_(hdr, ['всего за оборудование', 'сумма за оборудование']) || 8;
  var colInstallUnit = _colByName_(hdr, ['стоимость монтажа']) || 9;
  var colSumInstall = _colByName_(hdr, ['всего за монтаж']) || 10;
  var colTotal = _colByName_(hdr, ['итого']) || 11;

  var colNote = _colByName_(hdr, ['примечание', 'комментарий']) || 0;
  var colDiscount = _colByContainsAll_(hdr, ['скидка', '%']) || 0;

  var items = [];
  var lastRow = sh.getLastRow();

  for (var r = headerRow + 1; r <= lastRow; r++) {
    var art = String(sh.getRange(r, colArt).getDisplayValue() || '').trim();
    var name = String(sh.getRange(r, colName).getDisplayValue() || '').trim();

    if (!art && !name) break;
    if (!art) continue;

    var unit = String(sh.getRange(r, colUnit).getDisplayValue() || '').trim();
    var price = _toNumberLoose_(sh.getRange(r, colPrice).getValue());
    var qty = _toNumberLoose_(sh.getRange(r, colQty).getValue());
    var sumEquip = _toNumberLoose_(sh.getRange(r, colSumEquip).getValue());
    var installUnit = _toNumberLoose_(sh.getRange(r, colInstallUnit).getValue());
    var sumInstall = _toNumberLoose_(sh.getRange(r, colSumInstall).getValue());
    var total = _toNumberLoose_(sh.getRange(r, colTotal).getValue());
    var note = colNote ? String(sh.getRange(r, colNote).getDisplayValue() || '').trim() : '';
    var discountPct = colDiscount ? _toNumberLoose_(sh.getRange(r, colDiscount).getValue()) : '';

    items.push({
      art: art,
      name: name,
      unit: unit,
      qty: qty,
      price: price,
      sumEquip: sumEquip,
      installUnit: installUnit,
      sumInstall: sumInstall,
      total: total,
      note: note,
      discountPct: discountPct
    });
  }

  return { items: items, jsonString: JSON.stringify(items) };
}

function buildPdfFileName_(meta) {
  var safe = function (s) {
    return String(s || '')
      .replace(/[\\\/:*?"<>|]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80);
  };

  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var kpNo = safe(meta.kpNo);
  var cust = safe(meta.customer);

  var name = 'КП';
  if (kpNo) name += '_' + kpNo;
  if (cust) name += '_' + cust;
  name += '_' + dateStr;

  return name;
}

/**
 * Формирует массив значений в порядке KP_EXPORT_CFG.LOG_HEADERS.
 */
function buildLogRow_(meta, cartJson, fileUrl, downloadUrl, fileId) {
  return [
    meta.timestamp,
    meta.kpNo,
    meta.kpDate || '',
    meta.manager,
    meta.phone,
    meta.customer,
    meta.customerAddr,
    meta.contractNo,
    meta.discountPct,
    meta.installPct,
    meta.equipTotal,
    meta.installTotal,
    meta.delivery,
    meta.toPay,
    meta.vat,
    meta.prepayPctNorm ? (meta.prepayPctNorm * 100) : '',
    meta.prepaySum,
    meta.mainLead,
    meta.ecoLead,
    meta.validDays,
    cartJson,
    fileUrl,
    downloadUrl || '',
    fileId
    // "Статус" не добавляем здесь: appendToLog_() заполнит автоматически
  ];
}

/**
 * Запись в "Журнал КП" ПО ЗАГОЛОВКАМ листа
 * + автозаполнение "Статус" = "Новая"
 */
function appendToLog_(ss, rowValues) {
  var cfgLogName =
    (typeof KP_EXPORT_CFG !== 'undefined' && KP_EXPORT_CFG.LOG_SHEET)
      ? KP_EXPORT_CFG.LOG_SHEET
      : 'Журнал КП';

  var cfgHeaders =
    (typeof KP_EXPORT_CFG !== 'undefined' && KP_EXPORT_CFG.LOG_HEADERS)
      ? KP_EXPORT_CFG.LOG_HEADERS
      : null;

  var log = ss.getSheetByName(cfgLogName);
  if (!log) {
    log = ss.insertSheet(cfgLogName);
    if (cfgHeaders && cfgHeaders.length) {
      log.getRange(1, 1, 1, cfgHeaders.length).setValues([cfgHeaders]);
      log.setFrozenRows(1);
    }
  }

  if (log.getLastRow() === 0 && cfgHeaders && cfgHeaders.length) {
    log.getRange(1, 1, 1, cfgHeaders.length).setValues([cfgHeaders]);
    log.setFrozenRows(1);
  }

  var lastCol = log.getLastColumn();
  if (lastCol < 1) throw new Error('Журнал КП: нет колонок (пустая шапка).');

  var sheetHeadersRaw = log.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var sheetHeadersNorm = sheetHeadersRaw.map(_normHeader_);

  // карта: нормЗаголовка -> индексКолонки (1-based)
  var colByNorm = {};
  for (var c = 0; c < sheetHeadersNorm.length; c++) {
    var h = sheetHeadersNorm[c];
    if (h && !(h in colByNorm)) colByNorm[h] = c + 1;
  }

  // значения: нормЗаголовка -> значение
  var valuesByNorm = {};

  if (Array.isArray(rowValues)) {
    var srcHeaders = (cfgHeaders && cfgHeaders.length) ? cfgHeaders : sheetHeadersRaw;
    for (var i = 0; i < rowValues.length; i++) {
      var hName = srcHeaders[i];
      if (!hName) continue;
      valuesByNorm[_normHeader_(hName)] = rowValues[i];
    }
  } else if (rowValues && typeof rowValues === 'object') {
    Object.keys(rowValues).forEach(function (k) {
      valuesByNorm[_normHeader_(k)] = rowValues[k];
    });
  }

  // собираем rowOut под текущую шапку листа
  var rowOut = sheetHeadersNorm.map(function (hn) {
    return (hn in valuesByNorm ? valuesByNorm[hn] : '');
  });

  var nextRow = Math.max(2, log.getLastRow() + 1);
  log.getRange(nextRow, 1, 1, rowOut.length).setValues([rowOut]);

  // ✅ Автостатус: "Новая" в колонке "Статус" (или fallback "Статус КП")
  try {
    var statusDefault = 'Новая';

    var cStatus =
      colByNorm[_normHeader_('Статус')] ||
      colByNorm[_normHeader_('Статус КП')];

    if (cStatus) {
      var cell = log.getRange(nextRow, cStatus);
      var cur = String(cell.getDisplayValue() || '').trim();
      if (!cur) cell.setValue(statusDefault);
    }
  } catch (e) {}

  // формат времени (по заголовку "Дата/время выгрузки" если есть, иначе колонка 1)
  try {
    var cTime = colByNorm[_normHeader_('Дата/время выгрузки')] || 1;
    log.getRange(nextRow, cTime).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  } catch (e) {}

  // ✅ автоформат денег
  try {
    applyMoneyFormatsInLog_(log);
  } catch (e) {}
}

/**
 * Ставит "#,##0.00" на денежные колонки по заголовкам.
 */
function applyMoneyFormatsInLog_(logSheet) {
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) return;

  var lastCol = logSheet.getLastColumn();
  if (lastCol < 1) return;

  var headers = logSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(_normHeader_);

  var moneyHeaders = [
    'Итого оборудование, руб',
    'Монтаж, руб',
    'Доставка, руб',
    'Итого к оплате, руб',
    'Сумма предоплаты, руб',
    'НДС 22%, руб'
  ].map(_normHeader_);

  var fmtMoney = '#,##0.00';

  for (var i = 0; i < moneyHeaders.length; i++) {
    var h = moneyHeaders[i];
    var idx = headers.indexOf(h);
    if (idx === -1) continue;
    var col = idx + 1;
    logSheet.getRange(2, col, lastRow - 1, 1).setNumberFormat(fmtMoney);
  }
}

/* ========================= Cart header helpers ========================= */

function _findCartHeaderRow_(sh, anchorText) {
  var anchor = _normHeader_(anchorText);
  var last = Math.min(800, sh.getLastRow());

  for (var r = 1; r <= last; r++) {
    var v = _normHeader_(sh.getRange(r, 1).getDisplayValue());
    if (v === anchor) return r;
  }
  return 0;
}

function _colByName_(hdrNorm, variants) {
  for (var i = 0; i < variants.length; i++) {
    var vn = _normHeader_(variants[i]);
    var idx = hdrNorm.indexOf(vn);
    if (idx >= 0) return idx + 1;
  }
  return 0;
}

function _colByContainsAll_(hdrNorm, parts) {
  var p = parts.map(_normHeader_);

  for (var i = 0; i < hdrNorm.length; i++) {
    var t = hdrNorm[i];
    if (!t) continue;

    var ok = true;
    for (var j = 0; j < p.length; j++) {
      if (t.indexOf(p[j]) === -1) {
        ok = false;
        break;
      }
    }
    if (ok) return i + 1;
  }
  return 0;
}