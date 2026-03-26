/**
 * kp_log.gs — сбор данных для журнала КП и запись в лист "Журнал КП"
 *
 * ИСПРАВЛЕНО:
 * - buildLogRow_() теперь возвращает OBJECT по именам полей, а не массив
 * - appendToLog_() пишет по заголовкам листа, с поддержкой алиасов заголовков
 * - убран риск смещения данных, если CFG.SCHEMAS.KP_LOG и реальная шапка листа разошлись
 * - ensureKpLogSchema_() гарантирует наличие полной схемы журнала и колонки "Статус"
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
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v).replace(/\s+/g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function _normalizePercentLoose_(v) {
  if (typeof normalizePercent_ === 'function') return normalizePercent_(v);
  const n = _toNumberLoose_(v);
  if (!n) return 0;
  return n > 1 ? n / 100 : n;
}

function _safeFindValueByLabelInColD_(sh, label) {
  if (typeof findValueByLabelInColD_ === 'function') return findValueByLabelInColD_(sh, label);

  // fallback: ищем подпись в A, значение в D
  const maxScan = Math.min(400, sh.getLastRow());
  for (let r = 1; r <= maxScan; r++) {
    const a = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (a === label) return sh.getRange(r, 4).getValue();
  }
  return '';
}

function _safeFindTermsValueRight_(sh, label) {
  if (typeof findTermsValueRight_ === 'function') return findTermsValueRight_(sh, label);

  // fallback: ищем подпись в A, значение в J
  const maxScan = Math.min(700, sh.getLastRow());
  for (let r = 1; r <= maxScan; r++) {
    const a = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (a === label) return sh.getRange(r, 10).getValue(); // J
  }
  return '';
}

function _findCartHeaderRow_(sh, anchorText) {
  const needle = _normHeader_(anchorText || 'Артикул');
  const last = Math.min(800, sh.getLastRow());
  for (let r = 1; r <= last; r++) {
    const v = _normHeader_(sh.getRange(r, 1).getDisplayValue());
    if (v === needle) return r;
  }
  return 0;
}

function _colByName_(hdrNorm, names) {
  const variants = (names || []).map(_normHeader_);
  for (let i = 0; i < hdrNorm.length; i++) {
    if (variants.indexOf(hdrNorm[i]) >= 0) return i + 1;
  }
  return 0;
}

function _colByContainsAll_(hdrNorm, parts) {
  const p = (parts || []).map(_normHeader_);
  for (let i = 0; i < hdrNorm.length; i++) {
    const t = hdrNorm[i] || '';
    if (p.every(x => t.indexOf(x) >= 0)) return i + 1;
  }
  return 0;
}

function _kpLogSheetName_() {
  return (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.KP_LOG)
    ? CFG.SHEETS.KP_LOG
    : 'Журнал КП';
}

/**
 * Каноническая схема журнала КП.
 * Даже если CFG устарел или в нём нет одного из заголовков,
 * используем полный безопасный fallback.
 */
function _getKpLogHeadersCanonical_() {
  const fallback = [
    'Дата/время выгрузки',
    'КП №',
    'Дата КП',
    'Менеджер',
    'Телефон',
    'Заказчик',
    'Адрес заказчика',
    '№ договора',
    'Скидка (-) / Наценка (+), %',
    'Размер монтажа от стоимости оборудования, %',
    'Итого оборудование, руб',
    'Монтаж, руб',
    'Доставка, руб',
    'Итого к оплате, руб',
    'НДС 22%, руб',
    'Предоплата, %',
    'Сумма предоплаты, руб',
    'Срок (Основное)',
    'Срок (ЭКО)',
    'КП действительно, дней',
    'Позиции (JSON)',
    'PDF URL (Drive)',
    'PDF Download URL',
    'Drive File ID',
    'Статус'
  ];

  let fromCfg = null;

  if (typeof CFG !== 'undefined' && CFG.SCHEMAS && Array.isArray(CFG.SCHEMAS.KP_LOG) && CFG.SCHEMAS.KP_LOG.length) {
    fromCfg = CFG.SCHEMAS.KP_LOG.slice();
  } else if (typeof CFG !== 'undefined' && CFG.KP_EXPORT && Array.isArray(CFG.KP_EXPORT.LOG_HEADERS) && CFG.KP_EXPORT.LOG_HEADERS.length) {
    fromCfg = CFG.KP_EXPORT.LOG_HEADERS.slice();
  } else if (typeof KP_EXPORT_CFG !== 'undefined' && Array.isArray(KP_EXPORT_CFG.LOG_HEADERS) && KP_EXPORT_CFG.LOG_HEADERS.length) {
    fromCfg = KP_EXPORT_CFG.LOG_HEADERS.slice();
  }

  if (!Array.isArray(fromCfg) || !fromCfg.length) return fallback.slice();

  const fromCfgNorm = fromCfg.map(_normHeader_);
  const fallbackNorm = fallback.map(_normHeader_);

  const missing = fallbackNorm.filter(h => fromCfgNorm.indexOf(h) < 0);
  if (missing.length) return fallback.slice();

  return fromCfg.slice();
}

/**
 * Алиасы заголовков -> канонический заголовок.
 * Нужны для совместимости со старыми версиями шапки листа.
 */
function _getKpLogHeaderAliases_() {
  const map = {};

  function addAlias(aliasHeader, canonicalHeader) {
    map[_normHeader_(aliasHeader)] = _normHeader_(canonicalHeader);
  }

  // Канонические заголовки сами на себя
  _getKpLogHeadersCanonical_().forEach(function (h) {
    map[_normHeader_(h)] = _normHeader_(h);
  });

  // Старые / альтернативные варианты
  addAlias('Итого за оборудование, руб', 'Итого оборудование, руб');
  addAlias('В том числе НДС 22%, руб', 'НДС 22%, руб');
  addAlias('PDF URL', 'PDF URL (Drive)');
  addAlias('PDF Download', 'PDF Download URL');
  addAlias('Адрес Заказчика', 'Адрес заказчика');
  addAlias('Срок поставки (Основное)', 'Срок (Основное)');
  addAlias('Срок поставки (ЭКО)', 'Срок (ЭКО)');
  addAlias('КП действительно', 'КП действительно, дней');

  return map;
}

function _canonicalNormHeader_(header) {
  const aliases = _getKpLogHeaderAliases_();
  const norm = _normHeader_(header);
  return aliases[norm] || norm;
}

function _getKpStatusList_() {
  if (typeof CFG !== 'undefined' && CFG.DROPDOWNS && Array.isArray(CFG.DROPDOWNS.KP_LOG_STATUS)) {
    return CFG.DROPDOWNS.KP_LOG_STATUS.slice();
  }
  return ['Новая', 'Аннулирована', 'Отправлена в производство'];
}

/* ========================= API used by controller ========================= */

/**
 * Мета-данные для журнала.
 */
function extractMetaForLog_(sh) {
  const kpNo = String(_safeFindValueByLabelInColD_(sh, 'Коммерческое предложение №') || '').trim();
  const kpDate = _safeFindValueByLabelInColD_(sh, 'Дата КП');
  const manager = String(_safeFindValueByLabelInColD_(sh, 'Менеджер') || '').trim();
  const phone = String(_safeFindValueByLabelInColD_(sh, 'Телефон') || '').trim();
  const customer = String(_safeFindValueByLabelInColD_(sh, 'Наименование Заказчика') || '').trim();
  const customerAddr = String(_safeFindValueByLabelInColD_(sh, 'Адрес Заказчика') || '').trim();
  const contractNo = String(_safeFindValueByLabelInColD_(sh, '№ Договора') || '').trim();

  const discountPct = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Скидка (-) / Наценка (+), %'));
  const installPct = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Размер монтажа от стоимости оборудования, %'));

  let equipTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого за оборудование, руб'));
  if (!equipTotal) equipTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого оборудование, руб'));

  const installTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Монтаж, руб'));
  const delivery = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Доставка, руб'));
  const toPay = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого к оплате, руб'));

  let vat = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'НДС 22%, руб'));
  if (!vat) vat = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'В том числе НДС 22%, руб'));

  const prepayPctNorm = _normalizePercentLoose_(
    _safeFindTermsValueRight_(sh, 'Предоплата за оборудование составляет:')
  );

  let prepaySum = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Сумма предоплаты, руб'));
  if (!prepaySum && equipTotal && prepayPctNorm) prepaySum = equipTotal * prepayPctNorm;

  const mainLead =
    String(_safeFindTermsValueRight_(sh, 'Срок поставки Основное производство исчисляется с момента поступления предоплаты на р/счет и составляет:') || '').trim()
    || String(_safeFindValueByLabelInColD_(sh, 'Срок (Основное)') || '').trim();

  const ecoLead =
    String(_safeFindTermsValueRight_(sh, 'Срок поставки ЭКО-серия исчисляется с момента поступления предоплаты на р/счет и составляет:') || '').trim()
    || String(_safeFindValueByLabelInColD_(sh, 'Срок (ЭКО)') || '').trim();

  const validDays =
    _toNumberLoose_(_safeFindTermsValueRight_(sh, 'Данное КП действительно в течение:'))
    || _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'КП действительно, дней'));

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
 * Корзина КП -> JSON
 * Пропускает строки без UID (если колонка UID есть), чтобы не тащить строки групп.
 * Пропускает строки с qty <= 0.
 */
function extractCartAsJson_(sh) {
  const headerRow = _findCartHeaderRow_(sh, 'Артикул');
  if (!headerRow) return { items: [], jsonString: '[]' };

  const maxCol = Math.min(50, sh.getLastColumn());
  const hdr = sh.getRange(headerRow, 1, 1, maxCol).getDisplayValues()[0].map(_normHeader_);

  const colArt = _colByName_(hdr, ['артикул']) || 1;
  const colName = _colByContainsAll_(hdr, ['наименование']) || 4;
  const colUnit = _colByName_(hdr, ['ед. изм', 'ед.изм.', 'ед.']) || 5;
  const colPrice = _colByContainsAll_(hdr, ['стоимость', 'оборуд']) || 6;
  const colQty = _colByContainsAll_(hdr, ['кол']) || 7;
  const colSumEquip = _colByContainsAll_(hdr, ['всего', 'оборуд']) || 8;
  const colInstallUnit = _colByContainsAll_(hdr, ['стоимость', 'монтаж']) || 9;
  const colSumInstall = _colByContainsAll_(hdr, ['всего', 'монтаж']) || 10;
  const colTotal = _colByName_(hdr, ['итого']) || 11;
  const colNote = _colByName_(hdr, ['примечание', 'комментарий']) || 0;
  const colDiscount = _colByContainsAll_(hdr, ['скидка', '%']) || 0;
  const colUid = _colByName_(hdr, ['uid']) || 0;

  const items = [];
  const lastRow = sh.getLastRow();

  for (let r = headerRow + 1; r <= lastRow; r++) {
    const art = String(sh.getRange(r, colArt).getDisplayValue() || '').trim();
    const name = String(sh.getRange(r, colName).getDisplayValue() || '').trim();

    // пустой хвост
    if (!art && !name) break;

    // qty > 0 обязательно
    const qty = _toNumberLoose_(sh.getRange(r, colQty).getValue());
    if (!(qty > 0)) continue;

    // Если есть UID-колонка — строки без UID считаем служебными/группами и пропускаем
    const uid = colUid ? String(sh.getRange(r, colUid).getDisplayValue() || '').trim() : '';
    if (colUid && !uid) continue;

    // Артикул/наименование
    if (!art && !name) continue;

    const unit = String(sh.getRange(r, colUnit).getDisplayValue() || '').trim();
    const price = _toNumberLoose_(sh.getRange(r, colPrice).getValue());
    const sumEquip = _toNumberLoose_(sh.getRange(r, colSumEquip).getValue());
    const installUnit = _toNumberLoose_(sh.getRange(r, colInstallUnit).getValue());
    const sumInstall = _toNumberLoose_(sh.getRange(r, colSumInstall).getValue());
    const total = _toNumberLoose_(sh.getRange(r, colTotal).getValue());
    const note = colNote ? String(sh.getRange(r, colNote).getDisplayValue() || '').trim() : '';
    const discountPct = colDiscount ? _toNumberLoose_(sh.getRange(r, colDiscount).getValue()) : '';

    items.push({
      uid: uid || '',
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
  const safe = function (s) {
    return String(s || '')
      .replace(/[\\\/:*?"<>|]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80);
  };

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const kpNo = safe(meta.kpNo);
  const cust = safe(meta.customer);

  let name = 'КП';
  if (kpNo) name += '_' + kpNo;
  if (cust) name += '_' + cust;
  name += '_' + dateStr;

  return name;
}

/**
 * ВАЖНО:
 * Возвращаем ОБЪЕКТ по именам полей, а не массив.
 * Это убирает риск смещения, если схема заголовков листа и конфига разошлась.
 */
function buildLogRow_(meta, cartJson, fileUrl, downloadUrl, fileId) {
  return {
    'Дата/время выгрузки': meta.timestamp,
    'КП №': meta.kpNo,
    'Дата КП': meta.kpDate || '',
    'Менеджер': meta.manager,
    'Телефон': meta.phone,
    'Заказчик': meta.customer,
    'Адрес заказчика': meta.customerAddr,
    '№ договора': meta.contractNo,
    'Скидка (-) / Наценка (+), %': meta.discountPct,
    'Размер монтажа от стоимости оборудования, %': meta.installPct,
    'Итого оборудование, руб': meta.equipTotal,
    'Монтаж, руб': meta.installTotal,
    'Доставка, руб': meta.delivery,
    'Итого к оплате, руб': meta.toPay,
    'НДС 22%, руб': meta.vat,
    'Предоплата, %': meta.prepayPctNorm ? (meta.prepayPctNorm * 100) : '',
    'Сумма предоплаты, руб': meta.prepaySum,
    'Срок (Основное)': meta.mainLead,
    'Срок (ЭКО)': meta.ecoLead,
    'КП действительно, дней': meta.validDays,
    'Позиции (JSON)': cartJson,
    'PDF URL (Drive)': fileUrl,
    'PDF Download URL': downloadUrl || '',
    'Drive File ID': fileId || '',
    'Статус': 'Новая'
  };
}

/**
 * КРИТИЧЕСКОЕ:
 * Пишем строку ПО ЗАГОЛОВКАМ листа, а не "слева направо".
 * Поддерживается и object-based, и array-based режим.
 */
function appendToLog_(ss, rowArrayOrObject) {
  const sh = ensureKpLogSchema_(ss);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const normHeaders = headers.map(_normHeader_);

  const aliases = _getKpLogHeaderAliases_();
  const canonicalHeaders = _getKpLogHeadersCanonical_();

  // Нормализованный словарь значений: canonicalNormHeader -> value
  const valueMap = {};

  if (Array.isArray(rowArrayOrObject)) {
    // fallback для старых вызовов, если где-то ещё остался массив
    const srcHeaders = canonicalHeaders;
    for (let i = 0; i < srcHeaders.length; i++) {
      const hNorm = _canonicalNormHeader_(srcHeaders[i]);
      valueMap[hNorm] = (i < rowArrayOrObject.length) ? rowArrayOrObject[i] : '';
    }
  } else {
    const obj = rowArrayOrObject || {};
    Object.keys(obj).forEach(function (k) {
      const hNorm = _canonicalNormHeader_(k);
      valueMap[hNorm] = obj[k];
    });
  }

  // Гарантия статуса по умолчанию
  const statusKey = _canonicalNormHeader_('Статус');
  if (!String(valueMap[statusKey] || '').trim()) {
    valueMap[statusKey] = 'Новая';
  }

  // Собираем строку в реальном порядке заголовков листа
  const row = normHeaders.map(function (sheetHeaderNorm) {
    const canonicalSheetHeaderNorm = aliases[sheetHeaderNorm] || sheetHeaderNorm;
    return (canonicalSheetHeaderNorm in valueMap) ? valueMap[canonicalSheetHeaderNorm] : '';
  });

  sh.appendRow(row);
  const rowNum = sh.getLastRow();

  // Формат денег (по именам заголовков)
  const moneyNames = [
    'Итого оборудование, руб',
    'Монтаж, руб',
    'Доставка, руб',
    'Итого к оплате, руб',
    'НДС 22%, руб',
    'Сумма предоплаты, руб'
  ].map(_canonicalNormHeader_);

  for (let c = 1; c <= normHeaders.length; c++) {
    const canonicalHeader = aliases[normHeaders[c - 1]] || normHeaders[c - 1];
    if (moneyNames.indexOf(canonicalHeader) >= 0) {
      try {
        sh.getRange(rowNum, c).setNumberFormat('#,##0.00');
      } catch (e) {}
    }
  }

  // Формат даты/времени
  try {
    const timeCol = normHeaders.indexOf(_normHeader_('Дата/время выгрузки')) + 1;
    if (timeCol > 0) sh.getRange(rowNum, timeCol).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  } catch (e) {}

  // Валидация статуса на новой строке
  ensureKpLogStatusValidation_(sh);

  return rowNum;
}

/* ========================= Schema / Headers / Validation ========================= */

function ensureKpLogSchema_(ss) {
  const shName = _kpLogSheetName_();
  let sh = ss.getSheetByName(shName);
  if (!sh) sh = ss.insertSheet(shName);

  const needHeaders = _getKpLogHeadersCanonical_();
  if (!needHeaders.length) throw new Error('Не определены заголовки для "Журнал КП".');

  const hasHeaderRow = sh.getLastRow() >= 1;
  const currentHeaders = hasHeaderRow
    ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0]
    : [];

  const currentNorm = currentHeaders.map(_canonicalNormHeader_);
  const needNorm = needHeaders.map(_canonicalNormHeader_);

  // Если шапки нет вообще — пишем целиком
  if (!hasHeaderRow || currentHeaders.join('').trim() === '') {
    sh.getRange(1, 1, 1, needHeaders.length).setValues([needHeaders]);
    sh.setFrozenRows(1);
  } else {
    // Добавляем недостающие колонки справа
    needHeaders.forEach(function (h, i) {
      if (currentNorm.indexOf(needNorm[i]) < 0) {
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1, sh.getLastColumn()).setValue(h);
      }
    });

    // Если заголовок совпадает по смыслу, но написан по-старому — нормализуем название
    const finalHeadersTmp = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const finalNormTmp = finalHeadersTmp.map(_canonicalNormHeader_);

    for (let i = 0; i < finalHeadersTmp.length; i++) {
      const hNorm = finalNormTmp[i];
      const canonicalIdx = needNorm.indexOf(hNorm);
      if (canonicalIdx >= 0) {
        const canonicalTitle = needHeaders[canonicalIdx];
        if (String(finalHeadersTmp[i] || '').trim() !== canonicalTitle) {
          sh.getRange(1, i + 1).setValue(canonicalTitle);
        }
      }
    }
  }

  ensureKpLogStatusValidation_(sh);
  return sh;
}

function ensureKpLogStatusValidation_(sh) {
  if (!sh) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(_canonicalNormHeader_);
  const statusCol = headers.indexOf(_canonicalNormHeader_('Статус')) + 1;
  if (!statusCol) return;

  const statuses = _getKpStatusList_();
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statuses, true)
    .setAllowInvalid(false)
    .build();

  // На все строки ниже шапки (с запасом)
  const startRow = 2;
  const numRows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(startRow, statusCol, numRows, 1).setDataValidation(rule);

  // Пустым существующим статусам — "Новая"
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const rng = sh.getRange(2, statusCol, lastRow - 1, 1);
    const vals = rng.getValues();
    let changed = false;

    for (let i = 0; i < vals.length; i++) {
      if (!String(vals[i][0] || '').trim()) {
        vals[i][0] = 'Новая';
        changed = true;
      }
    }

    if (changed) rng.setValues(vals);
  }
}