/**
 * kp_log.gs — сбор данных для журнала КП и запись в лист "Журнал КП"
 *
 * Основное:
 * - appendToLog_() пишет ПО ЗАГОЛОВКАМ, а не "слева направо"
 * - ensureKpLogSchema_() гарантирует наличие колонки "Статус"
 * - Для новых записей статус = "Новая" + выпадающий список
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
  return (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.KP_LOG) ? CFG.SHEETS.KP_LOG : 'Журнал КП';
}

function _getKpLogHeadersCfg_() {
  // Приоритет: CFG.SCHEMAS.KP_LOG
  if (typeof CFG !== 'undefined' && CFG.SCHEMAS && Array.isArray(CFG.SCHEMAS.KP_LOG) && CFG.SCHEMAS.KP_LOG.length) {
    return CFG.SCHEMAS.KP_LOG.slice();
  }

  // fallback на старый KP_EXPORT_CFG.LOG_HEADERS
  let headers = [];
  if (typeof CFG !== 'undefined' && CFG.KP_EXPORT && Array.isArray(CFG.KP_EXPORT.LOG_HEADERS)) {
    headers = CFG.KP_EXPORT.LOG_HEADERS.slice();
  } else if (typeof KP_EXPORT_CFG !== 'undefined' && Array.isArray(KP_EXPORT_CFG.LOG_HEADERS)) {
    headers = KP_EXPORT_CFG.LOG_HEADERS.slice();
  }

  // гарантируем Статус
  if (headers.indexOf('Статус') < 0) headers.push('Статус');
  return headers;
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
    kpNo,
    kpDate,
    manager,
    phone,
    customer,
    customerAddr,
    contractNo,
    discountPct,
    installPct,
    equipTotal,
    installTotal,
    delivery,
    toPay,
    vat,
    prepayPctNorm,
    prepaySum,
    mainLead,
    ecoLead,
    validDays
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
      art,
      name,
      unit,
      qty,
      price,
      sumEquip,
      installUnit,
      sumInstall,
      total,
      note,
      discountPct
    });
  }

  return { items, jsonString: JSON.stringify(items) };
}

function buildPdfFileName_(meta) {
  const safe = (s) => String(s || '')
    .replace(/[\\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);

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
 * Массив значений в порядке заголовков журнала (без жёсткой привязки к позициям колонок).
 */
function buildLogRow_(meta, cartJson, fileUrl, downloadUrl, fileId) {
  return [
    meta.timestamp,                                    // Дата/время выгрузки
    meta.kpNo,                                         // КП №
    meta.kpDate || '',                                 // Дата КП
    meta.manager,                                      // Менеджер
    meta.phone,                                        // Телефон
    meta.customer,                                     // Заказчик
    meta.customerAddr,                                 // Адрес заказчика
    meta.contractNo,                                   // № договора
    meta.discountPct,                                  // Скидка %
    meta.installPct,                                   // Монтаж %
    meta.equipTotal,                                   // Итого оборудование
    meta.installTotal,                                 // Монтаж
    meta.delivery,                                     // Доставка
    meta.toPay,                                        // Итого к оплате
    meta.vat,                                          // НДС
    meta.prepayPctNorm ? (meta.prepayPctNorm * 100) : '', // Предоплата %
    meta.prepaySum,                                    // Сумма предоплаты
    meta.mainLead,                                     // Срок (Основное)
    meta.ecoLead,                                      // Срок (ЭКО)
    meta.validDays,                                    // КП действительно
    cartJson,                                          // Позиции (JSON)
    fileUrl,                                           // PDF URL
    downloadUrl || '',                                 // PDF Download URL
    fileId || '',                                      // Drive File ID
    'Новая'                                            // Статус (новый)
  ];
}

/**
 * КРИТИЧЕСКОЕ:
 * Пишем строку ПО ЗАГОЛОВКАМ листа, чтобы значения не съезжали.
 */
function appendToLog_(ss, rowArrayOrObject) {
  const sh = ensureKpLogSchema_(ss);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const normHeaders = headers.map(_normHeader_);

  // Конфигурируемые заголовки "эталона"
  const cfgHeaders = _getKpLogHeadersCfg_();

  // Нормализованный словарь значений
  const valueMap = {};

  if (Array.isArray(rowArrayOrObject)) {
    // Маппим массив по cfgHeaders
    for (let i = 0; i < cfgHeaders.length; i++) {
      valueMap[_normHeader_(cfgHeaders[i])] = (i < rowArrayOrObject.length) ? rowArrayOrObject[i] : '';
    }
  } else {
    const obj = rowArrayOrObject || {};
    Object.keys(obj).forEach(k => valueMap[_normHeader_(k)] = obj[k]);
  }

  // Гарантия статуса по умолчанию
  const statusKey = _normHeader_('Статус');
  if (!String(valueMap[statusKey] || '').trim()) {
    valueMap[statusKey] = 'Новая';
  }

  // Собираем строку в реальном порядке заголовков листа
  const row = normHeaders.map(h => (h in valueMap) ? valueMap[h] : '');

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
  ].map(_normHeader_);

  for (let c = 1; c <= normHeaders.length; c++) {
    if (moneyNames.indexOf(normHeaders[c - 1]) >= 0) {
      try { sh.getRange(rowNum, c).setNumberFormat('#,##0.00'); } catch (e) {}
    }
  }

  // Валидация статуса на новой строке
  ensureKpLogStatusValidation_(sh);

  return rowNum;
}

/* ========================= Schema / Headers / Validation ========================= */

function ensureKpLogSchema_(ss) {
  const shName = _kpLogSheetName_();
  let sh = ss.getSheetByName(shName);
  if (!sh) sh = ss.insertSheet(shName);

  const needHeaders = _getKpLogHeadersCfg_();
  if (!needHeaders.length) throw new Error('Не определены заголовки для "Журнал КП".');

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const currentHeaders = sh.getLastRow() >= 1
    ? sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    : [];

  const currentNorm = currentHeaders.map(_normHeader_);
  const needNorm = needHeaders.map(_normHeader_);

  // Если шапки нет вообще — пишем целиком
  if (currentHeaders.join('').trim() === '') {
    sh.getRange(1, 1, 1, needHeaders.length).setValues([needHeaders]);
  } else {
    // Добавляем недостающие колонки справа
    needHeaders.forEach((h, i) => {
      if (currentNorm.indexOf(needNorm[i]) < 0) {
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1, sh.getLastColumn()).setValue(h);
      }
    });
  }

  // Повторно читаем шапку
  const finalHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const finalNorm = finalHeaders.map(_normHeader_);

  // Приведём названия (если совпали по смыслу, но пусто/криво) — не трогаем порядок, только гарантируем "Статус"
  const statusIdx = finalNorm.indexOf(_normHeader_('Статус')) + 1;
  if (!statusIdx) {
    sh.insertColumnAfter(sh.getLastColumn());
    sh.getRange(1, sh.getLastColumn()).setValue('Статус');
  }

  ensureKpLogStatusValidation_(sh);
  return sh;
}

function ensureKpLogStatusValidation_(sh) {
  if (!sh) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(_normHeader_);
  const statusCol = headers.indexOf(_normHeader_('Статус')) + 1;
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