/**
 * kp_log.gs — сбор данных для журнала КП и запись в лист "Журнал КП"
 *
 * КЛЮЧЕВОЕ:
 * - appendToLog_() записывает строку ПО ЗАГОЛОВКАМ листа, чтобы значения не "съезжали"
 *   при удалении/перестановке/скрытии столбцов (например Drive File ID).
 * - Автоформат денежных колонок: "#,##0.00"
 *
 * Ожидается, что в проекте есть (или будут) файлы:
 * - КП pdf.gs (контроллер), где определён KP_EXPORT_CFG.LOG_HEADERS
 * - kp_pdf_export.gs (экспорт/Drive/UI)
 * - kp_utils.gs (утилиты) — опционально
 */

/* =========================
   Helpers / Safe wrappers
   ========================= */

function _normHeader_(s) {
  return String(s || '')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _toNumberLoose_(v) {
  // если в проекте есть toNumber_ — используем
  if (typeof toNumber_ === 'function') return toNumber_(v);

  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/\s+/g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function _normalizePercentLoose_(v) {
  // если в проекте есть normalizePercent_ — используем
  if (typeof normalizePercent_ === 'function') return normalizePercent_(v);

  const n = _toNumberLoose_(v);
  if (!n) return 0;
  // 70 -> 0.7, 0.7 -> 0.7
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

/* =========================
   API used by controller
   ========================= */

/**
 * Мета-данные для журнала.
 * (Логика близкая к “базовой” — при необходимости можно точечно подстроить подписи.)
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

  // оборудование
  let equipTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого за оборудование, руб'));
  if (!equipTotal) equipTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого оборудование, руб'));

  const installTotal = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Монтаж, руб'));
  const delivery = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Доставка, руб'));
  const toPay = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Итого к оплате, руб'));

  // НДС
  let vat = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'НДС 22%, руб'));
  if (!vat) vat = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'В том числе НДС 22%, руб'));

  // предоплата % (обычно в блоке условий)
  const prepayPctNorm = _normalizePercentLoose_(
    _safeFindTermsValueRight_(sh, 'Предоплата за оборудование составляет:')
  );

  // сумма предоплаты — если явной строки нет, считаем как в КП.gs: от оборудования
  let prepaySum = _toNumberLoose_(_safeFindValueByLabelInColD_(sh, 'Сумма предоплаты, руб'));
  if (!prepaySum && equipTotal && prepayPctNorm) prepaySum = equipTotal * prepayPctNorm;

  // сроки
  const mainLead =
    String(_safeFindTermsValueRight_(sh, 'Срок поставки Основное производство исчисляется с момента поступления предоплаты на р/счет и составляет:') || '').trim()
    || String(_safeFindValueByLabelInColD_(sh, 'Срок (Основное)') || '').trim();

  const ecoLead =
    String(_safeFindTermsValueRight_(sh, 'Срок поставки ЭКО-серия исчисляется с момента поступления предоплаты на р/счет и составляет:') || '').trim()
    || String(_safeFindValueByLabelInColD_(sh, 'Срок (ЭКО)') || '').trim();

  // действительность КП
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
 * Корзина → JSON (включая Примечание и индивидуальную скидку, если такие колонки есть).
 * Это нужно для сравнения "Примечаний" при проверке дублей.
 */
function extractCartAsJson_(sh) {
  const headerRow = _findCartHeaderRow_(sh, 'Артикул');
  if (!headerRow) return { items: [], jsonString: '[]' };

  const maxCol = Math.min(40, sh.getLastColumn());
  const hdr = sh.getRange(headerRow, 1, 1, maxCol).getDisplayValues()[0].map(_normHeader_);

  const colArt = _colByName_(hdr, ['артикул']) || 1;
  const colName = _colByName_(hdr, ['наименование']) || 4;
  const colUnit = _colByName_(hdr, ['ед. изм.', 'ед.изм.', 'ед.']) || 5;
  const colPrice = _colByName_(hdr, ['стоимость оборудования', 'цена', 'стоимость']) || 6;
  const colQty = _colByName_(hdr, ['кол-во', 'количество']) || 7;
  const colSumEquip = _colByName_(hdr, ['всего за оборудование', 'сумма за оборудование']) || 8;
  const colInstallUnit = _colByName_(hdr, ['стоимость монтажа']) || 9;
  const colSumInstall = _colByName_(hdr, ['всего за монтаж']) || 10;
  const colTotal = _colByName_(hdr, ['итого']) || 11;

  const colNote = _colByName_(hdr, ['примечание', 'комментарий']) || 0;
  const colDiscount = _colByContainsAll_(hdr, ['скидка', '%']) || 0;

  const items = [];
  const lastRow = sh.getLastRow();

  for (let r = headerRow + 1; r <= lastRow; r++) {
    const art = String(sh.getRange(r, colArt).getDisplayValue() || '').trim();
    const name = String(sh.getRange(r, colName).getDisplayValue() || '').trim();

    if (!art && !name) break;
    if (!art) continue;

    const unit = String(sh.getRange(r, colUnit).getDisplayValue() || '').trim();
    const price = _toNumberLoose_(sh.getRange(r, colPrice).getValue());
    const qty = _toNumberLoose_(sh.getRange(r, colQty).getValue());
    const sumEquip = _toNumberLoose_(sh.getRange(r, colSumEquip).getValue());
    const installUnit = _toNumberLoose_(sh.getRange(r, colInstallUnit).getValue());
    const sumInstall = _toNumberLoose_(sh.getRange(r, colSumInstall).getValue());
    const total = _toNumberLoose_(sh.getRange(r, colTotal).getValue());

    const note = colNote ? String(sh.getRange(r, colNote).getDisplayValue() || '').trim() : '';
    const discountPct = colDiscount ? _toNumberLoose_(sh.getRange(r, colDiscount).getValue()) : '';

    items.push({ art, name, unit, qty, price, sumEquip, installUnit, sumInstall, total, note, discountPct });
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
 * Формирует массив значений в порядке KP_EXPORT_CFG.LOG_HEADERS (как “базовая” логика).
 * ВАЖНО: даже если на листе "Журнал КП" столбцы переставлены — appendToLog_ положит всё по заголовкам.
 */
function buildLogRow_(meta, cartJson, fileUrl, downloadUrl, fileId) {
  // downloadUrl можно передавать '' если колонку убрали/не используете
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
  ];
}

/**
 * ✅ КРИТИЧЕСКАЯ ПРАВКА:
 * Записывает строку в "Журнал КП" ПО ЗАГОЛОВКАМ листа, а не "слева направо",
 * поэтому Drive File ID не будет съезжать в другую колонку.
 */
function appendToLog_(ss, rowValues) {
  const cfgLogName = (typeof KP_EXPORT_CFG !== 'undefined' && KP_EXPORT_CFG.LOG_SHEET) ? KP_EXPORT_CFG.LOG_SHEET : 'Журнал КП';
  const cfgHeaders = (typeof KP_EXPORT_CFG !== 'undefined' && KP_EXPORT_CFG.LOG_HEADERS) ? KP_EXPORT_CFG.LOG_HEADERS : null;

  let log = ss.getSheetByName(cfgLogName);
  if (!log) {
    log = ss.insertSheet(cfgLogName);
    if (cfgHeaders && cfgHeaders.length) {
      log.getRange(1, 1, 1, cfgHeaders.length).setValues([cfgHeaders]);
      log.setFrozenRows(1);
    }
  }

  // если лист пустой (на всякий)
  if (log.getLastRow() === 0 && cfgHeaders && cfgHeaders.length) {
    log.getRange(1, 1, 1, cfgHeaders.length).setValues([cfgHeaders]);
    log.setFrozenRows(1);
  }

  const lastCol = log.getLastColumn();
  if (lastCol < 1) throw new Error('Журнал КП: нет колонок (пустая шапка).');

  const sheetHeadersRaw = log.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const sheetHeadersNorm = sheetHeadersRaw.map(_normHeader_);

  // карта: нормЗаголовка -> индексКолонки (1-based)
  const colByNorm = {};
  for (let c = 0; c < sheetHeadersNorm.length; c++) {
    const h = sheetHeadersNorm[c];
    if (h && !(h in colByNorm)) colByNorm[h] = c + 1;
  }

  // значения: нормЗаголовка -> значение
  const valuesByNorm = {};

  if (Array.isArray(rowValues)) {
    // если контроллер формирует массив в порядке KP_EXPORT_CFG.LOG_HEADERS — мапим по этим заголовкам
    const srcHeaders = (cfgHeaders && cfgHeaders.length) ? cfgHeaders : sheetHeadersRaw;

    for (let i = 0; i < rowValues.length; i++) {
      const hName = srcHeaders[i];
      if (!hName) continue;
      valuesByNorm[_normHeader_(hName)] = rowValues[i];
    }
  } else if (rowValues && typeof rowValues === 'object') {
    // если вдруг когда-то перейдёте на object-based
    Object.keys(rowValues).forEach(k => {
      valuesByNorm[_normHeader_(k)] = rowValues[k];
    });
  }

  // собираем rowOut под текущую шапку листа
  const rowOut = sheetHeadersNorm.map(hn => (hn in valuesByNorm ? valuesByNorm[hn] : ''));

  const nextRow = Math.max(2, log.getLastRow() + 1);
  log.getRange(nextRow, 1, 1, rowOut.length).setValues([rowOut]);

  // формат времени (по заголовку "Дата/время выгрузки" если есть, иначе колонка 1)
  try {
    const cTime = colByNorm[_normHeader_('Дата/время выгрузки')] || 1;
    log.getRange(nextRow, cTime).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  } catch (e) {}

  // ✅ автоформат денег (на весь диапазон данных)
  try { applyMoneyFormatsInLog_(log); } catch (e) {}
}

/**
 * Ставит "#,##0.00" на денежные колонки по заголовкам.
 * Применяется со 2-й строки до последней.
 */
function applyMoneyFormatsInLog_(logSheet) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = logSheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = logSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(_normHeader_);

  const moneyHeaders = [
    'Итого оборудование, руб',
    'Монтаж, руб',
    'Доставка, руб',
    'Итого к оплате, руб',
    'Сумма предоплаты, руб',
    'НДС 22%, руб'
  ].map(_normHeader_);

  const fmtMoney = '#,##0.00';

  for (const h of moneyHeaders) {
    const idx = headers.indexOf(h);
    if (idx === -1) continue;
    const col = idx + 1;
    logSheet.getRange(2, col, lastRow - 1, 1).setNumberFormat(fmtMoney);
  }
}

/* =========================
   Cart header helpers
   ========================= */

function _findCartHeaderRow_(sh, anchorText) {
  const anchor = _normHeader_(anchorText);
  const last = Math.min(800, sh.getLastRow());
  for (let r = 1; r <= last; r++) {
    const v = _normHeader_(sh.getRange(r, 1).getDisplayValue());
    if (v === anchor) return r;
  }
  return 0;
}

function _colByName_(hdrNorm, variants) {
  for (const v of variants) {
    const vn = _normHeader_(v);
    const idx = hdrNorm.indexOf(vn);
    if (idx >= 0) return idx + 1;
  }
  return 0;
}

function _colByContainsAll_(hdrNorm, parts) {
  const p = parts.map(_normHeader_);
  for (let i = 0; i < hdrNorm.length; i++) {
    const t = hdrNorm[i];
    if (!t) continue;
    if (p.every(x => t.includes(x))) return i + 1;
  }
  return 0;
}