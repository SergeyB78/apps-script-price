/*************** KP_Export_PDF.gs ***************/
/**
 * Экспорт КП в PDF (A4, книжная), сохранение в Drive + запись в Журнал КП.
 * Блоки "Настройки расчёта" и "Условия и сроки поставки (изменяемые)" исключаются из PDF
 * (временно скрываем строки, затем возвращаем).
 *
 * ВАЛИДАЦИЯ: экспорт невозможен, если не заполнены:
 * - Наименование Заказчика
 * - Адрес Заказчика
 * - Менеджер
 * В сообщении показываем, какие именно поля пустые.
 */

const KP_EXPORT_CFG = {
  KP_SHEET: 'КП',
  LOG_SHEET: 'Журнал КП',

  DRIVE_FOLDER_ID: '1o8lqVv3DlUe4e3bMpWKvNZncf5r_2xDD',

  EXCLUDE_BLOCK_TITLES: {
    SETTINGS: 'Настройки расчёта',
    TERMS: 'Условия и сроки поставки (изменяемые)'
  },

  // сколько строк скрывать ПОСЛЕ заголовка (не включая строку заголовка)
  SETTINGS_ROWS_AFTER_TITLE: 3, // Скидка/Монтаж%/Доставка
  TERMS_ROWS_AFTER_TITLE: 4,    // 4 строки условий

  CART_HEADER: 'Артикул',

  LOG_HEADERS: [
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
    'Drive File ID'
  ]
};

/**
 * Меню вызывает эту функцию.
 */
function exportKpPdfAndLog() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(KP_EXPORT_CFG.KP_SHEET);
  if (!sh) throw new Error('Не найден лист "КП".');

  // ✅ ВАЛИДАЦИЯ обязательных полей (с перечислением пустых)
  const missing = validateRequiredKpFields_(sh);
  if (missing.length) {
    SpreadsheetApp.getUi().alert(
      'Экспорт КП невозможен.\n\nНе заполнены обязательные поля:\n- ' + missing.join('\n- ')
    );
    return;
  }

  const toHide = [];
  try {
    // --- исключаем блок "Настройки расчёта"
    const rSettings = findRowByExactText_(sh, KP_EXPORT_CFG.EXCLUDE_BLOCK_TITLES.SETTINGS);
    if (rSettings) {
      // если строка сверху пустая — тоже скрываем (ваша пустая строка-разделитель)
      const rPrev = rSettings - 1;
      if (rPrev >= 1 && isRowBlank_(sh, rPrev)) toHide.push(rPrev);

      toHide.push(rSettings); // заголовок
      toHide.push(...rangeRows_(rSettings + 1, rSettings + KP_EXPORT_CFG.SETTINGS_ROWS_AFTER_TITLE));
    }

    // --- исключаем блок "Условия и сроки поставки (изменяемые)"
    const rTerms = findRowByExactText_(sh, KP_EXPORT_CFG.EXCLUDE_BLOCK_TITLES.TERMS);
    if (rTerms) {
      toHide.push(rTerms); // заголовок
      toHide.push(...rangeRows_(rTerms + 1, rTerms + KP_EXPORT_CFG.TERMS_ROWS_AFTER_TITLE));
    }

    const uniqueHide = Array.from(new Set(toHide)).filter(r => r >= 1 && r <= sh.getMaxRows());
    if (uniqueHide.length) hideRowsBySortedList_(sh, uniqueHide);

    // meta + cart
    const meta = extractMetaForLog_(sh);
    const cart = extractCartAsJson_(sh);

    // pdf
    const pdfBlob = exportSheetToPdfBlob_(ss, sh, buildPdfFileName_(meta));
    const saved = savePdfToDriveFolder_(pdfBlob, KP_EXPORT_CFG.DRIVE_FOLDER_ID);

    // log
    const logRow = buildLogRow_(meta, cart.jsonString, saved.fileUrl, saved.downloadUrl, saved.fileId);
    appendToLog_(ss, logRow);

    // links
    showLinksDialog_(saved.fileUrl, saved.downloadUrl);

  } finally {
    // вернуть скрытые строки
    try {
      const uniqueHide = Array.from(new Set(toHide)).filter(r => r >= 1 && r <= sh.getMaxRows());
      if (uniqueHide.length) showRowsBySortedList_(sh, uniqueHide);
    } catch (e) {}
  }
}

/* ===================== VALIDATION ===================== */

function validateRequiredKpFields_(sh) {
  const requiredLabels = [
    'Наименование Заказчика',
    'Адрес Заказчика',
    'Менеджер'
  ];

  const missing = [];
  for (const label of requiredLabels) {
    const v = findValueByLabelInColD_(sh, label); // значение из колонки D по подписи в A (A:C merged)
    if (isEmptyValue_(v)) missing.push(label);
  }
  return missing;
}

function isEmptyValue_(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string') return v.trim() === '';
  const s = String(v).trim();
  return s === '';
}

/* ===================== PDF EXPORT ===================== */

function exportSheetToPdfBlob_(ss, sheet, fileName) {
  const ssId = ss.getId();
  const gid = sheet.getSheetId();

  const url =
    `https://docs.google.com/spreadsheets/d/${encodeURIComponent(ssId)}/export?` +
    [
      `format=pdf`,
      `gid=${gid}`,
      `size=A4`,
      `portrait=true`,
      `fitw=true`,
      `sheetnames=false`,
      `printtitle=false`,
      `pagenumbers=false`,
      `gridlines=false`,
      `fzr=false`,
      `top_margin=0.4`,
      `bottom_margin=0.4`,
      `left_margin=0.4`,
      `right_margin=0.4`
    ].join('&');

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error(`Не удалось сформировать PDF. Код: ${code}. Ответ: ${resp.getContentText()}`);
  }

  return resp.getBlob().setName(`${fileName}.pdf`);
}

function savePdfToDriveFolder_(blob, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  const fileId = file.getId();
  const fileUrl = file.getUrl();
  const downloadUrl = `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}`;

  return { fileId, fileUrl, downloadUrl };
}

/* ===================== LOG ===================== */

function appendToLog_(ss, rowValues) {
  let log = ss.getSheetByName(KP_EXPORT_CFG.LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(KP_EXPORT_CFG.LOG_SHEET);
    log.getRange(1, 1, 1, KP_EXPORT_CFG.LOG_HEADERS.length).setValues([KP_EXPORT_CFG.LOG_HEADERS]);
    log.setFrozenRows(1);
  } else {
    const firstRow = log.getRange(1, 1, 1, KP_EXPORT_CFG.LOG_HEADERS.length).getValues()[0];
    const hasHeader = String(firstRow[0] || '').trim() === KP_EXPORT_CFG.LOG_HEADERS[0];
    if (!hasHeader) {
      log.insertRowBefore(1);
      log.getRange(1, 1, 1, KP_EXPORT_CFG.LOG_HEADERS.length).setValues([KP_EXPORT_CFG.LOG_HEADERS]);
      log.setFrozenRows(1);
    }
  }

  const nextRow = log.getLastRow() + 1;
  log.getRange(nextRow, 1, 1, rowValues.length).setValues([rowValues]);

  // базовые форматы
  log.getRange(nextRow, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
}

/* ===================== META EXTRACTION ===================== */

function extractMetaForLog_(sh) {
  const getVal = (label) => findValueByLabelInColD_(sh, label);

  const kpNo = String(getVal('Коммерческое предложение №') || '').trim();
  const kpDate = getVal('Дата КП');
  const manager = String(getVal('Менеджер') || '').trim();
  const phone = String(getVal('Телефон') || '').trim();
  const customer = String(getVal('Наименование Заказчика') || '').trim();
  const customerAddr = String(getVal('Адрес Заказчика') || '').trim();
  const contractNo = String(getVal('№ Договора') || '').trim();

  // настройки расчёта
  const discountPct = toNumber_(getVal('Скидка (-) / Наценка (+), %'));
  const installPct = toNumber_(getVal('Размер монтажа от стоимости оборудования, %'));

  // итоги
  const equipTotal = toNumber_(findValueByLabelInColD_(sh, 'Итого за оборудование, руб'));
  const installTotal = toNumber_(findValueByLabelInColD_(sh, 'Монтаж, руб'));
  const delivery = toNumber_(findValueByLabelInColD_(sh, 'Доставка, руб'));
  const toPay = toNumber_(findValueByLabelInColD_(sh, 'Итого к оплате, руб'));
  const vat = toNumber_(findValueByLabelInColD_(sh, 'В том числе НДС 22%, руб'));

  // условия (верхний блок): слева текст, справа значение (J)
  const prepayPct = findTermsValueRight_(sh, 'Предоплата за оборудование составляет:');
  const prepayPctNorm = normalizePercent_(prepayPct); // 0..1
  const prepaySum = (isFiniteNumber_(equipTotal) ? equipTotal : 0) * (isFiniteNumber_(prepayPctNorm) ? prepayPctNorm : 0);

  const mainLead = findTermsValueRight_(sh, 'Срок поставки Основное производство исчисляется с момента поступления предоплаты на р/счет и составляет:');
  const ecoLead = findTermsValueRight_(sh, 'Срок поставки ЭКО-серия исчисляется с момента поступления предоплаты на р/счет и составляет:');
  const validDays = findTermsValueRight_(sh, 'Данное КП действительно в течение:');

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
    downloadUrl,
    fileId
  ];
}

function buildPdfFileName_(meta) {
  const safe = (s) => String(s || '')
    .replace(/[\\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const kpNo = safe(meta.kpNo || 'без_номера');
  const cust = safe(meta.customer || 'заказчик');
  return `КП_${kpNo}_${cust}_${dateStr}`;
}

/* ===================== CART -> JSON ===================== */

function extractCartAsJson_(sh) {
  const maxScan = Math.min(300, sh.getLastRow());
  let headerRow = 0;

  for (let r = 1; r <= maxScan; r++) {
    const v = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (v === KP_EXPORT_CFG.CART_HEADER) { headerRow = r; break; }
  }
  if (!headerRow) return { items: [], jsonString: '[]' };

  let r = headerRow + 1;
  const last = sh.getLastRow();
  const items = [];

  while (r <= last) {
    const art = String(sh.getRange(r, 1).getDisplayValue() || '').trim();  // A
    const name = String(sh.getRange(r, 4).getDisplayValue() || '').trim(); // D

    if (!art && !name) break;
    if (!art) { r++; continue; }

    const unit = String(sh.getRange(r, 5).getDisplayValue() || '').trim(); // E
    const price = toNumber_(sh.getRange(r, 6).getValue());     // F
    const qty = toNumber_(sh.getRange(r, 7).getValue());       // G
    const sumEquip = toNumber_(sh.getRange(r, 8).getValue());  // H
    const installUnit = toNumber_(sh.getRange(r, 9).getValue());  // I
    const sumInstall = toNumber_(sh.getRange(r, 10).getValue());  // J
    const total = toNumber_(sh.getRange(r, 11).getValue());       // K

    items.push({ art, name, unit, qty, price, sumEquip, installUnit, sumInstall, total });
    r++;
  }

  return { items, jsonString: JSON.stringify(items) };
}

/* ===================== HIDE/SHOW ROWS HELPERS ===================== */

function hideRowsBySortedList_(sh, rows) {
  const sorted = rows.slice().sort((a, b) => a - b);
  let start = sorted[0];
  let prev = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i];
    if (cur === prev + 1) { prev = cur; continue; }
    sh.hideRows(start, prev - start + 1);
    start = cur; prev = cur;
  }
  sh.hideRows(start, prev - start + 1);
}

function showRowsBySortedList_(sh, rows) {
  const sorted = rows.slice().sort((a, b) => a - b);
  let start = sorted[0];
  let prev = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i];
    if (cur === prev + 1) { prev = cur; continue; }
    sh.showRows(start, prev - start + 1);
    start = cur; prev = cur;
  }
  sh.showRows(start, prev - start + 1);
}

function rangeRows_(r1, r2) {
  const out = [];
  for (let r = r1; r <= r2; r++) out.push(r);
  return out;
}

/* ===================== FINDERS ===================== */

function findRowByExactText_(sh, text) {
  const target = String(text || '').trim();
  if (!target) return 0;

  const lastRow = sh.getLastRow();
  const lastCol = Math.min(11, sh.getLastColumn());
  const scanRows = Math.min(300, lastRow);

  const values = sh.getRange(1, 1, scanRows, lastCol).getDisplayValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (String(values[r][c] || '').trim() === target) return r + 1;
    }
  }
  return 0;
}

function isRowBlank_(sh, row) {
  const vals = sh.getRange(row, 1, 1, Math.min(11, sh.getLastColumn())).getDisplayValues()[0];
  return vals.every(v => String(v || '').trim() === '');
}

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

/* ===================== UI ===================== */

function showLinksDialog_(fileUrl, downloadUrl) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;font-size:14px;">
      <p>Скачать (попадёт в “Загрузки” браузера): <a href="${downloadUrl}" target="_blank">скачать PDF</a></p>
    </div>`
  ).setWidth(420).setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(html, 'КП → PDF');
}
