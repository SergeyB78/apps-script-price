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

  DRIVE_FOLDER_ID: ((typeof CFG !== 'undefined' && CFG.IDS && CFG.IDS.DRIVE_FOLDER_ID) ? CFG.IDS.DRIVE_FOLDER_ID : '1o8lqVv3DlUe4e3bMpWKvNZncf5r_2xDD'),

  EXCLUDE_BLOCK_TITLES: {
    SETTINGS: 'Настройки расчёта',
    TERMS: 'Условия и сроки поставки (изменяемые)'
  },

  // сколько строк скрывать ПОСЛЕ заголовка (не включая строку заголовка)
  SETTINGS_ROWS_AFTER_TITLE: 3, // Скидка/Монтаж%/Доставка
  TERMS_ROWS_AFTER_TITLE: 4,    // 4 строки условий

  CART_HEADER: 'Артикул',


  // Колонки корзины, которые НЕ должны попадать в PDF (временно скрываем перед экспортом)
  CART_EXCLUDE_COL_TITLES: ['Скидка (-) / Наценка (+), %'],
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

// ===================== ЖУРНАЛ КП: какие колонки оставляем видимыми =====================
// ВАЖНО: группировка делается ПО НАЗВАНИЯМ ЗАГОЛОВКОВ (не по цветам).
// Если хотите изменить "важные" колонки — правьте список VISIBLE_HEADERS.
const JOURNAL_LAYOUT = {
  HEADER_ROW: 1,

  // Эти колонки остаются открытыми (видимыми) — остальные будут свернуты группами
  VISIBLE_HEADERS: [
    'Дата/время выгрузки',
    'КП №',
    'Дата КП',
    'Менеджер',
    'Заказчик',
    'Адрес заказчика',

    'Итого оборудование, руб',
    'Монтаж, руб',
    'Доставка, руб',
    'Итого к оплате, руб',
    'Сумма предоплаты, руб',

    'PDF URL (Drive)',
    'Drive File ID'
  ],

  // Эти колонки скрываем всегда (не группируем)
  ALWAYS_HIDE_HEADERS: [
    'PDF Download URL'
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

  const colsHidden = []; // [{col, wasHidden}] для временного скрытия колонок в PDF
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

    // --- исключаем из PDF колонки корзины (например, "Скидка %" для менеджеров)
    try {
      const cartHeaderRow = findCartHeaderRow_(sh, KP_EXPORT_CFG.CART_HEADER);
      if (cartHeaderRow) {
        const maxCol = sh.getLastColumn();
        const headerVals = sh.getRange(cartHeaderRow, 1, 1, maxCol).getDisplayValues()[0].map(s => String(s || '').trim());
        const titles = (KP_EXPORT_CFG.CART_EXCLUDE_COL_TITLES || []).map(t => String(t || '').trim()).filter(Boolean);
        for (const t of titles) {
          const col = headerVals.findIndex(v => v === t) + 1;
          if (col > 0) {
            const wasHidden = sh.isColumnHiddenByUser(col);
            colsHidden.push({ col, wasHidden });
            if (!wasHidden) sh.hideColumns(col);
          }
        }
      }
    } catch (e) {
      // не блокируем экспорт, если не нашли корзину/колонку
      Logger.log('PDF: не удалось скрыть колонку корзины: ' + e);
    }


    SpreadsheetApp.flush(); // применить скрытие строк/колонок перед экспортом PDF
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
    // вернуть скрытые колонки
    try {
      if (colsHidden.length) {
        // восстанавливаем только те, которые скрыли мы
        for (const it of colsHidden) {
          if (it && it.col && !it.wasHidden) {
            sh.showColumns(it.col);
          }
        }
      }
    } catch (e) {}
  }
}

/* ===================== VALIDATION ===================== */

function validateRequiredKpFields_(sh) {
  // Возвращает массив недостающих "человеческих" названий обязательных полей
  const required = [
    { title: 'КП №', labels: ['КП №', 'КП№', 'Коммерческое предложение №', 'Коммерческое предложение N', 'Коммерческое предложение №:'] },
    { title: 'Дата КП', labels: ['Дата КП', 'Дата КП:'] },
    { title: 'Менеджер', labels: ['Менеджер', 'Менеджер:'] },
    { title: 'Заказчик', labels: ['Заказчик', 'Заказчик:', 'Наименование Заказчика', 'Наименование заказчика'] },
    { title: 'Адрес заказчика', labels: ['Адрес заказчика', 'Адрес Заказчика', 'Адрес заказчика:'] },
  ];

  const missing = [];

  required.forEach((f) => {
    const v = findValueByAnyLabelInColD_(sh, f.labels);
    if (!String(v || '').trim()) missing.push(f.title);
  });

  // Корзина: должны быть позиции с кол-вом > 0
  const cart = extractCartAsJson_(sh);
  if (!cart.items || cart.items.length === 0) {
    missing.push('Корзина: нет позиций (кол-во > 0)');
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

  // Сворачиваем "белые" колонки (неважные), чтобы важные (зелёные) были видны без горизонтального скролла
  try {
    applyLogColumnGrouping_(log);
  } catch (e) {
    // не критично
  }
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


function findCartHeaderRow_(sh, cartHeaderText) {
  const target = String(cartHeaderText || '').trim();
  if (!target) return 0;
  const maxScan = Math.min(300, sh.getLastRow());
  for (let r = 1; r <= maxScan; r++) {
    const v = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (v === target) return r;
  }
  return 0;
}


function extractCartAsJson_(sh) {
  const CART_HEADER = KP_EXPORT_CFG.CART_HEADER; // "Артикул"
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return { items: [], jsonString: '[]' };

  // Ищем строку заголовков корзины по "Артикул" в колонке A
  let headerRow = 0;
  const maxScan = Math.min(lastRow, 250);
  for (let r = 1; r <= maxScan; r++) {
    const v = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (v === CART_HEADER) { headerRow = r; break; }
  }
  if (!headerRow) return { items: [], jsonString: '[]' };

  const startRow = headerRow + 1;
  const numRows = Math.max(0, lastRow - startRow + 1);

  // Ширина корзины: A..K базово + Примечание (L) + Скидка (M) = 13
  // (если колонок меньше — Google вернёт только существующие значения)
  const width = 13;

  const rng = sh.getRange(startRow, 1, numRows, width);
  const values = rng.getValues();

  const items = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const art = String(row[0] || '').trim(); // A
    if (!art) continue;

    const qty = toNumber_(row[6]); // G: Кол-во
    if (!qty || qty <= 0) continue;

    const item = {
      art,
      // D: Наименование/размеры
      name: String(row[3] || '').trim(),
      // E: Ед.изм.
      unit: String(row[4] || '').trim(),
      qty,
      // F: Стоимость оборудования
      price: toNumber_(row[5]),
      // H: Всего за оборудование
      sumEquip: toNumber_(row[7]),
      // I: Стоимость монтажа (за единицу)
      installUnit: toNumber_(row[8]),
      // J: Всего за монтаж
      sumInstall: toNumber_(row[9]),
      // K: Итого
      total: toNumber_(row[10]),
      // L: Примечание (ручное)
      note: String(row[11] || '').trim(),
      // M: Скидка(-)/Наценка(+), % (индивидуальная)
      discountPct: toNumber_(row[12]),
    };

    items.push(item);
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
function findValueByAnyLabelInColD_(sh, labels) {
  for (let i = 0; i < labels.length; i++) {
    const v = findValueByLabelInColD_(sh, labels[i]);
    if (String(v || '').trim()) return v;
  }
  return '';
}


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

/* ===================== JOURNAL COLUMN GROUPING ===================== */

function applyLogColumnGrouping_(sh) {
  const headerRow = JOURNAL_LAYOUT.HEADER_ROW || 1;
  const lastCol = sh.getLastColumn();
  if (lastCol < 2) return;

  // 1) Сброс: показать все колонки и снять старые группы
  try { sh.showColumns(1, sh.getMaxColumns()); } catch (e) {}
  try {
    // снимаем глубину группировки (если есть)
    sh.getRange(headerRow, 1, 1, lastCol).shiftColumnGroupDepth(-10);
  } catch (e) {}

  // 2) Заголовки
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0]
    .map(h => normalizeHeader_(h));

  const headerToCol = {};
  for (let c = 0; c < headers.length; c++) {
    if (headers[c]) headerToCol[headers[c]] = c + 1; // 1-based
  }

  // 3) Какие колонки видимы
  const visibleCols = new Set();
  (JOURNAL_LAYOUT.VISIBLE_HEADERS || []).forEach((h) => {
    const col = headerToCol[normalizeHeader_(h)];
    if (col) visibleCols.add(col);
  });

  // 4) Скрываем "всегда скрытые" колонки
  (JOURNAL_LAYOUT.ALWAYS_HIDE_HEADERS || []).forEach((h) => {
    const col = headerToCol[normalizeHeader_(h)];
    if (col) {
      try { sh.hideColumns(col); } catch (e) {}
      visibleCols.delete(col);
    }
  });

  // 5) Группируем и сворачиваем ВСЕ НЕ-видимые колонки блоками
  let start = null;
  for (let col = 1; col <= lastCol; col++) {
    const isHidden = safeIsColumnHidden_(sh, col);
    const isVisible = visibleCols.has(col);

    const shouldGroup = !isVisible && !isHidden;

    if (shouldGroup) {
      if (start === null) start = col;
    } else {
      if (start !== null) {
        collapseRange_(sh, start, col - 1, headerRow);
        start = null;
      }
    }
  }
  if (start !== null) collapseRange_(sh, start, lastCol, headerRow);
}



function normalizeHeader_(v) {
  return String(v || '')
    .replace(/\s+/g, ' ')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toLowerCase();
}

function safeIsColumnHidden_(sh, col) {
  try { return sh.isColumnHiddenByUser(col); } catch (e) { return false; }
}

function collapseRange_(sh, startCol, endCol, headerRow) {
  if (!startCol || !endCol || endCol < startCol) return;
  const width = endCol - startCol + 1;
  try {
    sh.getRange(headerRow || 1, startCol, 1, width).shiftColumnGroupDepth(1);
    const g = sh.getColumnGroup(startCol, 1);
    if (g) g.collapse();
  } catch (e) {
    // если группировка недоступна в этом аккаунте/документе — просто скрываем диапазон
    try { sh.hideColumns(startCol, width); } catch (e2) {}
  }
}
