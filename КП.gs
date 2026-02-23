/**
 * КП.gs — buildKP (onOpen в main.gs НЕ трогаем)
 *
 * ДОРАБОТКИ:
 * 1) UID-связка Прайс → КП:
 *    - UID переносим в скрытую колонку N (14) в корзине КП
 * 2) Исключаем строки-группы из корзины:
 *    - Берём в КП только строки, где (Кол-во > 0) И (UID не пустой)
 * 3) НОВОЕ ТРЕБОВАНИЕ:
 *    - Если в "Прайс" не выбрано ни одной позиции (qty>0 и UID есть),
 *      то КП НЕ формируем (не чистим лист КП), показываем предупреждение и выходим.
 * 4) FIX "Those columns are out of bounds":
 *    - Перед setColumnWidth гарантируем, что лист КП имеет минимум 14 колонок.
 */

const KP_CFG = {
  SHEET_KP: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.KP) ? CFG.SHEETS.KP : 'КП',
  SHEET_PRICE: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.PRICE) ? CFG.SHEETS.PRICE : 'Прайс',
  HEADER_FILE_ID: (typeof CFG !== 'undefined' && CFG.IDS && CFG.IDS.KP_HEADER_FILE_ID)
    ? CFG.IDS.KP_HEADER_FILE_ID
    : '1E-6KvJH6CPFkEHgG0nLX_NE7rYlm67bw',

  // N = 14: M (13) = скидка %, N (14) = UID (скрыто)
  COL_END: 14,
  COL_PRINT_END: 12, // L

  DEFAULT_COL_WIDTH: 100,
  DEFAULT_ROW_HEIGHT: 21,
  FONT_SIZE: 20,

  // форматы
  NUM_FMT_MONEY: '#,##0.00',
  NUM_FMT_INT: '#,##0',
  NUM_FMT_PCT: '0.00',
  NUM_FMT_TEXT: '@',

  // строки параметров
  PARAM_ROW_DISCOUNT: 19,     // D19
  PARAM_ROW_INSTALL_PCT: 20,  // D20
  PARAM_ROW_DELIVERY: 21,     // D21

  // заголовки "Прайс"
  PRICE_HEADERS: {
    UID: 'UID',
    ART: 'Артикул',
    VIEW1: 'Вид 1',
    VIEW2: 'Вид 2',
    NAME: 'Наименование изделия/ размеры',
    UNIT: 'Ед. изм.',
    COST: 'Стоимость оборудования',
    QTY: 'Кол-во',
  },

  // колонки КП корзины
  CART_COLS: {
    DISCOUNT_PCT: 13, // M
    UID: 14,          // N (скрытая)
  },
};

/** FIX: гарантируем размер листа (иначе setColumnWidth(14, ...) падает) */
function ensureSheetSize_(sh, minRows, minCols) {
  const curRows = sh.getMaxRows();
  const curCols = sh.getMaxColumns();

  if (curRows < minRows) sh.insertRowsAfter(curRows, minRows - curRows);
  if (curCols < minCols) sh.insertColumnsAfter(curCols, minCols - curCols);
}

/**
 * НОВОЕ: быстрый подсчёт выбранных позиций в "Прайс"
 * Условие: qty > 0 И uid не пустой (строки групп отсекаются).
 * Возвращает число выбранных строк-товаров.
 */
function countSelectedPositionsInPrice_(shPrice) {
  const lastRow = shPrice.getLastRow();
  const lastCol = shPrice.getLastColumn();
  if (lastRow < 2) return 0;

  // Прайс обычно с заголовками в 1-й строке, но на всякий случай поддержим 2-ю
  const cand1 = shPrice.getRange(1, 1, 1, lastCol).getValues()[0];
  const cand2 = shPrice.getRange(2, 1, 1, lastCol).getValues()[0];

  let headerRow = 1;
  let map = headerMap_(cand1);
  if (!map[KP_CFG.PRICE_HEADERS.UID] || !map[KP_CFG.PRICE_HEADERS.QTY]) {
    map = headerMap_(cand2);
    headerRow = 2;
  }
  if (!map[KP_CFG.PRICE_HEADERS.UID] || !map[KP_CFG.PRICE_HEADERS.QTY]) {
    // fallback: если заголовки не нашли — используем "тяжёлый" метод через readPriceRows_
    try {
      const rows = readPriceRows_(shPrice);
      return rows.filter(r => (Number(r.qty) || 0) > 0 && String(r.uid || '').trim() !== '').length;
    } catch (e) {
      return 0;
    }
  }

  const colUid = map[KP_CFG.PRICE_HEADERS.UID]; // 1-based
  const colQty = map[KP_CFG.PRICE_HEADERS.QTY]; // 1-based

  const dataStartRow = headerRow + 1;
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) return 0;

  const uidVals = shPrice.getRange(dataStartRow, colUid, numRows, 1).getValues();
  const qtyVals = shPrice.getRange(dataStartRow, colQty, numRows, 1).getValues();

  let cnt = 0;
  for (let i = 0; i < numRows; i++) {
    const uid = String(uidVals[i][0] ?? '').trim();
    if (!uid) continue;

    const q = qtyVals[i][0];
    const n = (typeof q === 'number') ? q : Number(String(q ?? '').replace(',', '.'));
    if ((n || 0) > 0) cnt++;
  }
  return cnt;
}

function buildKP() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    const shPrice = ss.getSheetByName(KP_CFG.SHEET_PRICE);
    if (!shPrice) throw new Error('Не найден лист "Прайс".');

    // НОВОЕ: если нет выбранных позиций — КП НЕ формируем
    const selectedCount = countSelectedPositionsInPrice_(shPrice);
    if (selectedCount <= 0) {
      ui.alert('КП не сформировано', 'В листе "Прайс" не выбрана ни одна позиция (Кол-во > 0).', ui.ButtonSet.OK);
      ss.setActiveSheet(shPrice);
      return;
    }

    const shKP = ss.getSheetByName(KP_CFG.SHEET_KP) || ss.insertSheet(KP_CFG.SHEET_KP);
    ss.setActiveSheet(shKP);

    // FIX: перед любой работой с ширинами/скрытием обеспечиваем колонки
    ensureSheetSize_(shKP, 200, KP_CFG.COL_END);

    clearSheetHardAndResetSizes_(shKP);
    setupKpColumns_(shKP);

    // Шапка A1:L7
    insertHeaderImageOverCells_(shKP, KP_CFG.HEADER_FILE_ID, 1, 1, 7, KP_CFG.COL_PRINT_END);

    // Параметры
    let row = 9;
    row = buildParamsBlock_(shKP, row);

    // Условия (верхний блок)
    row += 1;
    const termsInfo = buildTermsBlock_(shKP, row);
    row = termsInfo.afterRow;

    // Корзина
    row += 2;
    const cartInfo = buildCartBlockFromPrice_(shKP, shPrice, row);
    row = cartInfo.afterRow;

    // Итоги
    row += 1;
    const totalsInfo = buildTotalsBlock_(shKP, row, cartInfo.firstDataRow, cartInfo.lastDataRow);
    row = totalsInfo.afterRow;

    // Нижний блок "УСЛОВИЯ..."
    row += 1;
    buildFooterTermsBlock_(shKP, row, termsInfo, totalsInfo);

    // Шрифт 20 на весь лист
    applyGlobalFontSize_(shKP, KP_CFG.FONT_SIZE);

    // Форматы чисел
    applyNumberFormats_(shKP, cartInfo.firstDataRow, cartInfo.lastDataRow, totalsInfo);

    // Скрываем UID колонку (N) всегда
    try { shKP.hideColumns(KP_CFG.CART_COLS.UID); } catch (e) {}

    shKP.setFrozenRows(0);
    shKP.setFrozenColumns(0);
    shKP.setActiveSelection('A1');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ----------------- ОЧИСТКА + СБРОС ----------------- */

function clearSheetHardAndResetSizes_(sheet) {
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  try { sheet.getImages().forEach(img => img.remove()); } catch (e) {}
  try { const f = sheet.getFilter(); if (f) f.remove(); } catch (e) {}

  const maxR = Math.max(sheet.getMaxRows(), 1);
  const maxC = Math.max(sheet.getMaxColumns(), 1);
  const rngAll = sheet.getRange(1, 1, maxR, maxC);

  try { rngAll.breakApart(); } catch (e) {}
  sheet.clear();

  try {
    sheet.showRows(1, maxR);
    sheet.showColumns(1, maxC);
  } catch (e) {}

  sheet.setRowHeights(1, sheet.getMaxRows(), KP_CFG.DEFAULT_ROW_HEIGHT);
  sheet.setColumnWidths(1, sheet.getMaxColumns(), KP_CFG.DEFAULT_COL_WIDTH);
  sheet.setHiddenGridlines(false);
}

function setupKpColumns_(sh) {
  // FIX: гарантируем колонки до N
  ensureSheetSize_(sh, 50, KP_CFG.COL_END);

  sh.setColumnWidth(1, 180);  // A (Артикул)
  sh.setColumnWidth(2, 200);  // B (Вид 1)
  sh.setColumnWidth(3, 200);  // C (Вид 2)
  sh.setColumnWidth(4, 360);  // D (Наименование)
  sh.setColumnWidth(5, 60);   // E
  sh.setColumnWidth(6, 150);  // F
  sh.setColumnWidth(7, 70);   // G
  sh.setColumnWidth(8, 160);  // H
  sh.setColumnWidth(9, 160);  // I
  sh.setColumnWidth(10, 160); // J
  sh.setColumnWidth(11, 160); // K
  sh.setColumnWidth(12, 260); // L (Примечание)
  sh.setColumnWidth(13, 120); // M (Скидка %)
  sh.setColumnWidth(14, 60);  // N (UID, скрыто)

  for (let r = 1; r <= 7; r++) sh.setRowHeight(r, 30);
}

/* ---------------------- ШАПКА ---------------------- */

function insertHeaderImageOverCells_(sh, fileId, startRow, startCol, endRow, endCol) {
  try { sh.getImages().forEach(img => img.remove()); } catch (e) {}

  const blob = DriveApp.getFileById(fileId).getBlob();
  const img = sh.insertImage(blob, startCol, startRow);

  const widthPx = sumColumnWidths_(sh, startCol, endCol);
  const heightPx = sumRowHeights_(sh, startRow, endRow);

  img.setAnchorCell(sh.getRange(startRow, startCol));
  img.setWidth(widthPx);
  img.setHeight(heightPx);
}

function sumColumnWidths_(sh, c1, c2) {
  let w = 0;
  for (let c = c1; c <= c2; c++) w += sh.getColumnWidth(c);
  return w;
}

function sumRowHeights_(sh, r1, r2) {
  let h = 0;
  for (let r = r1; r <= r2; r++) h += sh.getRowHeight(r);
  return h;
}

/* ---------------------- ПАРАМЕТРЫ + НАСТРОЙКИ РАСЧЁТА ---------------------- */

function buildParamsBlock_(sh, startRow) {
  // Строка 9: пустая
  sh.getRange(startRow, 1, 1, KP_CFG.COL_END)
    .merge()
    .setValue('')
    .setBackground(null)
    .setFontWeight('normal')
    .setHorizontalAlignment('left');

  const baseLabels = [
    'Наименование Заказчика',
    'Адрес Заказчика',
    '№ Договора',
    'Менеджер',
    'Телефон',
    'Дата КП',
    'Коммерческое предложение №',
  ];

  const row0 = startRow + 1; // 10
  const textRows = new Set([10, 11, 12, 13, 14, 16]);

  for (let i = 0; i < baseLabels.length; i++) {
    const r = row0 + i;

    sh.getRange(r, 1, 1, 3)
      .merge()
      .setValue(baseLabels[i])
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    const input = sh.getRange(r, 4, 1, (KP_CFG.COL_END - 3))
      .merge()
      .setValue('')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    if (textRows.has(r)) input.setNumberFormat(KP_CFG.NUM_FMT_TEXT);

    if (baseLabels[i] === 'Дата КП') {
      input.setNumberFormat('dd.mm.yyyy');
      sh.getRange(r, 4).setValue(new Date());
    }

    // выпадающий список менеджеров
    if (baseLabels[i] === 'Менеджер') {
      try {
        const ss = SpreadsheetApp.getActive();
        const refName = (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.REF) ? CFG.SHEETS.REF : 'Справочник';
        const shRef = ss.getSheetByName(refName);
        if (shRef) {
          const last = Math.max(shRef.getLastRow(), 2);
          const num = Math.max(1, last - 1);
          const listRange = shRef.getRange(2, 1, num, 1); // A2:A
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(listRange, true)
            .setAllowInvalid(false)
            .build();
          input.setDataValidation(rule);
        }
      } catch (e) {}
    }

    // телефон по менеджеру
    if (baseLabels[i] === 'Телефон') {
      try {
        const rMgr = row0 + baseLabels.indexOf('Менеджер');
        const mgrCell = `$D$${rMgr}`;
        const refName = (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.REF) ? CFG.SHEETS.REF : 'Справочник';
        sh.getRange(r, 4).setFormula(`=IF(${mgrCell}="";"";IFERROR(VLOOKUP(${mgrCell};${refName}!$A$2:$B;2;FALSE);""))`);
        input.setNumberFormat(KP_CFG.NUM_FMT_TEXT);
      } catch (e) {}
    }
  }

  // пустая строка 17
  const rBlank = row0 + baseLabels.length; // 17
  sh.getRange(rBlank, 1, 1, KP_CFG.COL_END).merge().setValue('').setBackground(null);

  // заголовок "Настройки расчёта" — 18
  const rSettingsTitle = rBlank + 1; // 18
  sh.getRange(rSettingsTitle, 1, 1, KP_CFG.COL_END)
    .merge()
    .setValue('Настройки расчёта')
    .setFontWeight('bold')
    .setBackground('#eeeeee')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(true);

  const settingsLabels = [
    'Скидка (-) / Наценка (+), %',
    'Размер монтажа от стоимости оборудования, %',
    'Доставка, руб',
  ];

  for (let i = 0; i < settingsLabels.length; i++) {
    const r = rSettingsTitle + 1 + i; // 19..21

    sh.getRange(r, 1, 1, 3).merge()
      .setValue(settingsLabels[i])
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    const input = sh.getRange(r, 4, 1, (KP_CFG.COL_END - 3)).merge()
      .setValue('')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    if (settingsLabels[i] === 'Скидка (-) / Наценка (+), %') {
      input.setNumberFormat(KP_CFG.NUM_FMT_PCT);
      sh.getRange(r, 4).setValue(0);
    }
    if (settingsLabels[i] === 'Размер монтажа от стоимости оборудования, %') {
      input.setNumberFormat(KP_CFG.NUM_FMT_PCT);
      sh.getRange(r, 4).setValue(25);
    }
    if (settingsLabels[i] === 'Доставка, руб') {
      input.setNumberFormat(KP_CFG.NUM_FMT_MONEY);
      sh.getRange(r, 4).setValue(0);
    }
  }

  const lastRow = rSettingsTitle + settingsLabels.length; // 21
  sh.getRange(startRow, 1, lastRow - startRow + 1, KP_CFG.COL_END).setBorder(true, true, true, true, true, true);

  return lastRow + 1;
}

/* ------------------- УСЛОВИЯ / СРОКИ (верхний блок) ------------------- */

function buildTermsBlock_(sh, startRow) {
  sh.getRange(startRow, 1, 1, KP_CFG.COL_END)
    .merge()
    .setValue('Условия и сроки поставки (изменяемые)')
    .setFontWeight('bold')
    .setBackground('#eeeeee')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(true);

  const rPrepay = startRow + 1;
  const rMain = startRow + 2;
  const rEco = startRow + 3;
  const rValid = startRow + 4;

  sh.getRange(rPrepay, 1, 1, 9).merge()
    .setValue('Предоплата за оборудование составляет:')
    .setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');

  const prepayCell = sh.getRange(rPrepay, 10, 1, 4).merge();
  prepayCell.setValue(0.7).setNumberFormat('0%')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sh.getRange(rMain, 1, 1, 9).merge()
    .setValue('Срок поставки Основное производство исчисляется с момента поступления предоплаты на р/счет и составляет:')
    .setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(rMain, 10, 1, 4).merge()
    .setValue('35-40 рабочих дней')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

  sh.getRange(rEco, 1, 1, 9).merge()
    .setValue('Срок поставки ЭКО-серия исчисляется с момента поступления предоплаты на р/счет и составляет:')
    .setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(rEco, 10, 1, 4).merge()
    .setValue('40-45 рабочих дней')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

  sh.getRange(rValid, 1, 1, 9).merge()
    .setValue('Данное КП действительно в течение:')
    .setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(rValid, 10, 1, 4).merge()
    .setValue(7)
    .setNumberFormat('0')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  const afterRow = rValid + 1;
  sh.getRange(startRow, 1, (afterRow - startRow), KP_CFG.COL_END).setBorder(true, true, true, true, true, true);

  return {
    afterRow,
    cells: {
      prepayPct: `J${rPrepay}`,
      mainLead: `J${rMain}`,
      ecoLead: `J${rEco}`,
      validityDays: `J${rValid}`,
    },
  };
}

/* ------------------------ КОРЗИНА ------------------------ */

function buildCartBlockFromPrice_(shKP, shPrice, startRow) {
  const headerRow = startRow;

  const headers = [
    'Артикул',
    'Вид 1',
    'Вид 2',
    'Наименование / размеры',
    'Ед.',
    'Стоимость оборудования',
    'Кол-во',
    'Всего за оборудование',
    'Стоимость монтажа',
    'Всего за монтаж',
    'Итого',
    'Примечание',
    'Скидка (-) / Наценка (+), %',
    'UID',
  ];

  shKP.getRange(headerRow, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f5f5f5')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  try { shKP.hideColumns(KP_CFG.CART_COLS.UID); } catch (e) {}

  const priceImgByRow = buildImagesByRow_(shPrice);
  const priceRows = readPriceRows_(shPrice);

  // ВАЖНО: берём только реальные товары (есть UID) и qty > 0
  const cartRows = priceRows.filter(r =>
    (Number(r.qty) || 0) > 0 && String(r.uid || '').trim() !== ''
  );

  let outRow = headerRow + 1;
  const firstDataRow = outRow;

  // (теоретически сюда уже не должны попадать 0 позиций, т.к. мы проверили в buildKP,
  //  но оставляем защиту)
  if (cartRows.length === 0) {
    shKP.getRange(outRow, 1, 1, KP_CFG.COL_END)
      .merge()
      .setValue('В «Прайс» нет выбранных товарных строк (Кол-во > 0 и UID заполнен).')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    shKP.getRange(headerRow, 1, (outRow - headerRow) + 1, KP_CFG.COL_END)
      .setBorder(true, true, true, true, true, true);

    return { headerRow, firstDataRow, lastDataRow: outRow, afterRow: outRow + 1 };
  }

  for (const item of cartRows) {
    shKP.setRowHeight(outRow, 200);

    shKP.getRange(outRow, 1).setValue(item.art);

    setViewImageFromPrice_(shKP, shPrice, priceImgByRow, outRow, 2, item.sourceRow, item.sourceColView1, item.view1);
    setViewImageFromPrice_(shKP, shPrice, priceImgByRow, outRow, 3, item.sourceRow, item.sourceColView2, item.view2);

    shKP.getRange(outRow, 4).setValue(item.name);
    shKP.getRange(outRow, 5).setValue(item.unit);

    // M скидка%
    shKP.getRange(outRow, KP_CFG.CART_COLS.DISCOUNT_PCT).setFormula(`=$D$${KP_CFG.PARAM_ROW_DISCOUNT}`);

    // N UID (скрыто)
    shKP.getRange(outRow, KP_CFG.CART_COLS.UID).setValue(item.uid || '').setNumberFormat(KP_CFG.NUM_FMT_TEXT);

    // L Примечание
    shKP.getRange(outRow, 12).setValue('');

    // F цена с учетом скидки%
    const ruPrice = toRuNumberLiteral_(item.cost);
    shKP.getRange(outRow, 6).setFormula(`=${ruPrice}*(1+M${outRow}/100)`);

    shKP.getRange(outRow, 7).setValue(item.qty); // G
    shKP.getRange(outRow, 8).setFormulaR1C1('=RC[-2]*RC[-1]'); // H
    shKP.getRange(outRow, 9).setFormula(`=F${outRow}*$D$${KP_CFG.PARAM_ROW_INSTALL_PCT}/100`); // I
    shKP.getRange(outRow, 10).setFormulaR1C1('=RC[-1]*RC[-3]'); // J
    shKP.getRange(outRow, 11).setFormulaR1C1('=RC[-3]+RC[-1]'); // K

    outRow++;
  }

  const lastDataRow = outRow - 1;
  const dataCount = lastDataRow - firstDataRow + 1;

  if (dataCount > 0) {
    shKP.getRange(firstDataRow, 1, dataCount, KP_CFG.COL_END)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  shKP.getRange(headerRow, 1, outRow - headerRow, KP_CFG.COL_END)
    .setBorder(true, true, true, true, true, true);

  return { headerRow, firstDataRow, lastDataRow, afterRow: outRow };
}

function toRuNumberLiteral_(v) {
  if (typeof v === 'number') return String(v).replace('.', ',');
  let s = String(v ?? '').trim();
  if (!s) return '0';
  s = s.replace(/\s+/g, '');
  s = s.replace('.', ',');
  return s;
}

function setViewImageFromPrice_(shKP, shPrice, priceImgByRow, targetRow, targetCol, sourceRow, sourceCol, anyVal) {
  const targetCell = shKP.getRange(targetRow, targetCol);
  const url = extractImageUrl_(anyVal);

  if (url) {
    const w = shKP.getColumnWidth(targetCol);
    const h = shKP.getRowHeight(targetRow);
    const fw = Math.max(20, w - 16);
    const fh = Math.max(20, h - 10);
    targetCell.setFormula(`=IMAGE("${url}";4;${fh};${fw})`);
    return;
  }

  if (!sourceRow || !sourceCol) return;

  const img = findNearestImageInRow_(priceImgByRow, sourceRow, sourceCol, 3);
  if (!img) return;

  const blob = img.getBlob();
  const inserted = shKP.insertImage(blob, targetCol, targetRow);
  inserted.setAnchorCell(targetCell);

  const w = shKP.getColumnWidth(targetCol);
  const h = shKP.getRowHeight(targetRow);

  inserted.setWidth(Math.max(10, w - 16));
  inserted.setHeight(Math.max(10, h - 10));
}

function buildImagesByRow_(sheet) {
  const byRow = {};
  let imgs = [];
  try { imgs = sheet.getImages(); } catch (e) { imgs = []; }

  imgs.forEach(img => {
    try {
      const cell = img.getAnchorCell();
      const r = cell.getRow();
      const c = cell.getColumn();
      if (!byRow[r]) byRow[r] = [];
      byRow[r].push({ col: c, img });
    } catch (e) {}
  });

  Object.keys(byRow).forEach(r => byRow[Number(r)].sort((a, b) => a.col - b.col));
  return byRow;
}

function findNearestImageInRow_(byRow, row, desiredCol, radiusCols) {
  const arr = byRow[row];
  if (!arr || arr.length === 0) return null;

  let best = null;
  let bestDist = 9999;

  for (const it of arr) {
    const d = Math.abs(it.col - desiredCol);
    if (d <= radiusCols && d < bestDist) {
      bestDist = d;
      best = it.img;
    }
  }
  return best;
}

/* --------------------- ЧТЕНИЕ ПРАЙСА (RichText URL) --------------------- */

function readPriceRows_(shPrice) {
  const lastRow = shPrice.getLastRow();
  const lastCol = shPrice.getLastColumn();
  if (lastRow < 2) return [];

  const cand1 = shPrice.getRange(1, 1, 1, lastCol).getValues()[0];
  const cand2 = shPrice.getRange(2, 1, 1, lastCol).getValues()[0];

  let headerRow = 0;
  let map = headerMap_(cand1);

  if (!map[KP_CFG.PRICE_HEADERS.QTY] || !map[KP_CFG.PRICE_HEADERS.COST]) {
    map = headerMap_(cand2);
    headerRow = 1;
  }
  if (!map[KP_CFG.PRICE_HEADERS.QTY] || !map[KP_CFG.PRICE_HEADERS.COST]) return [];

  const col = (name) => (map[name] || 0);

  const idx = {
    uid: col(KP_CFG.PRICE_HEADERS.UID),
    art: col(KP_CFG.PRICE_HEADERS.ART),
    v1: col(KP_CFG.PRICE_HEADERS.VIEW1),
    v2: col(KP_CFG.PRICE_HEADERS.VIEW2),
    name: col(KP_CFG.PRICE_HEADERS.NAME),
    unit: col(KP_CFG.PRICE_HEADERS.UNIT),
    cost: col(KP_CFG.PRICE_HEADERS.COST),
    qty: col(KP_CFG.PRICE_HEADERS.QTY),
  };

  const dataStartRow = headerRow + 2;
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) return [];

  const rng = shPrice.getRange(dataStartRow, 1, numRows, lastCol);
  const values = rng.getValues();
  const formulas = rng.getFormulas();

  let rtV1 = null;
  let rtV2 = null;
  try { if (idx.v1) rtV1 = shPrice.getRange(dataStartRow, idx.v1, numRows, 1).getRichTextValues(); } catch (e) {}
  try { if (idx.v2) rtV2 = shPrice.getRange(dataStartRow, idx.v2, numRows, 1).getRichTextValues(); } catch (e) {}

  const out = [];

  for (let i = 0; i < values.length; i++) {
    const sheetRow = dataStartRow + i;
    const row = values[i];
    const rowF = formulas[i];

    const uid = idx.uid ? String(row[idx.uid - 1] ?? '').trim() : '';
    const art = idx.art ? String(row[idx.art - 1] ?? '').trim() : '';
    const qty = idx.qty ? row[idx.qty - 1] : '';
    const cost = idx.cost ? row[idx.cost - 1] : '';

    if (!uid && !art && !qty && !cost) continue;

    const rich1 = (rtV1 && rtV1[i]) ? rtV1[i][0] : null;
    const rich2 = (rtV2 && rtV2[i]) ? rtV2[i][0] : null;

    const url1 = extractUrlFromRichText_(rich1);
    const url2 = extractUrlFromRichText_(rich2);

    const fallbackV1 = idx.v1 ? (rowF[idx.v1 - 1] || row[idx.v1 - 1]) : '';
    const fallbackV2 = idx.v2 ? (rowF[idx.v2 - 1] || row[idx.v2 - 1]) : '';

    out.push({
      uid,
      art,
      view1: url1 || fallbackV1,
      view2: url2 || fallbackV2,
      name: idx.name ? row[idx.name - 1] : '',
      unit: idx.unit ? row[idx.unit - 1] : '',
      cost: cost,
      qty: qty,
      sourceRow: sheetRow,
      sourceColView1: idx.v1 || null,
      sourceColView2: idx.v2 || null,
    });
  }

  return out;
}

function extractUrlFromRichText_(rich) {
  try {
    if (!rich) return '';
    const u = rich.getLinkUrl && rich.getLinkUrl();
    if (u) return u;

    const runs = rich.getRuns && rich.getRuns();
    if (runs && runs.length) {
      for (const run of runs) {
        const ru = run.getLinkUrl && run.getLinkUrl();
        if (ru) return ru;
      }
    }
  } catch (e) {}
  return '';
}

function headerMap_(headerRowValues) {
  const m = {};
  for (let c = 0; c < headerRowValues.length; c++) {
    const h = String(headerRowValues[c] ?? '').trim();
    if (h) m[h] = c + 1; // 1-based
  }
  return m;
}

function extractImageUrl_(val) {
  if (!val) return '';
  const s = String(val).trim();
  const m = s.match(/IMAGE\(\s*"([^"]+)"\s*/i);
  if (m && m[1]) return m[1];
  const h = s.match(/HYPERLINK\(\s*"([^"]+)"/i);
  if (h && h[1]) return h[1];
  if (/^https?:\/\//i.test(s)) return s;
  return '';
}

/* ----------------------------- ИТОГИ ----------------------------- */

function buildTotalsBlock_(sh, startRow, firstDataRow, lastDataRow) {
  const r0 = startRow;
  const ds = Math.max(firstDataRow || 0, 1);
  const de = Math.max(lastDataRow || ds, ds);

  const setLine = (row, label, formula, fmt) => {
    sh.getRange(row, 1, 1, 3).merge()
      .setValue(label)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    const cell = sh.getRange(row, 4); // D
    cell.setFormula(formula)
      .setNumberFormat(fmt || KP_CFG.NUM_FMT_MONEY)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  };

  setLine(r0, 'Итого за оборудование, руб', `=SUM(H${ds}:H${de})`, KP_CFG.NUM_FMT_MONEY);
  setLine(r0 + 1, 'Доставка, руб', `=D${KP_CFG.PARAM_ROW_DELIVERY}`, KP_CFG.NUM_FMT_MONEY);
  setLine(r0 + 2, 'Монтаж, руб', `=SUM(J${ds}:J${de})`, KP_CFG.NUM_FMT_MONEY);

  setLine(r0 + 3, 'Итого к оплате, руб', `=D${r0}+D${r0 + 2}+D${r0 + 1}`, KP_CFG.NUM_FMT_MONEY);
  setLine(r0 + 4, 'В том числе НДС 22%, руб', `=D${r0 + 3}*22/122`, KP_CFG.NUM_FMT_MONEY);

  sh.getRange(r0, 1, 5, 4).setBorder(true, true, true, true, true, true);

  return {
    afterRow: r0 + 5,
    cells: {
      equipmentTotal: `D${r0}`,
      delivery: `D${r0 + 1}`,
      installTotal: `D${r0 + 2}`,
      toPay: `D${r0 + 3}`,
      vat: `D${r0 + 4}`,
    },
  };
}

/* ------------------- НИЖНИЙ БЛОК УСЛОВИЙ ------------------- */

function buildFooterTermsBlock_(sh, startRow, termsInfo, totalsInfo) {
  sh.getRange(startRow, 1, 1, KP_CFG.COL_END)
    .merge()
    .setValue('УСЛОВИЯ И СРОКИ ПОСТАВКИ :')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(true);

  const r1 = startRow + 1;
  const r2 = startRow + 2;
  const r3 = startRow + 3;
  const r4 = startRow + 4;
  const r5 = startRow + 5;
  const r6 = startRow + 6;
  const r7 = startRow + 7;

  const lineFormula = (r, formula) => {
    sh.getRange(r, 1, 1, KP_CFG.COL_END)
      .merge()
      .setFormula(formula)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  };

  const lineValue = (r, value) => {
    sh.getRange(r, 1, 1, KP_CFG.COL_END)
      .merge()
      .setValue(value)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  };

  lineFormula(
    r1,
    `="Предоплата за оборудование составляет: "&TEXT(${termsInfo.cells.prepayPct};"0%")&", что составляет "&TEXT(${totalsInfo.cells.equipmentTotal}*${termsInfo.cells.prepayPct};"#,##0.00")&" рублей"`
  );

  lineValue(
    r2,
    'Оставшаяся сумма оплачивается в течение 3-х дней с даты уведомления о готовности к отгрузке в Раменском районе, сельском поселении Ганусовское, промпарк A107'
  );

  lineFormula(r3, `="Срок поставки Основное производство исчисляется с момента поступления предоплаты на р/счет и составляет: "&${termsInfo.cells.mainLead}`);
  lineFormula(r4, `="Срок поставки ЭКО-серия исчисляется с момента поступления предоплаты на р/счет и составляет: "&${termsInfo.cells.ecoLead}`);

  lineValue(r5, 'Оборудование поставляется в упаковке в разобранном/частично собранном виде');
  lineValue(r6, 'Срок поставки может быть изменен в случае изменения технического задания, либо при необходимости внесения изменений заказчиком в конструкцию изделий');

  lineFormula(r7, `="Данное КП действительно в течение: "&${termsInfo.cells.validityDays}&" (семи) дней"`);

  return r7 + 1;
}

/* ------------------- Форматы ------------------- */

function applyGlobalFontSize_(sh, size) {
  try {
    const maxR = Math.max(sh.getMaxRows(), sh.getLastRow(), 1);
    const maxC = Math.max(sh.getMaxColumns(), sh.getLastColumn(), 1);
    sh.getRange(1, 1, maxR, Math.min(maxC, KP_CFG.COL_END)).setFontSize(size);
  } catch (e) {}
}

function applyNumberFormats_(sh, firstDataRow, lastDataRow, totalsInfo) {
  if (firstDataRow && lastDataRow && lastDataRow >= firstDataRow) {
    const n = lastDataRow - firstDataRow + 1;

    sh.getRange(firstDataRow, 6, n, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // F
    sh.getRange(firstDataRow, 8, n, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // H
    sh.getRange(firstDataRow, 9, n, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // I
    sh.getRange(firstDataRow, 10, n, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // J
    sh.getRange(firstDataRow, 11, n, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // K

    sh.getRange(firstDataRow, 7, n, 1).setNumberFormat(KP_CFG.NUM_FMT_INT);   // G
    sh.getRange(firstDataRow, 12, n, 1).setNumberFormat(KP_CFG.NUM_FMT_TEXT); // L
    sh.getRange(firstDataRow, 13, n, 1).setNumberFormat(KP_CFG.NUM_FMT_PCT);  // M
    sh.getRange(firstDataRow, 14, n, 1).setNumberFormat(KP_CFG.NUM_FMT_TEXT); // N (UID)
  }

  if (totalsInfo && totalsInfo.cells) {
    try { sh.getRange(totalsInfo.cells.equipmentTotal).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.delivery).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.installTotal).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.toPay).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.vat).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
  }
}