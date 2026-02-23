/**
 * КП.gs — формирование КП
 *
 * Добавлено:
 * 1) Диалог при запуске "Сформировать КП":
 *    - ДА  : сформировать заново С ПОЛНОЙ ОЧИСТКОЙ листа "КП"
 *    - НЕТ : обновить КП БЕЗ ПОЛНОЙ ОЧИСТКИ (сохранить ручные данные; обновить корзину/итоги)
 *    - ОТМЕНА: ничего не делать
 *
 * 2) Проверка выбора в "Прайс":
 *    - Если не выбрано ни одной позиции (qty>0 и UID заполнен) — КП не формируем
 *
 * 3) Исключаем строки-группы из корзины:
 *    - В КП попадают только строки где (Кол-во > 0) И (UID не пустой)
 *
 * 4) UID в КП:
 *    - UID переносим в скрытую колонку N (14) корзины КП
 *
 * 5) FIX "Those columns are out of bounds":
 *    - Гарантируем, что на листе КП есть минимум 14 колонок.
 *
 * 6) FIX ошибки "#REF!" в D21 при выборе "НЕТ":
 *    - При обновлении итогов ищем строки итогов ТОЛЬКО начиная с блока итогов (totalsStartRow),
 *      чтобы не попасть в "Доставка, руб" в блоке "Настройки расчёта".
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

  NUM_FMT_MONEY: '#,##0.00',
  NUM_FMT_INT: '#,##0',
  NUM_FMT_PCT: '0.00',
  NUM_FMT_TEXT: '@',

  // строки параметров
  PARAM_ROW_DISCOUNT: 19,     // D19
  PARAM_ROW_INSTALL_PCT: 20,  // D20
  PARAM_ROW_DELIVERY: 21,     // D21

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

  CART_COLS: {
    DISCOUNT_PCT: 13, // M
    UID: 14,          // N (скрытая)
  },
};

/* ========================= ENTRY POINT (меню) ========================= */

function buildKP() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const shPrice = ss.getSheetByName(KP_CFG.SHEET_PRICE);
  if (!shPrice) {
    ui.alert('Ошибка', 'Не найден лист "Прайс".', ui.ButtonSet.OK);
    return;
  }

  // 1) Сначала проверяем, что есть выбранные позиции (qty>0 и UID заполнен)
  const selectedCount = kp_countSelectedPositionsInPrice_(shPrice);
  if (selectedCount <= 0) {
    ui.alert('КП не сформировано', 'В листе "Прайс" не выбрана ни одна позиция (Кол-во > 0).', ui.ButtonSet.OK);
    ss.setActiveSheet(shPrice);
    return;
  }

  // 2) Диалог выбора режима
  const msg =
    'Выберите режим формирования КП:\n\n' +
    'ДА — сформировать заново с полной очисткой листа «КП» (все данные, введённые вручную, будут удалены).\n\n' +
    'НЕТ — обновить КП без полной очистки (ручные данные сохранятся; корзина/итоги будут обновлены).\n\n' +
    'ОТМЕНА — ничего не делать.';

  const btn = ui.alert('Формирование КП', msg, ui.ButtonSet.YES_NO_CANCEL);

  if (btn === ui.Button.CANCEL || btn === ui.Button.CLOSE) return;

  if (btn === ui.Button.YES) {
    kp_buildKP_reset_();
  } else if (btn === ui.Button.NO) {
    kp_buildKP_update_();
  }
}

/* ========================= MODE: RESET (полная очистка) ========================= */

function kp_buildKP_reset_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const ss = SpreadsheetApp.getActive();
    const shPrice = ss.getSheetByName(KP_CFG.SHEET_PRICE);
    if (!shPrice) throw new Error('Не найден лист "Прайс".');

    // защита
    const selectedCount = kp_countSelectedPositionsInPrice_(shPrice);
    if (selectedCount <= 0) {
      SpreadsheetApp.getUi().alert('КП не сформировано', 'В листе "Прайс" не выбрана ни одна позиция (Кол-во > 0).', SpreadsheetApp.getUi().ButtonSet.OK);
      ss.setActiveSheet(shPrice);
      return;
    }

    const shKP = ss.getSheetByName(KP_CFG.SHEET_KP) || ss.insertSheet(KP_CFG.SHEET_KP);
    ss.setActiveSheet(shKP);

    kp_ensureSheetSize_(shKP, 200, KP_CFG.COL_END);

    kp_clearSheetHardAndResetSizes_(shKP);
    kp_setupKpColumns_(shKP);

    kp_insertHeaderImageOverCells_(shKP, KP_CFG.HEADER_FILE_ID, 1, 1, 7, KP_CFG.COL_PRINT_END);

    let row = 9;
    row = kp_buildParamsBlock_(shKP, row);

    row += 1;
    const termsInfo = kp_buildTermsBlock_(shKP, row);
    row = termsInfo.afterRow;

    row += 2;
    const cartInfo = kp_buildCartBlockFromPrice_(shKP, shPrice, row);
    row = cartInfo.afterRow;

    row += 1;
    const totalsInfo = kp_buildTotalsBlock_(shKP, row, cartInfo.firstDataRow, cartInfo.lastDataRow);
    row = totalsInfo.afterRow;

    row += 1;
    kp_buildFooterTermsBlock_(shKP, row, termsInfo, totalsInfo);

    kp_applyGlobalFontSize_(shKP, KP_CFG.FONT_SIZE);
    kp_applyNumberFormats_(shKP, cartInfo.firstDataRow, cartInfo.lastDataRow, totalsInfo);

    try { shKP.hideColumns(KP_CFG.CART_COLS.UID); } catch (e) {}

    shKP.setFrozenRows(0);
    shKP.setFrozenColumns(0);
    shKP.setActiveSelection('A1');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ========================= MODE: UPDATE (без полной очистки) ========================= */

function kp_buildKP_update_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    const shPrice = ss.getSheetByName(KP_CFG.SHEET_PRICE);
    if (!shPrice) throw new Error('Не найден лист "Прайс".');

    const selectedCount = kp_countSelectedPositionsInPrice_(shPrice);
    if (selectedCount <= 0) {
      ui.alert('КП не обновлено', 'В листе "Прайс" не выбрана ни одна позиция (Кол-во > 0).', ui.ButtonSet.OK);
      ss.setActiveSheet(shPrice);
      return;
    }

    const shKP = ss.getSheetByName(KP_CFG.SHEET_KP);
    if (!shKP) {
      kp_buildKP_reset_();
      return;
    }

    ss.setActiveSheet(shKP);
    kp_ensureSheetSize_(shKP, 200, KP_CFG.COL_END);

    const ok = kp_updateCartAndTotalsInPlace_(shKP, shPrice);

    if (!ok) {
      ui.alert('Обновление без очистки невозможно', 'Не удалось найти структуру КП. Выполняю формирование заново с очисткой.', ui.ButtonSet.OK);
      kp_buildKP_reset_();
      return;
    }

    try { shKP.hideColumns(KP_CFG.CART_COLS.UID); } catch (e) {}
    shKP.setActiveSelection('A1');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Обновление "в месте":
 * - НЕ трогаем верхние поля (заказчик/адрес/договор/менеджер и т.п.)
 * - Пересобираем ТОЛЬКО корзину и обновляем формулы итогов под новый диапазон
 * - Сохраняем "Примечание" и индивидуальную скидку по UID (если были заполнены вручную)
 *
 * Возвращает true/false (смогли ли обновить без полной очистки).
 */
function kp_updateCartAndTotalsInPlace_(shKP, shPrice) {
  // 1) Находим строку заголовка корзины ("Артикул" в колонке A)
  const cartHeaderRow = kp_findRowByExactTextInColA_(shKP, 'Артикул', 1, 800);
  if (!cartHeaderRow) return false;

  // 2) Находим строку, где начинается блок итогов
  let totalsStartRow =
    kp_findRowByExactTextInColA_(shKP, 'Итого за оборудование, руб', 1, 3000) ||
    kp_findRowByExactTextInColA_(shKP, 'Итого оборудование, руб', 1, 3000);
  if (!totalsStartRow) return false;

  // 3) Читаем текущую шапку корзины, определяем колонки UID/Примечание/Скидка
  const hdrVals = shKP.getRange(cartHeaderRow, 1, 1, KP_CFG.COL_END).getDisplayValues()[0];
  const hdrNorm = hdrVals.map(kp_normHeader_);

  const colArt = kp_colByExactHeader_(hdrNorm, 'артикул') || 1;
  const colNote = kp_colByExactHeader_(hdrNorm, 'примечание') || 12;
  const colDiscount = kp_colByContainsAll_(hdrNorm, ['скидка', '%']) || 13;
  const colUid = kp_colByExactHeader_(hdrNorm, 'uid') || 14;

  // 4) Считываем существующие Примечания/Скидку по UID
  const old = kp_readExistingCartManualByUid_(shKP, cartHeaderRow, totalsStartRow, colUid, colNote, colDiscount, colArt);

  // 5) Берём новые позиции из Прайса (ТОЛЬКО qty>0 и uid != '')
  const priceImgByRow = kp_buildImagesByRow_(shPrice);
  const priceRows = kp_readPriceRows_(shPrice);
  const cartRows = priceRows.filter(r => (Number(r.qty) || 0) > 0 && String(r.uid || '').trim() !== '');
  if (!cartRows.length) return false;

  // 6) Гарантируем, что между корзиной и итогами есть место
  const targetTotalsStart = cartHeaderRow + cartRows.length + 2;
  if (totalsStartRow < targetTotalsStart) {
    const delta = targetTotalsStart - totalsStartRow;
    shKP.insertRowsBefore(totalsStartRow, delta);
    totalsStartRow += delta; // сдвинулась вниз
  }

  // 7) Удаляем изображения в старой зоне корзины (чтобы не наслаивались)
  kp_removeImagesInRowRange_(shKP, cartHeaderRow + 1, totalsStartRow - 2);

  // 8) Обновляем шапку корзины
  const headers = [
    'Артикул', 'Вид 1', 'Вид 2', 'Наименование / размеры', 'Ед.',
    'Стоимость оборудования', 'Кол-во', 'Всего за оборудование',
    'Стоимость монтажа', 'Всего за монтаж', 'Итого',
    'Примечание', 'Скидка (-) / Наценка (+), %', 'UID'
  ];
  shKP.getRange(cartHeaderRow, 1, 1, KP_CFG.COL_END).setValues([headers]);
  shKP.getRange(cartHeaderRow, 1, 1, KP_CFG.COL_END)
    .setFontWeight('bold')
    .setBackground('#f5f5f5')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // 9) Чистим старые строки корзины (в A:N)
  const dataStartRow = cartHeaderRow + 1;
  const clearRows = Math.max(old.oldRowsCount, cartRows.length);
  if (clearRows > 0) {
    shKP.getRange(dataStartRow, 1, clearRows, KP_CFG.COL_END).clearContent();
    for (let rr = dataStartRow + cartRows.length; rr < dataStartRow + clearRows; rr++) {
      try { shKP.setRowHeight(rr, KP_CFG.DEFAULT_ROW_HEIGHT); } catch (e) {}
    }
  }

  // 10) Записываем новую корзину
  let outRow = dataStartRow;
  for (const item of cartRows) {
    shKP.setRowHeight(outRow, 200);

    shKP.getRange(outRow, 1).setValue(item.art);

    kp_setViewImageFromPrice_(shKP, shPrice, priceImgByRow, outRow, 2, item.sourceRow, item.sourceColView1, item.view1);
    kp_setViewImageFromPrice_(shKP, shPrice, priceImgByRow, outRow, 3, item.sourceRow, item.sourceColView2, item.view2);

    shKP.getRange(outRow, 4).setValue(item.name);
    shKP.getRange(outRow, 5).setValue(item.unit);

    // UID (N)
    shKP.getRange(outRow, KP_CFG.CART_COLS.UID).setValue(item.uid || '').setNumberFormat(KP_CFG.NUM_FMT_TEXT);

    // Примечание (L) — восстановим по UID если было
    const prevManual = old.byUid[item.uid] || null;
    const note = prevManual ? (prevManual.note || '') : '';
    shKP.getRange(outRow, 12).setValue(note).setNumberFormat(KP_CFG.NUM_FMT_TEXT);

    // Скидка % (M) — если было вручную, восстановим значением; иначе формула на общий дисконт
    if (prevManual && prevManual.discount !== null && prevManual.discount !== '') {
      shKP.getRange(outRow, 13).setValue(prevManual.discount);
    } else {
      shKP.getRange(outRow, 13).setFormula(`=$D$${KP_CFG.PARAM_ROW_DISCOUNT}`);
    }
    shKP.getRange(outRow, 13).setNumberFormat(KP_CFG.NUM_FMT_PCT);

    // F: цена с учетом M (скидка по строке)
    const ruPrice = kp_toRuNumberLiteral_(item.cost);
    shKP.getRange(outRow, 6).setFormula(`=${ruPrice}*(1+M${outRow}/100)`);

    // G: qty
    shKP.getRange(outRow, 7).setValue(item.qty);

    // H..K
    shKP.getRange(outRow, 8).setFormulaR1C1('=RC[-2]*RC[-1]');                // H = F*G
    shKP.getRange(outRow, 9).setFormula(`=F${outRow}*$D$${KP_CFG.PARAM_ROW_INSTALL_PCT}/100`); // I
    shKP.getRange(outRow, 10).setFormulaR1C1('=RC[-1]*RC[-3]');               // J = I*G
    shKP.getRange(outRow, 11).setFormulaR1C1('=RC[-3]+RC[-1]');               // K = H+J

    outRow++;
  }

  const lastDataRow = outRow - 1;
  const dataCount = Math.max(0, lastDataRow - dataStartRow + 1);

  // 11) Формат/границы для корзины
  if (dataCount > 0) {
    shKP.getRange(dataStartRow, 1, dataCount, KP_CFG.COL_END)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    shKP.getRange(dataStartRow, 6, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // F
    shKP.getRange(dataStartRow, 8, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // H
    shKP.getRange(dataStartRow, 9, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // I
    shKP.getRange(dataStartRow, 10, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // J
    shKP.getRange(dataStartRow, 11, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_MONEY); // K
    shKP.getRange(dataStartRow, 7, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_INT);   // G
    shKP.getRange(dataStartRow, 12, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_TEXT); // L
    shKP.getRange(dataStartRow, 14, dataCount, 1).setNumberFormat(KP_CFG.NUM_FMT_TEXT); // N
  }

  shKP.getRange(cartHeaderRow, 1, 1 + dataCount, KP_CFG.COL_END)
    .setBorder(true, true, true, true, true, true);

  // 12) Обновляем формулы итогов под новый диапазон ds..de
  const ds = dataStartRow;
  const de = lastDataRow;

  // КЛЮЧЕВО: ищем строки итогов ТОЛЬКО начиная с totalsStartRow (иначе поймаем "Доставка, руб" из настроек)
  const searchFrom = totalsStartRow;
  const searchTo = 6000;

  const rEquip =
    kp_findRowByExactTextInColA_(shKP, 'Итого за оборудование, руб', searchFrom, searchTo) ||
    kp_findRowByExactTextInColA_(shKP, 'Итого оборудование, руб', searchFrom, searchTo);

  const rDel = kp_findRowByExactTextInColA_(shKP, 'Доставка, руб', searchFrom, searchTo);
  const rInst = kp_findRowByExactTextInColA_(shKP, 'Монтаж, руб', searchFrom, searchTo);
  const rPay = kp_findRowByExactTextInColA_(shKP, 'Итого к оплате, руб', searchFrom, searchTo);

  const rVat =
    kp_findRowByExactTextInColA_(shKP, 'В том числе НДС 22%, руб', searchFrom, searchTo) ||
    kp_findRowByExactTextInColA_(shKP, 'НДС 22%, руб', searchFrom, searchTo);

  if (rEquip) shKP.getRange(rEquip, 4).setFormula(`=SUM(H${ds}:H${de})`).setNumberFormat(KP_CFG.NUM_FMT_MONEY);
  if (rDel) shKP.getRange(rDel, 4).setFormula(`=D${KP_CFG.PARAM_ROW_DELIVERY}`).setNumberFormat(KP_CFG.NUM_FMT_MONEY);
  if (rInst) shKP.getRange(rInst, 4).setFormula(`=SUM(J${ds}:J${de})`).setNumberFormat(KP_CFG.NUM_FMT_MONEY);

  if (rPay && rEquip && rInst && rDel) {
    shKP.getRange(rPay, 4).setFormula(`=D${rEquip}+D${rInst}+D${rDel}`).setNumberFormat(KP_CFG.NUM_FMT_MONEY);
  }
  if (rVat && rPay) {
    shKP.getRange(rVat, 4).setFormula(`=D${rPay}*22/122`).setNumberFormat(KP_CFG.NUM_FMT_MONEY);
  }

  // 13) UID колонка скрыта
  try { shKP.hideColumns(KP_CFG.CART_COLS.UID); } catch (e) {}

  return true;
}

/* ========================= HELPERS (общие) ========================= */

function kp_ensureSheetSize_(sh, minRows, minCols) {
  const curRows = sh.getMaxRows();
  const curCols = sh.getMaxColumns();
  if (curRows < minRows) sh.insertRowsAfter(curRows, minRows - curRows);
  if (curCols < minCols) sh.insertColumnsAfter(curCols, minCols - curCols);
}

function kp_normHeader_(s) {
  return String(s || '').replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}

function kp_colByExactHeader_(hdrNorm, exact) {
  const e = kp_normHeader_(exact);
  const i = hdrNorm.indexOf(e);
  return i >= 0 ? i + 1 : 0;
}

function kp_colByContainsAll_(hdrNorm, parts) {
  const p = parts.map(kp_normHeader_);
  for (let i = 0; i < hdrNorm.length; i++) {
    const t = hdrNorm[i];
    if (!t) continue;
    if (p.every(x => t.includes(x))) return i + 1;
  }
  return 0;
}

function kp_findRowByExactTextInColA_(sh, text, r1, r2) {
  const target = String(text || '').trim();
  const last = Math.min(r2 || sh.getLastRow(), sh.getLastRow());
  const start = Math.max(1, r1 || 1);
  for (let r = start; r <= last; r++) {
    const v = String(sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (v === target) return r;
  }
  return 0;
}

function kp_readExistingCartManualByUid_(shKP, cartHeaderRow, totalsStartRow, colUid, colNote, colDiscount, colArt) {
  const byUid = {};
  const start = cartHeaderRow + 1;
  const end = Math.max(start, totalsStartRow - 2);

  let oldRowsCount = 0;

  for (let r = start; r <= end; r++) {
    const art = String(shKP.getRange(r, colArt).getDisplayValue() || '').trim();
    const uid = String(shKP.getRange(r, colUid).getDisplayValue() || '').trim();

    if (!art && !uid) break;

    oldRowsCount++;

    if (!uid) continue;

    const note = String(shKP.getRange(r, colNote).getDisplayValue() || '').trim();

    const discountCell = shKP.getRange(r, colDiscount);
    const f = String(discountCell.getFormula() || '');
    let discount = null;
    if (!f) {
      const v = discountCell.getValue();
      if (v !== '' && v !== null && v !== undefined) discount = v;
    }

    byUid[uid] = { note, discount };
  }

  return { byUid, oldRowsCount };
}

function kp_removeImagesInRowRange_(sh, rowFrom, rowTo) {
  let imgs = [];
  try { imgs = sh.getImages(); } catch (e) { imgs = []; }
  imgs.forEach(img => {
    try {
      const r = img.getAnchorCell().getRow();
      if (r >= rowFrom && r <= rowTo) img.remove();
    } catch (e) {}
  });
}

/* ========================= SELECTED COUNT IN PRICE ========================= */

function kp_countSelectedPositionsInPrice_(shPrice) {
  const lastRow = shPrice.getLastRow();
  const lastCol = shPrice.getLastColumn();
  if (lastRow < 2) return 0;

  const cand1 = shPrice.getRange(1, 1, 1, lastCol).getValues()[0];
  const cand2 = shPrice.getRange(2, 1, 1, lastCol).getValues()[0];

  let headerRow = 1;
  let map = kp_headerMap_(cand1);
  if (!map[KP_CFG.PRICE_HEADERS.UID] || !map[KP_CFG.PRICE_HEADERS.QTY]) {
    map = kp_headerMap_(cand2);
    headerRow = 2;
  }

  if (!map[KP_CFG.PRICE_HEADERS.UID] || !map[KP_CFG.PRICE_HEADERS.QTY]) {
    const rows = kp_readPriceRows_(shPrice);
    return rows.filter(r => (Number(r.qty) || 0) > 0 && String(r.uid || '').trim() !== '').length;
  }

  const colUid = map[KP_CFG.PRICE_HEADERS.UID];
  const colQty = map[KP_CFG.PRICE_HEADERS.QTY];

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

/* ========================= RESET MODE BUILD BLOCKS ========================= */

function kp_clearSheetHardAndResetSizes_(sheet) {
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

function kp_setupKpColumns_(sh) {
  kp_ensureSheetSize_(sh, 50, KP_CFG.COL_END);

  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 360);
  sh.setColumnWidth(5, 60);
  sh.setColumnWidth(6, 150);
  sh.setColumnWidth(7, 70);
  sh.setColumnWidth(8, 160);
  sh.setColumnWidth(9, 160);
  sh.setColumnWidth(10, 160);
  sh.setColumnWidth(11, 160);
  sh.setColumnWidth(12, 260);
  sh.setColumnWidth(13, 120);
  sh.setColumnWidth(14, 60);

  for (let r = 1; r <= 7; r++) sh.setRowHeight(r, 30);
}

function kp_insertHeaderImageOverCells_(sh, fileId, startRow, startCol, endRow, endCol) {
  try { sh.getImages().forEach(img => img.remove()); } catch (e) {}

  const blob = DriveApp.getFileById(fileId).getBlob();
  const img = sh.insertImage(blob, startCol, startRow);

  const widthPx = kp_sumColumnWidths_(sh, startCol, endCol);
  const heightPx = kp_sumRowHeights_(sh, startRow, endRow);

  img.setAnchorCell(sh.getRange(startRow, startCol));
  img.setWidth(widthPx);
  img.setHeight(heightPx);
}

function kp_sumColumnWidths_(sh, c1, c2) {
  let w = 0;
  for (let c = c1; c <= c2; c++) w += sh.getColumnWidth(c);
  return w;
}
function kp_sumRowHeights_(sh, r1, r2) {
  let h = 0;
  for (let r = r1; r <= r2; r++) h += sh.getRowHeight(r);
  return h;
}

function kp_buildParamsBlock_(sh, startRow) {
  sh.getRange(startRow, 1, 1, KP_CFG.COL_END).merge().setValue('').setBackground(null);

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

    sh.getRange(r, 1, 1, 3).merge()
      .setValue(baseLabels[i])
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    const input = sh.getRange(r, 4, 1, (KP_CFG.COL_END - 3)).merge()
      .setValue('')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    if (textRows.has(r)) input.setNumberFormat(KP_CFG.NUM_FMT_TEXT);

    if (baseLabels[i] === 'Дата КП') {
      input.setNumberFormat('dd.mm.yyyy');
      sh.getRange(r, 4).setValue(new Date());
    }

    if (baseLabels[i] === 'Менеджер') {
      try {
        const ss = SpreadsheetApp.getActive();
        const refName = (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.REF) ? CFG.SHEETS.REF : 'Справочник';
        const shRef = ss.getSheetByName(refName);
        if (shRef) {
          const last = Math.max(shRef.getLastRow(), 2);
          const num = Math.max(1, last - 1);
          const listRange = shRef.getRange(2, 1, num, 1);
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(listRange, true)
            .setAllowInvalid(false)
            .build();
          input.setDataValidation(rule);
        }
      } catch (e) {}
    }

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

  const rBlank = row0 + baseLabels.length; // 17
  sh.getRange(rBlank, 1, 1, KP_CFG.COL_END).merge().setValue('').setBackground(null);

  const rSettingsTitle = rBlank + 1; // 18
  sh.getRange(rSettingsTitle, 1, 1, KP_CFG.COL_END).merge()
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
    const r = rSettingsTitle + 1 + i;

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

function kp_buildTermsBlock_(sh, startRow) {
  sh.getRange(startRow, 1, 1, KP_CFG.COL_END).merge()
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

  sh.getRange(rPrepay, 10, 1, 4).merge()
    .setValue(0.7)
    .setNumberFormat('0%')
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

function kp_buildCartBlockFromPrice_(shKP, shPrice, startRow) {
  const headerRow = startRow;

  const headers = [
    'Артикул', 'Вид 1', 'Вид 2', 'Наименование / размеры', 'Ед.',
    'Стоимость оборудования', 'Кол-во', 'Всего за оборудование',
    'Стоимость монтажа', 'Всего за монтаж', 'Итого',
    'Примечание', 'Скидка (-) / Наценка (+), %', 'UID'
  ];

  shKP.getRange(headerRow, 1, 1, KP_CFG.COL_END).setValues([headers]);
  shKP.getRange(headerRow, 1, 1, KP_CFG.COL_END)
    .setFontWeight('bold')
    .setBackground('#f5f5f5')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  try { shKP.hideColumns(KP_CFG.CART_COLS.UID); } catch (e) {}

  const priceImgByRow = kp_buildImagesByRow_(shPrice);
  const priceRows = kp_readPriceRows_(shPrice);

  // ВАЖНО: только реальные товары (есть UID) и qty > 0
  const cartRows = priceRows.filter(r => (Number(r.qty) || 0) > 0 && String(r.uid || '').trim() !== '');

  let outRow = headerRow + 1;
  const firstDataRow = outRow;

  for (const item of cartRows) {
    shKP.setRowHeight(outRow, 200);

    shKP.getRange(outRow, 1).setValue(item.art);

    kp_setViewImageFromPrice_(shKP, shPrice, priceImgByRow, outRow, 2, item.sourceRow, item.sourceColView1, item.view1);
    kp_setViewImageFromPrice_(shKP, shPrice, priceImgByRow, outRow, 3, item.sourceRow, item.sourceColView2, item.view2);

    shKP.getRange(outRow, 4).setValue(item.name);
    shKP.getRange(outRow, 5).setValue(item.unit);

    shKP.getRange(outRow, 13).setFormula(`=$D$${KP_CFG.PARAM_ROW_DISCOUNT}`).setNumberFormat(KP_CFG.NUM_FMT_PCT); // M
    shKP.getRange(outRow, 14).setValue(item.uid || '').setNumberFormat(KP_CFG.NUM_FMT_TEXT); // N

    shKP.getRange(outRow, 12).setValue('').setNumberFormat(KP_CFG.NUM_FMT_TEXT); // L note

    const ruPrice = kp_toRuNumberLiteral_(item.cost);
    shKP.getRange(outRow, 6).setFormula(`=${ruPrice}*(1+M${outRow}/100)`);

    shKP.getRange(outRow, 7).setValue(item.qty);
    shKP.getRange(outRow, 8).setFormulaR1C1('=RC[-2]*RC[-1]');
    shKP.getRange(outRow, 9).setFormula(`=F${outRow}*$D$${KP_CFG.PARAM_ROW_INSTALL_PCT}/100`);
    shKP.getRange(outRow, 10).setFormulaR1C1('=RC[-1]*RC[-3]');
    shKP.getRange(outRow, 11).setFormulaR1C1('=RC[-3]+RC[-1]');

    outRow++;
  }

  const lastDataRow = outRow - 1;
  const dataCount = Math.max(0, lastDataRow - firstDataRow + 1);

  if (dataCount > 0) {
    shKP.getRange(firstDataRow, 1, dataCount, KP_CFG.COL_END)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  shKP.getRange(headerRow, 1, 1 + dataCount, KP_CFG.COL_END)
    .setBorder(true, true, true, true, true, true);

  return { headerRow, firstDataRow, lastDataRow, afterRow: outRow };
}

function kp_buildTotalsBlock_(sh, startRow, firstDataRow, lastDataRow) {
  const r0 = startRow;
  const ds = Math.max(firstDataRow || 0, 1);
  const de = Math.max(lastDataRow || ds, ds);

  const setLine = (row, label, formula) => {
    sh.getRange(row, 1, 1, 3).merge()
      .setValue(label)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);

    sh.getRange(row, 4)
      .setFormula(formula)
      .setNumberFormat(KP_CFG.NUM_FMT_MONEY)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  };

  setLine(r0, 'Итого за оборудование, руб', `=SUM(H${ds}:H${de})`);
  setLine(r0 + 1, 'Доставка, руб', `=D${KP_CFG.PARAM_ROW_DELIVERY}`);
  setLine(r0 + 2, 'Монтаж, руб', `=SUM(J${ds}:J${de})`);
  setLine(r0 + 3, 'Итого к оплате, руб', `=D${r0}+D${r0 + 2}+D${r0 + 1}`);
  setLine(r0 + 4, 'В том числе НДС 22%, руб', `=D${r0 + 3}*22/122`);

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

function kp_buildFooterTermsBlock_(sh, startRow, termsInfo, totalsInfo) {
  sh.getRange(startRow, 1, 1, KP_CFG.COL_END).merge()
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
    sh.getRange(r, 1, 1, KP_CFG.COL_END).merge()
      .setFormula(formula)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  };

  const lineValue = (r, value) => {
    sh.getRange(r, 1, 1, KP_CFG.COL_END).merge()
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

  lineFormula(r7, `="Данное КП действительно в течение: "&${termsInfo.cells.validityDays}&" дней"`);

  return r7 + 1;
}

function kp_applyGlobalFontSize_(sh, size) {
  try {
    const maxR = Math.max(sh.getMaxRows(), sh.getLastRow(), 1);
    const maxC = Math.max(sh.getMaxColumns(), sh.getLastColumn(), 1);
    sh.getRange(1, 1, maxR, Math.min(maxC, KP_CFG.COL_END)).setFontSize(size);
  } catch (e) {}
}

function kp_applyNumberFormats_(sh, firstDataRow, lastDataRow, totalsInfo) {
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
    sh.getRange(firstDataRow, 14, n, 1).setNumberFormat(KP_CFG.NUM_FMT_TEXT); // N
  }

  if (totalsInfo && totalsInfo.cells) {
    try { sh.getRange(totalsInfo.cells.equipmentTotal).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.delivery).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.installTotal).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.toPay).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
    try { sh.getRange(totalsInfo.cells.vat).setNumberFormat(KP_CFG.NUM_FMT_MONEY); } catch (e) {}
  }
}

/* ========================= PRICE READ + IMAGES ========================= */

function kp_toRuNumberLiteral_(v) {
  if (typeof v === 'number') return String(v).replace('.', ',');
  let s = String(v ?? '').trim();
  if (!s) return '0';
  s = s.replace(/\s+/g, '');
  s = s.replace('.', ',');
  return s;
}

function kp_setViewImageFromPrice_(shKP, shPrice, priceImgByRow, targetRow, targetCol, sourceRow, sourceCol, anyVal) {
  const targetCell = shKP.getRange(targetRow, targetCol);
  const url = kp_extractImageUrl_(anyVal);

  if (url) {
    const w = shKP.getColumnWidth(targetCol);
    const h = shKP.getRowHeight(targetRow);
    const fw = Math.max(20, w - 16);
    const fh = Math.max(20, h - 10);
    targetCell.setFormula(`=IMAGE("${url}";4;${fh};${fw})`);
    return;
  }

  if (!sourceRow || !sourceCol) return;

  const img = kp_findNearestImageInRow_(priceImgByRow, sourceRow, sourceCol, 3);
  if (!img) return;

  const blob = img.getBlob();
  const inserted = shKP.insertImage(blob, targetCol, targetRow);
  inserted.setAnchorCell(targetCell);

  const w = shKP.getColumnWidth(targetCol);
  const h = shKP.getRowHeight(targetRow);

  inserted.setWidth(Math.max(10, w - 16));
  inserted.setHeight(Math.max(10, h - 10));
}

function kp_buildImagesByRow_(sheet) {
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

function kp_findNearestImageInRow_(byRow, row, desiredCol, radiusCols) {
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

function kp_readPriceRows_(shPrice) {
  const lastRow = shPrice.getLastRow();
  const lastCol = shPrice.getLastColumn();
  if (lastRow < 2) return [];

  const cand1 = shPrice.getRange(1, 1, 1, lastCol).getValues()[0];
  const cand2 = shPrice.getRange(2, 1, 1, lastCol).getValues()[0];

  let headerRowOffset = 0; // 0 => заголовок в 1-й строке
  let map = kp_headerMap_(cand1);

  if (!map[KP_CFG.PRICE_HEADERS.QTY] || !map[KP_CFG.PRICE_HEADERS.COST]) {
    map = kp_headerMap_(cand2);
    headerRowOffset = 1; // заголовок во 2-й строке
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

  const dataStartRow = headerRowOffset + 2;
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

    const url1 = kp_extractUrlFromRichText_(rich1);
    const url2 = kp_extractUrlFromRichText_(rich2);

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

function kp_extractUrlFromRichText_(rich) {
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

function kp_headerMap_(headerRowValues) {
  const m = {};
  for (let c = 0; c < headerRowValues.length; c++) {
    const h = String(headerRowValues[c] ?? '').trim();
    if (h) m[h] = c + 1; // 1-based
  }
  return m;
}

function kp_extractImageUrl_(val) {
  if (!val) return '';
  const s = String(val).trim();
  const m = s.match(/IMAGE\(\s*"([^"]+)"\s*/i);
  if (m && m[1]) return m[1];
  const h = s.match(/HYPERLINK\(\s*"([^"]+)"/i);
  if (h && h[1]) return h[1];
  if (/^https?:\/\//i.test(s)) return s;
  return '';
}