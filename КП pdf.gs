/**
 * КП pdf.gs (контроллер)
 *
 * ДОРАБОТКА:
 * 1) Расширенная проверка дублей ПЕРЕД формированием PDF и записью в журнал.
 * 2) Колонка "Скидка (-) / Наценка (+), %" в корзине НЕ должна попадать в PDF:
 *    перед экспортом временно скрываем колонку M (13) + ищем по заголовку.
 * 3) UID-колонка (N, 14) тоже НЕ должна попадать в PDF — скрываем аналогично.
 */

const KP_EXPORT_CFG = {
  KP_SHEET: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.KP) ? CFG.SHEETS.KP : 'КП',
  LOG_SHEET: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.KP_LOG) ? CFG.SHEETS.KP_LOG : 'Журнал КП',
  DRIVE_FOLDER_ID: (typeof CFG !== 'undefined' && CFG.IDS && CFG.IDS.DRIVE_FOLDER_ID) ? CFG.IDS.DRIVE_FOLDER_ID : '1o8lqVv3DlUe4e3bMpWKvNZncf5r_2xDD',

  EXCLUDE_BLOCK_TITLES: {
    SETTINGS: 'Настройки расчёта',
    TERMS: 'Условия и сроки поставки (изменяемые)',
  },
  SETTINGS_ROWS_AFTER_TITLE: 3,
  TERMS_ROWS_AFTER_TITLE: 4,

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
    'Drive File ID',
  ],
};

function exportKpPdfAndLog() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(KP_EXPORT_CFG.KP_SHEET);
  if (!sh) throw new Error('Не найден лист "КП".');

  const missing = validateRequiredKpFields_(sh);
  if (missing && missing.length) {
    SpreadsheetApp.getUi().alert(
      'Экспорт КП невозможен.\n\nНе заполнены обязательные поля:\n- ' + missing.join('\n- ')
    );
    return;
  }

  // Считываем meta+cart ДО скрытий и ДО PDF (для дублей)
  const meta = extractMetaForLog_(sh);
  const cart = extractCartAsJson_(sh);
  const notesSig = notesSignatureFromCartJson_(cart && cart.jsonString ? cart.jsonString : '[]');

  const dup = findDuplicateInLogExt_(ss, meta, notesSig);
  if (dup && dup.found) {
    showDuplicateDialogExt_(dup);
    return;
  }

  const toHide = [];
  let colsState = null;

  try {
    // исключаем блоки из PDF
    const rSettings = findRowByExactText_(sh, KP_EXPORT_CFG.EXCLUDE_BLOCK_TITLES.SETTINGS);
    if (rSettings) {
      const rPrev = rSettings - 1;
      if (rPrev >= 1 && isRowBlank_(sh, rPrev)) toHide.push(rPrev);
      toHide.push(rSettings);
      toHide.push(...rangeRows_(rSettings + 1, rSettings + KP_EXPORT_CFG.SETTINGS_ROWS_AFTER_TITLE));
    }

    const rTerms = findRowByExactText_(sh, KP_EXPORT_CFG.EXCLUDE_BLOCK_TITLES.TERMS);
    if (rTerms) {
      toHide.push(rTerms);
      toHide.push(...rangeRows_(rTerms + 1, rTerms + KP_EXPORT_CFG.TERMS_ROWS_AFTER_TITLE));
    }

    const uniqueHide = Array.from(new Set(toHide)).filter(r => r >= 1 && r <= sh.getMaxRows());
    if (uniqueHide.length) hideRowsBySortedList_(sh, uniqueHide);

    // скрываем колонку скидки + UID в корзине
    colsState = hideServiceColumnsForPdf_(sh);

    // PDF
    const pdfBlob = exportSheetToPdfBlob_(ss, sh, buildPdfFileName_(meta));
    const saved = savePdfToDriveFolder_(pdfBlob, KP_EXPORT_CFG.DRIVE_FOLDER_ID);

    // log
    const logRow = buildLogRow_(meta, cart.jsonString, saved.fileUrl, saved.downloadUrl, saved.fileId);
    appendToLog_(ss, logRow);

    showLinksDialog_(saved.fileUrl, saved.downloadUrl);
  } finally {
    try { if (colsState) restoreServiceColumnsAfterPdf_(sh, colsState); } catch (e) {}
    try {
      const uniqueHide = Array.from(new Set(toHide)).filter(r => r >= 1 && r <= sh.getMaxRows());
      if (uniqueHide.length) showRowsBySortedList_(sh, uniqueHide);
    } catch (e) {}
  }
}

/* ========================== Дубликаты: сравнение по условиям + ТОЛЬКО примечания ========================== */

function findDuplicateInLogExt_(ss, meta, notesSig) {
  const log = ss.getSheetByName(KP_EXPORT_CFG.LOG_SHEET);
  if (!log) return { found: false };

  const lastRow = log.getLastRow();
  const lastCol = log.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { found: false };

  const hdr = log.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(_norm_);
  const col = (name) => hdr.indexOf(_norm_(name)) + 1;

  const cKp = col('КП №');
  const cCust = col('Заказчик');
  const cAddr = col('Адрес заказчика');
  const cToPay = col('Итого к оплате, руб');
  const cMain = col('Срок (Основное)');
  const cEco = col('Срок (ЭКО)');
  const cValid = col('КП действительно, дней');
  const cJson = col('Позиции (JSON)');
  const cPdf = col('PDF URL (Drive)');

  if (!cKp || !cCust || !cAddr || !cToPay || !cMain || !cEco || !cValid || !cJson) return { found: false };

  const n = lastRow - 1;

  const kpVals = log.getRange(2, cKp, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const custVals = log.getRange(2, cCust, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const addrVals = log.getRange(2, cAddr, n, 1).getDisplayValues().map(r => _normText_(r[0]));
  const mainVals = log.getRange(2, cMain, n, 1).getDisplayValues().map(r => _normText_(r[0]));
  const ecoVals = log.getRange(2, cEco, n, 1).getDisplayValues().map(r => _normText_(r[0]));
  const validVals = log.getRange(2, cValid, n, 1).getValues().map(r => r[0]);
  const payVals = log.getRange(2, cToPay, n, 1).getValues().map(r => r[0]);
  const jsonVals = log.getRange(2, cJson, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const pdfVals = cPdf ? log.getRange(2, cPdf, n, 1).getDisplayValues().map(r => String(r[0] || '').trim()) : [];

  const keyKp = String(meta.kpNo || '').trim();
  const keyCust = String(meta.customer || '').trim();
  const keyAddr = _normText_(meta.customerAddr || '');
  const keyMain = _normText_(meta.mainLead || '');
  const keyEco = _normText_(meta.ecoLead || '');
  const keyValid = _round0_(_toNum_(meta.validDays));
  const keyPay = _round2_(_toNum_(meta.toPay));
  const keyNotes = String(notesSig || '');

  for (let i = n - 1; i >= 0; i--) {
    if (String(kpVals[i] || '').trim() !== keyKp) continue;
    if (String(custVals[i] || '').trim() !== keyCust) continue;
    if (addrVals[i] !== keyAddr) continue;
    if (mainVals[i] !== keyMain) continue;
    if (ecoVals[i] !== keyEco) continue;

    const vValid = _round0_(_toNum_(validVals[i]));
    if (vValid !== keyValid) continue;

    const vPay = _round2_(_toNum_(payVals[i]));
    if (Math.abs(vPay - keyPay) > 0.01) continue;

    // Сравниваем только примечания (из JSON)
    const vNotes = notesSignatureFromCartJson_(jsonVals[i]);
    if (vNotes !== keyNotes) continue;

    return {
      found: true,
      row: i + 2,
      kpNo: keyKp,
      customer: keyCust,
      customerAddr: meta.customerAddr || '',
      toPay: keyPay,
      pdfUrl: (pdfVals[i] || ''),
    };
  }

  return { found: false };
}

function showDuplicateDialogExt_(dup) {
  const url = dup.pdfUrl || '';
  const html = `
    Такое КП уже сформировано (совпали условия + примечания)
    <br><br>
    <b>КП №:</b> ${_esc_(dup.kpNo)}<br>
    <b>Заказчик:</b> ${_esc_(dup.customer)}<br>
    <b>Адрес:</b> ${_esc_(dup.customerAddr)}<br>
    <b>Итого к оплате:</b> ${_esc_(String(dup.toPay))}<br><br>
    <b>Запись в журнале:</b> строка ${dup.row}<br><br>
    ${url ? `<a href="${url}" target="_blank">Открыть PDF (Drive)</a><br><br>` : 'Ссылка на PDF в журнале не найдена.<br><br>'}
    Новая выгрузка не выполнена, чтобы избежать дублирования.
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(310),
    'Дубликат КП'
  );
}

/**
 * Вытаскиваем "сигнатуру примечаний" из JSON позиций:
 * берём пары art|note, нормализуем пробелы/регистр, сортируем по art.
 * UID игнорируем (это нормально).
 */
function notesSignatureFromCartJson_(jsonStr) {
  const raw = String(jsonStr || '').trim();
  if (!raw) return '';

  try {
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return '';

    const pairs = arr.map(it => {
      const art = String(it.art || '').trim();
      const note = String(it.note ?? it['Примечание'] ?? '').trim();
      return art + '|' + _normNote_(note);
    });

    pairs.sort();
    return pairs.join('\n');
  } catch (e) {
    return '';
  }
}

function _norm_(s) {
  return String(s || '').replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}
function _normText_(s) { return _norm_(s); }
function _normNote_(s) { return String(s || '').replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); }

function _toNum_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/\s+/g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
function _round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function _round0_(n) { return Math.round(Number(n) || 0); }
function _esc_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* ========================== Скрытие служебных колонок в PDF ========================== */

function hideServiceColumnsForPdf_(sh) {
  const state = { hidden: [] };

  const hideCol = (col) => {
    if (!col || col < 1 || col > sh.getLastColumn()) return;
    const wasHidden = sh.isColumnHiddenByUser(col);
    if (!wasHidden) sh.hideColumns(col);
    state.hidden.push({ col, wasHidden });
  };

  // 1) фикс: скидка в M (13) — как было
  hideCol(13);

  // 2) фикс: UID в N (14) — новый
  hideCol(14);

  // 3) доп. поиск по заголовку корзины (если вдруг колонки переместятся)
  try {
    const headerRow = findCartHeaderRowByA_(sh, KP_EXPORT_CFG.CART_HEADER);
    if (headerRow) {
      const maxCol = Math.min(60, sh.getLastColumn());
      const hdrs = sh.getRange(headerRow, 1, 1, maxCol).getDisplayValues()[0].map(_norm_);

      for (let c = 0; c < hdrs.length; c++) {
        const t = hdrs[c];
        if (!t) continue;
        if (t.includes('скидка') && t.includes('%')) hideCol(c + 1);
        if (t === 'uid') hideCol(c + 1);
      }
    }
  } catch (e) {}

  try { SpreadsheetApp.flush(); } catch (e) {}
  try { Utilities.sleep(600); } catch (e) {}

  return state;
}

function restoreServiceColumnsAfterPdf_(sh, state) {
  if (!state || !state.hidden) return;

  for (const it of state.hidden) {
    try {
      if (it && it.col && !it.wasHidden) sh.showColumns(it.col);
    } catch (e) {}
  }

  try { SpreadsheetApp.flush(); } catch (e) {}
}

function findCartHeaderRowByA_(sh, anchorText) {
  const anchor = _norm_(anchorText);
  const last = Math.min(600, sh.getLastRow());
  for (let r = 1; r <= last; r++) {
    const v = _norm_(sh.getRange(r, 1).getDisplayValue());
    if (v === anchor) return r;
  }
  return 0;
}