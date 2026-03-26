/**
 * КП pdf.gs (контроллер экспорта КП в PDF + запись в Журнал КП)
 *
 * ВАЖНО:
 * - НЕ объявляем глобальный const KP_EXPORT_CFG
 * - Структуру "Журнал КП" обеспечивает kp_log.gs (ensureKpLogSchema_)
 * - Экспорт PDF/Drive/UI — kp_pdf_export.gs
 */

function exportKpPdfAndLog() {
  const ss = SpreadsheetApp.getActive();
  const cfg = getKpExportCfg_();

  const sh = ss.getSheetByName(cfg.KP_SHEET);
  if (!sh) throw new Error('Не найден лист "' + cfg.KP_SHEET + '".');

  // 1) Проверка обязательных полей
  if (typeof validateRequiredKpFields_ === 'function') {
    const missing = validateRequiredKpFields_(sh) || [];
    if (missing.length) {
      SpreadsheetApp.getUi().alert(
        'Экспорт КП невозможен.\n\nНе заполнены обязательные поля:\n- ' + missing.join('\n- ')
      );
      return;
    }
  }

  // 2) Готовим/проверяем структуру Журнала КП
  if (typeof ensureKpLogSchema_ === 'function') {
    ensureKpLogSchema_(ss);
  }

  // 3) Считываем meta/cart ДО скрытий и ДО PDF
  const meta = extractMetaForLog_(sh);
  const cart = extractCartAsJson_(sh);

  // 3.1) Если в корзине нет ни одной позиции — не формируем КП
  const cartItems = (cart && cart.items && Array.isArray(cart.items)) ? cart.items : [];
  if (!cartItems.length) {
    SpreadsheetApp.getUi().alert(
      'Формирование КП остановлено.\n\nВ корзине нет выбранных позиций (количество > 0).'
    );
    return;
  }

  // 4) Проверка дубля (по условиям + только примечаниям)
  const notesSig = notesSignatureFromCartJson_(cart && cart.jsonString ? cart.jsonString : '[]');
  const dup = findDuplicateInLogExt_(ss, meta, notesSig);
  if (dup && dup.found) {
    showDuplicateDialogExt_(dup);
    return;
  }

  // 5) Скрытия для PDF
  const toHide = [];
  let discountState = null;

  try {
    if (typeof findRowByExactText_ === 'function') {
      const rSettings = findRowByExactText_(sh, cfg.EXCLUDE_BLOCK_TITLES.SETTINGS);
      if (rSettings) {
        const rPrev = rSettings - 1;
        if (rPrev >= 1 && typeof isRowBlank_ === 'function' && isRowBlank_(sh, rPrev)) {
          toHide.push(rPrev);
        }
        toHide.push(rSettings);
        toHide.push.apply(toHide, rangeRowsSafe_(rSettings + 1, rSettings + cfg.SETTINGS_ROWS_AFTER_TITLE));
      }

      const rTerms = findRowByExactText_(sh, cfg.EXCLUDE_BLOCK_TITLES.TERMS);
      if (rTerms) {
        toHide.push(rTerms);
        toHide.push.apply(toHide, rangeRowsSafe_(rTerms + 1, rTerms + cfg.TERMS_ROWS_AFTER_TITLE));
      }
    }

    const uniqueHide = Array.from(new Set(toHide))
      .filter(r => r >= 1 && r <= sh.getMaxRows())
      .sort((a, b) => a - b);

    if (uniqueHide.length && typeof hideRowsBySortedList_ === 'function') {
      hideRowsBySortedList_(sh, uniqueHide);
    }

    // скрываем колонку скидки
    discountState = (typeof hideDiscountColumnForPdf_ === 'function')
      ? hideDiscountColumnForPdf_(sh)
      : hideDiscountColumnForPdfSafe_(sh);

    SpreadsheetApp.flush();
    Utilities.sleep(300);

    // 6) PDF
    if (typeof exportSheetToPdfBlob_ !== 'function' || typeof savePdfToDriveFolder_ !== 'function') {
      throw new Error('Не найдены функции экспорта PDF (kp_pdf_export.gs).');
    }

    const pdfBlob = exportSheetToPdfBlob_(ss, sh, buildPdfFileName_(meta));
    const saved = savePdfToDriveFolder_(pdfBlob, cfg.DRIVE_FOLDER_ID);

    // 7) Журнал КП
    const logRow = buildLogRow_(meta, cart.jsonString, saved.fileUrl, saved.downloadUrl, saved.fileId);
    appendToLog_(ss, logRow);

    // 8) Диалог со ссылками
    if (typeof showLinksDialog_ === 'function') {
      showLinksDialog_(saved.fileUrl, saved.downloadUrl);
    } else {
      SpreadsheetApp.getUi().alert('PDF создан.\n\nDrive: ' + saved.fileUrl);
    }

  } finally {
    try {
      if (discountState) {
        if (typeof restoreDiscountColumnAfterPdf_ === 'function') {
          restoreDiscountColumnAfterPdf_(sh, discountState);
        } else {
          restoreDiscountColumnAfterPdfSafe_(sh, discountState);
        }
      }
    } catch (e) {}

    try {
      const uniqueHide = Array.from(new Set(toHide))
        .filter(r => r >= 1 && r <= sh.getMaxRows())
        .sort((a, b) => a - b);

      if (uniqueHide.length && typeof showRowsBySortedList_ === 'function') {
        showRowsBySortedList_(sh, uniqueHide);
      }
    } catch (e) {}
  }
}

/* ====================================================================== */
/* Конфиг                                                                 */
/* ====================================================================== */

function getKpExportCfg_() {
  const sheets = (typeof CFG !== 'undefined' && CFG.SHEETS) ? CFG.SHEETS : {};
  const ids = (typeof CFG !== 'undefined' && CFG.IDS) ? CFG.IDS : {};
  const kpExport = (typeof CFG !== 'undefined' && (CFG.KP_EXPORT || CFG.KP_PDF))
    ? (CFG.KP_EXPORT || CFG.KP_PDF)
    : {};

  return {
    KP_SHEET: sheets.KP || 'КП',
    LOG_SHEET: sheets.KP_LOG || 'Журнал КП',
    DRIVE_FOLDER_ID: kpExport.DRIVE_FOLDER_ID || ids.KP_PDF_FOLDER_ID || ids.DRIVE_FOLDER_ID || '',
    EXCLUDE_BLOCK_TITLES: {
      SETTINGS: (kpExport.EXCLUDE_BLOCK_TITLES && kpExport.EXCLUDE_BLOCK_TITLES.SETTINGS) || 'Настройки расчёта',
      TERMS: (kpExport.EXCLUDE_BLOCK_TITLES && kpExport.EXCLUDE_BLOCK_TITLES.TERMS) || 'Условия и сроки поставки (изменяемые)'
    },
    SETTINGS_ROWS_AFTER_TITLE: kpExport.SETTINGS_ROWS_AFTER_TITLE || 3,
    TERMS_ROWS_AFTER_TITLE: kpExport.TERMS_ROWS_AFTER_TITLE || 4,
    CART_HEADER: kpExport.CART_HEADER || 'Артикул'
  };
}

function rangeRowsSafe_(fromRow, toRow) {
  const out = [];
  if (!fromRow || !toRow) return out;
  for (let r = fromRow; r <= toRow; r++) out.push(r);
  return out;
}

/* ====================================================================== */
/* Дубликаты: сравнение по условиям + только примечания из Позиции(JSON)   */
/* ====================================================================== */

function findDuplicateInLogExt_(ss, meta, notesSig) {
  const cfg = getKpExportCfg_();
  const log = ss.getSheetByName(cfg.LOG_SHEET);
  if (!log) return { found: false };

  const lastRow = log.getLastRow();
  const lastCol = log.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { found: false };

  const hdr = log.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(_norm_);
  const col = (name) => hdr.indexOf(_norm_(name)) + 1;

  const cKp    = col('КП №');
  const cCust  = col('Заказчик');
  const cAddr  = col('Адрес заказчика');
  const cToPay = col('Итого к оплате, руб');
  const cMain  = col('Срок (Основное)');
  const cEco   = col('Срок (ЭКО)');
  const cValid = col('КП действительно, дней');
  const cJson  = col('Позиции (JSON)');
  const cPdf   = col('PDF URL (Drive)');

  if (!cKp || !cCust || !cAddr || !cToPay || !cMain || !cEco || !cValid || !cJson) {
    return { found: false };
  }

  const n = lastRow - 1;

  const kpVals    = log.getRange(2, cKp, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const custVals  = log.getRange(2, cCust, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const addrVals  = log.getRange(2, cAddr, n, 1).getDisplayValues().map(r => _normText_(r[0]));
  const mainVals  = log.getRange(2, cMain, n, 1).getDisplayValues().map(r => _normText_(r[0]));
  const ecoVals   = log.getRange(2, cEco, n, 1).getDisplayValues().map(r => _normText_(r[0]));
  const validVals = log.getRange(2, cValid, n, 1).getValues().map(r => r[0]);
  const payVals   = log.getRange(2, cToPay, n, 1).getValues().map(r => r[0]);
  const jsonVals  = log.getRange(2, cJson, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const pdfVals   = cPdf ? log.getRange(2, cPdf, n, 1).getDisplayValues().map(r => String(r[0] || '').trim()) : [];

  const keyKp    = String(meta.kpNo || '').trim();
  const keyCust  = String(meta.customer || '').trim();
  const keyAddr  = _normText_(meta.customerAddr || '');
  const keyMain  = _normText_(meta.mainLead || '');
  const keyEco   = _normText_(meta.ecoLead || '');
  const keyValid = _round0_(_toNum_(meta.validDays));
  const keyPay   = _round2_(_toNum_(meta.toPay));
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

    const vNotes = notesSignatureFromCartJson_(jsonVals[i]);
    if (vNotes !== keyNotes) continue;

    return {
      found: true,
      row: i + 2,
      kpNo: keyKp,
      customer: keyCust,
      customerAddr: meta.customerAddr || '',
      toPay: keyPay,
      pdfUrl: (pdfVals[i] || '')
    };
  }

  return { found: false };
}

function notesSignatureFromCartJson_(jsonString) {
  let arr = [];
  try {
    arr = JSON.parse(jsonString || '[]');
    if (!Array.isArray(arr)) arr = [];
  } catch (e) {
    arr = [];
  }

  const pairs = arr.map(item => {
    const art = _normText_(item && (item.art || item.article || ''));
    const note = _normText_(item && (item.note || item['Примечание'] || ''));
    return art + '|' + note;
  });

  pairs.sort();
  return pairs.join('||');
}

function showDuplicateDialogExt_(dup) {
  const url = dup.pdfUrl || '';
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.45">
      <div style="font-size:16px;font-weight:700;margin-bottom:8px;">Дубликат КП</div>
      <div>Такое КП уже сформировано (совпали условия + примечания).</div>
      <div style="margin-top:10px;"><b>КП №:</b> ${_esc_(dup.kpNo)}</div>
      <div><b>Заказчик:</b> ${_esc_(dup.customer)}</div>
      <div><b>Адрес:</b> ${_esc_(dup.customerAddr)}</div>
      <div><b>Итого к оплате:</b> ${_esc_(String(dup.toPay))}</div>
      <div><b>Строка в журнале:</b> ${dup.row}</div>
      ${url ? `<div style="margin-top:12px;"><a href="${url}" target="_blank">Открыть PDF (Drive)</a></div>` : `<div style="margin-top:12px;color:#666;">Ссылка на PDF в журнале не найдена.</div>`}
      <div style="margin-top:12px;color:#a33;"><b>Новая выгрузка не выполнена</b>, чтобы избежать дублирования.</div>
    </div>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(310),
    'Дубликат КП'
  );
}

/* ====================================================================== */
/* Скрытие / восстановление колонки скидки                                 */
/* ====================================================================== */

function hideDiscountColumnForPdfSafe_(sh) {
  const state = { wasHidden: false, col: 0 };

  const candidates = [];

  // 1. Пытаемся найти по заголовку корзины
  const headerRow = findCartHeaderRowByA_(sh, 'Артикул');
  if (headerRow) {
    const maxCol = Math.min(60, sh.getLastColumn());
    const hdr = sh.getRange(headerRow, 1, 1, maxCol).getDisplayValues()[0].map(_norm_);

    for (let c = 0; c < hdr.length; c++) {
      const t = hdr[c];
      if (t && t.indexOf('скидк') >= 0 && t.indexOf('%') >= 0) {
        candidates.push(c + 1);
        break;
      }
    }
  }

  // 2. Из KP.gs
  try {
    if (
      typeof KP_CFG !== 'undefined' &&
      KP_CFG.CART_COLS &&
      KP_CFG.CART_COLS.DISCOUNT_PCT
    ) {
      candidates.push(Number(KP_CFG.CART_COLS.DISCOUNT_PCT));
    }
  } catch (e) {}

  // 3. Фолбэк — колонка M
  candidates.push(13);

  const uniq = [];
  const seen = {};
  for (let i = 0; i < candidates.length; i++) {
    const col = Number(candidates[i] || 0);
    if (!col || seen[col]) continue;
    if (col < 1 || col > sh.getMaxColumns()) continue;
    seen[col] = true;
    uniq.push(col);
  }

  for (let j = 0; j < uniq.length; j++) {
    const tryCol = uniq[j];

    try {
      state.col = tryCol;
      state.wasHidden = sh.isColumnHiddenByUser(tryCol);

      if (!state.wasHidden) {
        sh.hideColumns(tryCol);
        SpreadsheetApp.flush();
        Utilities.sleep(250);
      }

      if (sh.isColumnHiddenByUser(tryCol)) {
        return state;
      }
    } catch (e) {
      // идём дальше
    }
  }

  return { wasHidden: false, col: 0 };
}

function restoreDiscountColumnAfterPdfSafe_(sh, state) {
  if (!state || !state.col) return;
  if (!state.wasHidden) {
    try {
      sh.showColumns(state.col);
      SpreadsheetApp.flush();
    } catch (e) {}
  }
}

function findCartHeaderRowByA_(sh, anchorText) {
  const needle = _norm_(anchorText);
  const last = Math.min(800, sh.getLastRow());
  for (let r = 1; r <= last; r++) {
    const v = _norm_(sh.getRange(r, 1).getDisplayValue());
    if (v === needle) return r;
  }
  return 0;
}

/* ====================================================================== */
/* Утилиты                                                                 */
/* ====================================================================== */

function _norm_(s) {
  return String(s || '')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _normText_(s) {
  return String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _toNum_(v) {
  if (typeof toNumber_ === 'function') return toNumber_(v);
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v).replace(/\s+/g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function _round2_(n) {
  return Math.round((_toNum_(n) + Number.EPSILON) * 100) / 100;
}

function _round0_(n) {
  return Math.round(_toNum_(n));
}

function _esc_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}