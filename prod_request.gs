/**
 * prod_request.gs
 * Полная исправленная версия:
 * - Журнал заявок на производство пишется ПО ЗАГОЛОВКАМ, без смещения
 * - Временный лист для PDF удаляется после формирования файла
 * - После создания заявки показывается диалог со ссылкой на скачивание PDF
 * - Шапка заявки компактная, поля идут друг под другом
 * - Подтягиваются картинки "Вид 1" и "Вид 2" (если есть URL в БД/Прайсе)
 *
 * ВАЖНО:
 * - onOpen() не меняем
 * - main.gs может вызывать createProductionOrderFromSelectedKp()
 */

/* ========================================================================== */
/* CONFIG                                                                     */
/* ========================================================================== */

var PROD_REQ_CFG = {
  SHEET_KP_LOG: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.KP_LOG) ? CFG.SHEETS.KP_LOG : 'Журнал КП',
  SHEET_PROD_LOG: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.PROD_REQUEST_LOG) ? CFG.SHEETS.PROD_REQUEST_LOG : 'Журнал заявок на производство',
  SHEET_PRICE: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.PRICE) ? CFG.SHEETS.PRICE : 'Прайс',
  SHEET_DB: (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.DB) ? CFG.SHEETS.DB : 'БД Оборудования',

  PDF_FOLDER_ID:
    (typeof CFG !== 'undefined' && CFG.IDS && (CFG.IDS.KP_PDF_FOLDER_ID || CFG.IDS.DRIVE_FOLDER_ID))
      ? (CFG.IDS.KP_PDF_FOLDER_ID || CFG.IDS.DRIVE_FOLDER_ID)
      : '',

  TEMP_SHEET_PREFIX: '_tmp_prod_request_',

  FORM_COLS: 7,
  COL_WIDTHS: [95, 95, 95, 250, 50, 55, 135], // A:G
  DEFAULT_ROW_HEIGHT: 22,
  TITLE_FONT_SIZE: 16,
  BODY_FONT_SIZE: 10,
  HEADER_FONT_SIZE: 11,

  DATA_ROW_HEIGHT: 92,
  IMG_SIZE: 78,

  PDF_PORTRAIT: false, // false = landscape
  PDF_MARGINS: {
    top: 0.20,
    bottom: 0.20,
    left: 0.20,
    right: 0.20
  }
};

/* ========================================================================== */
/* ENTRY POINTS                                                               */
/* ========================================================================== */

function createProductionOrderFromSelectedKp() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const kpLog = ensureKpLogSchema_(ss);
  const prodLog = ensureProductionRequestLogSchema_(ss);

  const activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getName() !== kpLog.getName()) {
    ui.alert(
      'Для создания заявки на производство перейдите на лист "' + kpLog.getName() + '" и выделите нужную строку.'
    );
    return;
  }

  const row = activeSheet.getActiveRange() ? activeSheet.getActiveRange().getRow() : 0;
  if (row < 2) {
    ui.alert('Выберите строку в "Журнал КП" (не шапку).');
    return;
  }

  const kpRow = getKpLogRowAsObject_(kpLog, row);

  const kpDriveFileId = String(_prodPickKpField_(kpRow, ['Drive File ID', 'КП Drive File ID']) || '').trim();
  const kpNo = String(kpRow['КП №'] || '').trim();
  const customer = String(kpRow['Заказчик'] || '').trim();
  const status = String(kpRow['Статус'] || '').trim() || 'Новая';

  if (!kpDriveFileId) {
    ui.alert(
      'Нельзя создать заявку на производство.\n\nВ выбранной строке "Журнал КП" не заполнен "Drive File ID".'
    );
    return;
  }

  if (!kpNo) {
    ui.alert('Нельзя создать заявку на производство: не заполнен "КП №" в журнале.');
    return;
  }

  if (!customer) {
    ui.alert('Нельзя создать заявку на производство: не заполнен "Заказчик" в журнале.');
    return;
  }

  if (status === 'Отправлена в производство') {
    ui.alert('По этой записи уже создана заявка на производство.');
    return;
  }

  if (status === 'Аннулирована') {
    ui.alert('Выбранная запись имеет статус "Аннулирована". Создание заявки запрещено.');
    return;
  }

  const existingSent = findAnotherSentProductionByKpAndCustomer_(kpLog, kpNo, customer, kpDriveFileId);
  if (existingSent) {
    ui.alert(
      'Заявка на производство уже создана по дублю этого КП.\n\n' +
      'Строка журнала КП: ' + existingSent.row + '\n' +
      'Статус: Отправлена в производство\n\n' +
      'Для текущей строки создание заявки запрещено.'
    );
    return;
  }

  const requestNo = _prodPromptRequired_(
    'Номер заявки на производство',
    'Введите номер заявки на производство:',
    ''
  );
  if (requestNo === null) return;

  const invoiceNo = _prodPromptRequired_(
    'Счет на оплату',
    'Введите номер счета на оплату:',
    ''
  );
  if (invoiceNo === null) return;

  const plannedShipDate = _prodPromptRequired_(
    'Плановая дата отгрузки',
    'Введите плановую дату отгрузки (например 15.05.26):',
    ''
  );
  if (plannedShipDate === null) return;

  const previewText = buildProdRequestPreviewText_(kpRow, requestNo, invoiceNo, plannedShipDate);
  const confirm = ui.alert('Создать заявку на производство', previewText + '\n\nПродолжить?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const dupProd = findProdRequestDuplicate_(prodLog, requestNo, kpDriveFileId);
  if (dupProd) {
    ui.alert(
      'Такая заявка уже есть в "Журнал заявок на производство".\n\n' +
      'Строка: ' + dupProd.row + '\n' +
      'Номер заявки: ' + requestNo
    );
    return;
  }

  const now = new Date();
  const reqData = buildProductionRequestData_(ss, kpRow, requestNo, invoiceNo, plannedShipDate, now);

  let tempSheet = null;
  let saved = null;

  try {
    tempSheet = buildProductionRequestTempSheet_(ss, reqData);

    SpreadsheetApp.flush();
    Utilities.sleep(500);

    saved = exportProductionRequestPdf_(ss, tempSheet, reqData.fileName);

    const prodRow = buildProdRequestLogRow_(kpRow, requestNo, invoiceNo, plannedShipDate, now, saved);
    const prodRowNum = appendProductionRequestToLog_(prodLog, prodRow);

    setKpLogStatusByRow_(kpLog, row, 'Отправлена в производство');

    const cancelledCount = cancelDuplicateKpRows_(kpLog, {
      kpNo: kpNo,
      customer: customer,
      exceptDriveFileId: kpDriveFileId,
      keepRow: row
    });

    ss.setActiveSheet(kpLog);

    showProdRequestLinksDialog_(saved.fileUrl, saved.downloadUrl, {
      prodRowNum: prodRowNum,
      requestNo: requestNo,
      cancelledCount: cancelledCount
    });

  } finally {
    try {
      if (tempSheet && ss.getSheetByName(tempSheet.getName())) {
        ss.deleteSheet(tempSheet);
      }
    } catch (e) {}
  }
}

/**
 * Алиас для совместимости
 */
function createProductionRequestFromSelectedKp() {
  return createProductionOrderFromSelectedKp();
}

/* ========================================================================== */
/* DATA BUILD                                                                 */
/* ========================================================================== */

function buildProductionRequestData_(ss, kpRow, requestNo, invoiceNo, plannedShipDate, createdAt) {
  const requestDateTime = createdAt || new Date();
  const requestDateOnly = _prodFormatDateOnly_(requestDateTime);

  const kpNo = String(kpRow['КП №'] || '').trim();
  const kpDateRaw = kpRow['Дата КП'] || '';
  const kpDateOnly = _prodFormatDateOnly_(kpDateRaw);

  const customer = String(kpRow['Заказчик'] || '').trim();
  const customerAddr = String(kpRow['Адрес заказчика'] || '').trim();
  const contractNo = String(kpRow['№ договора'] || '').trim();
  const manager = String(kpRow['Менеджер'] || '').trim();
  const phone = String(kpRow['Телефон'] || '').trim();

  const docsBase =
    'Договора №: ' + contractNo +
    '; Коммерческое предложение № ' + kpNo + ' от ' + kpDateOnly +
    '; Счет на оплату: ' + invoiceNo;

  const items = _prodParseItemsJson_(_prodPickKpField_(kpRow, ['Позиции (JSON)']) || '[]');
  const catalogIndex = _prodBuildCatalogIndex_(ss);
  const enrichedItems = _prodEnrichItemsFromCatalog_(items, catalogIndex);

  const fileName =
    'Заявка на производство № ' + requestNo +
    ' от ' + requestDateOnly +
    ' - ' + _prodSafeFilePart_(customer) +
    ' - КП ' + kpNo;

  return {
    createdAt: requestDateTime,
    requestNo: requestNo,
    requestDateOnly: requestDateOnly,
    title: 'ЗАЯВКА НА ПРОИЗВОДСТВО № ' + requestNo + ' от ' + requestDateOnly,

    customer: customer,
    customerAddr: customerAddr,
    docsBase: docsBase,
    managerPhone: manager + ' / ' + phone,
    plannedShipDate: plannedShipDate,

    kpNo: kpNo,
    kpDateOnly: kpDateOnly,
    contractNo: contractNo,
    invoiceNo: invoiceNo,

    items: enrichedItems,
    fileName: fileName
  };
}

function _prodParseItemsJson_(jsonString) {
  try {
    const arr = JSON.parse(jsonString || '[]');
    if (Array.isArray(arr)) return arr;
  } catch (e) {}
  return [];
}

function _prodEnrichItemsFromCatalog_(items, catalogIndex) {
  return (items || []).map(function (item) {
    const uid = String(item.uid || '').trim();
    const art = String(item.art || '').trim();

    let src = null;
    if (uid && catalogIndex.byUid[uid]) src = catalogIndex.byUid[uid];
    if (!src && art && catalogIndex.byArt[art]) src = catalogIndex.byArt[art];

    return {
      uid: uid,
      art: art,
      name: (src && src.name) ? src.name : String(item.name || '').trim(),
      unit: (src && src.unit) ? src.unit : String(item.unit || '').trim(),
      qty: _prodToNumber_(item.qty),
      note: String(item.note || '').trim(),
      view1: (src && src.view1) ? src.view1 : '',
      view2: (src && src.view2) ? src.view2 : ''
    };
  }).filter(function (x) {
    return x.art || x.name || x.qty > 0;
  });
}

/* ========================================================================== */
/* TEMP SHEET                                                                 */
/* ========================================================================== */

function buildProductionRequestTempSheet_(ss, data) {
  const sheetName = PROD_REQ_CFG.TEMP_SHEET_PREFIX + new Date().getTime();
  const sh = ss.insertSheet(sheetName);

  _prodResetSheet_(sh);
  _prodEnsureSheetSize_(sh, Math.max(60, data.items.length * 3 + 20), PROD_REQ_CFG.FORM_COLS);
  _prodSetupSheetLayout_(sh);

  // Title
  sh.getRange(1, 1, 1, 7).merge()
    .setValue(data.title)
    .setFontWeight('bold')
    .setFontSize(PROD_REQ_CFG.TITLE_FONT_SIZE)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 28);

  // Header block
  let r = 3;
  _prodSetMetaRow_(sh, r++, 'Заказчик:', data.customer, 22);
  _prodSetMetaRow_(sh, r++, 'Адрес заказчика:', data.customerAddr, 22);
  _prodSetMetaRow_(sh, r++, 'Документы основание:', data.docsBase, 38);
  _prodSetMetaRow_(sh, r++, 'Менеджер / телефон:', data.managerPhone, 22);
  _prodSetMetaRow_(sh, r++, 'Плановая дата отгрузки:', data.plannedShipDate, 22);

  r += 1;

  // Table header
  const headerRow = r;
  const headers = [
    'Артикул',
    'Вид 1',
    'Вид 2',
    'Наименование / размеры',
    'Ед.',
    'Кол-во',
    'Примечание'
  ];

  sh.getRange(headerRow, 1, 1, 7).setValues([headers]);
  sh.getRange(headerRow, 1, 1, 7)
    .setFontWeight('bold')
    .setFontSize(PROD_REQ_CFG.HEADER_FONT_SIZE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true);
  sh.setRowHeight(headerRow, 30);

  let row = headerRow + 1;
  data.items.forEach(function (item) {
    sh.setRowHeight(row, PROD_REQ_CFG.DATA_ROW_HEIGHT);

    sh.getRange(row, 1).setValue(item.art || '');
    sh.getRange(row, 4).setValue(item.name || '');
    sh.getRange(row, 5).setValue(item.unit || '');
    sh.getRange(row, 6).setValue(item.qty || '');
    sh.getRange(row, 7).setValue(item.note || '');

    _prodSetImageFormula_(sh, row, 2, item.view1);
    _prodSetImageFormula_(sh, row, 3, item.view2);

    sh.getRange(row, 1, 1, 7)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true);

    sh.getRange(row, 1).setHorizontalAlignment('left');
    sh.getRange(row, 4).setHorizontalAlignment('left');
    sh.getRange(row, 5).setHorizontalAlignment('center');
    sh.getRange(row, 6).setHorizontalAlignment('center');
    sh.getRange(row, 7).setHorizontalAlignment('left');

    row++;
  });

  const lastDataRow = Math.max(headerRow, row - 1);
  sh.getRange(1, 1, lastDataRow, 7).setFontSize(PROD_REQ_CFG.BODY_FONT_SIZE);
  sh.getRange(1, 1, 1, 7).setFontSize(PROD_REQ_CFG.TITLE_FONT_SIZE);
  sh.getRange(headerRow, 1, 1, 7).setFontSize(PROD_REQ_CFG.HEADER_FONT_SIZE);

  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.setHiddenGridlines(false);
  sh.setActiveSelection('A1');

  return sh;
}

function _prodSetMetaRow_(sh, row, label, value, height) {
  sh.getRange(row, 1, 1, 2).merge()
    .setValue(label)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sh.getRange(row, 3, 1, 5).merge()
    .setValue(value || '')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sh.setRowHeight(row, height || 22);
}

function _prodSetImageFormula_(sh, row, col, url) {
  const cell = sh.getRange(row, col);
  const cleanUrl = String(url || '').trim();
  if (!cleanUrl) {
    cell.setValue('');
    return;
  }

  cell.setFormula(
    '=IMAGE("' + cleanUrl + '";4;' + PROD_REQ_CFG.IMG_SIZE + ';' + PROD_REQ_CFG.IMG_SIZE + ')'
  );
  cell.setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function _prodResetSheet_(sh) {
  const maxR = Math.max(sh.getMaxRows(), 1);
  const maxC = Math.max(sh.getMaxColumns(), 1);

  try {
    sh.getImages().forEach(function (img) { img.remove(); });
  } catch (e) {}

  try {
    const f = sh.getFilter();
    if (f) f.remove();
  } catch (e) {}

  const rng = sh.getRange(1, 1, maxR, maxC);
  try { rng.breakApart(); } catch (e) {}

  sh.clear();
  try {
    sh.showRows(1, maxR);
    sh.showColumns(1, maxC);
  } catch (e) {}

  sh.setRowHeights(1, maxR, PROD_REQ_CFG.DEFAULT_ROW_HEIGHT);
}

function _prodEnsureSheetSize_(sh, minRows, minCols) {
  const curRows = sh.getMaxRows();
  const curCols = sh.getMaxColumns();
  if (curRows < minRows) sh.insertRowsAfter(curRows, minRows - curRows);
  if (curCols < minCols) sh.insertColumnsAfter(curCols, minCols - curCols);
}

function _prodSetupSheetLayout_(sh) {
  for (var c = 1; c <= PROD_REQ_CFG.COL_WIDTHS.length; c++) {
    sh.setColumnWidth(c, PROD_REQ_CFG.COL_WIDTHS[c - 1]);
  }
}

/* ========================================================================== */
/* CATALOG / IMAGES                                                           */
/* ========================================================================== */

function _prodBuildCatalogIndex_(ss) {
  const out = { byUid: {}, byArt: {} };

  const candidates = [
    ss.getSheetByName(PROD_REQ_CFG.SHEET_DB),
    ss.getSheetByName(PROD_REQ_CFG.SHEET_PRICE)
  ].filter(Boolean);

  candidates.forEach(function (sh) {
    const rows = _prodReadCatalogRows_(sh);
    rows.forEach(function (r) {
      if (r.uid && !out.byUid[r.uid]) out.byUid[r.uid] = r;
      if (r.art && !out.byArt[r.art]) out.byArt[r.art] = r;
    });
  });

  return out;
}

function _prodReadCatalogRows_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  const header1 = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const header2 = lastRow >= 2 ? sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] : [];
  let headerRow = 1;

  let map = _prodBuildHeaderMap_(header1);
  if (!map.uid && !map.art && !map.name) {
    map = _prodBuildHeaderMap_(header2);
    headerRow = 2;
  }

  if (!map.art && !map.name) return [];

  const dataStart = headerRow + 1;
  const numRows = lastRow - dataStart + 1;
  if (numRows <= 0) return [];

  const values = sh.getRange(dataStart, 1, numRows, lastCol).getValues();
  const display = sh.getRange(dataStart, 1, numRows, lastCol).getDisplayValues();
  const formulas = sh.getRange(dataStart, 1, numRows, lastCol).getFormulas();

  let richV1 = null;
  let richV2 = null;
  try {
    if (map.view1) richV1 = sh.getRange(dataStart, map.view1, numRows, 1).getRichTextValues();
  } catch (e) {}
  try {
    if (map.view2) richV2 = sh.getRange(dataStart, map.view2, numRows, 1).getRichTextValues();
  } catch (e) {}

  const out = [];

  for (var i = 0; i < numRows; i++) {
    const row = values[i];
    const rowDisp = display[i];
    const rowFormula = formulas[i];

    const uid = map.uid ? String(row[map.uid - 1] || '').trim() : '';
    const art = map.art ? String(row[map.art - 1] || '').trim() : '';
    const name = map.name ? String(row[map.name - 1] || '').trim() : '';
    const unit = map.unit ? String(row[map.unit - 1] || '').trim() : '';

    if (!uid && !art && !name) continue;

    const v1rich = (richV1 && richV1[i]) ? richV1[i][0] : null;
    const v2rich = (richV2 && richV2[i]) ? richV2[i][0] : null;

    const v1 = map.view1
      ? _prodExtractImageUrl_(v1rich, rowFormula[map.view1 - 1], rowDisp[map.view1 - 1], row[map.view1 - 1])
      : '';
    const v2 = map.view2
      ? _prodExtractImageUrl_(v2rich, rowFormula[map.view2 - 1], rowDisp[map.view2 - 1], row[map.view2 - 1])
      : '';

    out.push({
      uid: uid,
      art: art,
      name: name,
      unit: unit,
      view1: v1,
      view2: v2
    });
  }

  return out;
}

function _prodBuildHeaderMap_(headers) {
  const norm = (headers || []).map(_prodNormHeader_);
  return {
    uid: _prodFindColByVariants_(norm, ['uid']),
    art: _prodFindColByVariants_(norm, ['артикул']),
    view1: _prodFindColByVariants_(norm, ['вид 1']),
    view2: _prodFindColByVariants_(norm, ['вид 2']),
    name: _prodFindColByContainsAll_(norm, ['наименование']),
    unit: _prodFindColByVariants_(norm, ['ед. изм.', 'ед. изм', 'ед.', 'ед'])
  };
}

function _prodFindColByVariants_(normHeaders, variants) {
  const vv = (variants || []).map(_prodNormHeader_);
  for (var i = 0; i < normHeaders.length; i++) {
    if (vv.indexOf(normHeaders[i]) >= 0) return i + 1;
  }
  return 0;
}

function _prodFindColByContainsAll_(normHeaders, parts) {
  const pp = (parts || []).map(_prodNormHeader_);
  for (var i = 0; i < normHeaders.length; i++) {
    const h = normHeaders[i] || '';
    if (pp.every(function (p) { return h.indexOf(p) >= 0; })) return i + 1;
  }
  return 0;
}

function _prodExtractImageUrl_(rich, formula, displayValue, rawValue) {
  try {
    if (rich) {
      const direct = rich.getLinkUrl && rich.getLinkUrl();
      if (direct) return direct;

      const runs = rich.getRuns && rich.getRuns();
      if (runs && runs.length) {
        for (var i = 0; i < runs.length; i++) {
          const u = runs[i].getLinkUrl && runs[i].getLinkUrl();
          if (u) return u;
        }
      }
    }
  } catch (e) {}

  const f = String(formula || '');
  let m = f.match(/IMAGE\(\s*"([^"]+)"/i);
  if (m && m[1]) return m[1];

  m = f.match(/HYPERLINK\(\s*"([^"]+)"/i);
  if (m && m[1]) return m[1];

  const candidates = [displayValue, rawValue];
  for (var j = 0; j < candidates.length; j++) {
    const s = String(candidates[j] || '').trim();
    if (/^https?:\/\//i.test(s)) return s;
  }

  return '';
}

/* ========================================================================== */
/* PDF EXPORT                                                                 */
/* ========================================================================== */

function exportProductionRequestPdf_(ss, sheet, fileName) {
  if (!PROD_REQ_CFG.PDF_FOLDER_ID) {
    throw new Error('Не задан ID папки для PDF.');
  }

  const blob = _prodExportSheetToPdfBlob_(ss, sheet, fileName);
  return _prodSavePdfToDriveFolder_(blob, PROD_REQ_CFG.PDF_FOLDER_ID);
}

function _prodExportSheetToPdfBlob_(ss, sheet, fileName) {
  const ssId = ss.getId();
  const gid = sheet.getSheetId();

  const url =
    'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(ssId) + '/export?' +
    [
      'format=pdf',
      'gid=' + gid,
      'size=A4',
      'portrait=' + (PROD_REQ_CFG.PDF_PORTRAIT ? 'true' : 'false'),
      'fitw=true',
      'sheetnames=false',
      'printtitle=false',
      'pagenumbers=false',
      'gridlines=false',
      'fzr=false',
      'top_margin=' + PROD_REQ_CFG.PDF_MARGINS.top,
      'bottom_margin=' + PROD_REQ_CFG.PDF_MARGINS.bottom,
      'left_margin=' + PROD_REQ_CFG.PDF_MARGINS.left,
      'right_margin=' + PROD_REQ_CFG.PDF_MARGINS.right
    ].join('&');

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Не удалось сформировать PDF заявки. Код: ' + code + '. Ответ: ' + resp.getContentText());
  }

  return resp.getBlob().setName(fileName + '.pdf');
}

function _prodSavePdfToDriveFolder_(blob, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  return {
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(file.getId())
  };
}

function showProdRequestLinksDialog_(fileUrl, downloadUrl, meta) {
  const rowNum = meta && meta.prodRowNum ? String(meta.prodRowNum) : '';
  const requestNo = meta && meta.requestNo ? String(meta.requestNo) : '';
  const cancelledCount = meta && typeof meta.cancelledCount !== 'undefined' ? String(meta.cancelledCount) : '0';

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px;">' +
      '<p><b>Заявка на производство № ' + _prodEsc_(requestNo) + ' создана.</b></p>' +
      '<p>Строка в журнале заявок: ' + _prodEsc_(rowNum) + '</p>' +
      '<p>Аннулировано дублей КП: ' + _prodEsc_(cancelledCount) + '</p>' +
      '<p>Скачать (попадёт в “Загрузки” браузера): <a href="' + downloadUrl + '" target="_blank">скачать PDF</a></p>' +
    '</div>'
  ).setWidth(460).setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(html, 'Заявка на производство → PDF');
}

/* ========================================================================== */
/* PROD LOG: CANONICAL HEADERS                                                */
/* ========================================================================== */

function _getProdReqHeadersCanonical_() {
  return [
    'Дата/время создания',
    'Номер заявки ПР',
    'Дата заявки ПР',
    'КП №',
    'Дата КП',
    'Заказчик',
    'Адрес заказчика',
    '№ договора',
    'Счет на оплату',
    'Менеджер',
    'Телефон',
    'Плановая дата отгрузки',
    'Итого к оплате, руб',
    'КП PDF URL (Drive)',
    'КП PDF Download URL',
    'КП Drive File ID',
    'Строка в журнал КП',
    'Позиции (JSON)',
    'Заявка PDF URL (Drive)',
    'Заявка PDF Download URL',
    'Заявка Drive File ID'
  ];
}

function _getProdReqHeaderAliases_() {
  const map = {};

  function addAlias(aliasHeader, canonicalHeader) {
    map[_prodNormHeader_(aliasHeader)] = _prodNormHeader_(canonicalHeader);
  }

  _getProdReqHeadersCanonical_().forEach(function (h) {
    map[_prodNormHeader_(h)] = _prodNormHeader_(h);
  });

  addAlias('Номер заявки на производство', 'Номер заявки ПР');
  addAlias('Дата заявки', 'Дата заявки ПР');

  addAlias('PDF URL (Drive)', 'КП PDF URL (Drive)');
  addAlias('PDF Download URL', 'КП PDF Download URL');
  addAlias('Drive File ID', 'КП Drive File ID');

  addAlias('PDF заявки URL (Drive)', 'Заявка PDF URL (Drive)');
  addAlias('PDF заявки Download URL', 'Заявка PDF Download URL');
  addAlias('Drive File ID заявки', 'Заявка Drive File ID');

  addAlias('Строка в Журнал КП', 'Строка в журнал КП');

  return map;
}

function _canonicalProdHeader_(header) {
  const aliases = _getProdReqHeaderAliases_();
  const norm = _prodNormHeader_(header);
  return aliases[norm] || norm;
}

function ensureProductionRequestLogSchema_(ss) {
  const name = PROD_REQ_CFG.SHEET_PROD_LOG;
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const need = _getProdReqHeadersCanonical_();
  const hasHeaderRow = sh.getLastRow() >= 1;
  const current = hasHeaderRow
    ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0]
    : [];

  const currentNorm = current.map(_canonicalProdHeader_);
  const needNorm = need.map(_canonicalProdHeader_);

  if (!hasHeaderRow || current.join('').trim() === '') {
    sh.getRange(1, 1, 1, need.length).setValues([need]);
    return sh;
  }

  // Добавляем недостающие колонки справа
  need.forEach(function (h, i) {
    if (currentNorm.indexOf(needNorm[i]) < 0) {
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(h);
    }
  });

  // Нормализуем названия уже существующих заголовков
  const finalHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const finalNorm = finalHeaders.map(_canonicalProdHeader_);

  for (var i = 0; i < finalHeaders.length; i++) {
    const hNorm = finalNorm[i];
    const canonicalIdx = needNorm.indexOf(hNorm);
    if (canonicalIdx >= 0) {
      const canonicalTitle = need[canonicalIdx];
      if (String(finalHeaders[i] || '').trim() !== canonicalTitle) {
        sh.getRange(1, i + 1).setValue(canonicalTitle);
      }
    }
  }

  return sh;
}

function buildProdRequestLogRow_(kpRow, requestNo, invoiceNo, plannedShipDate, createdAt, savedPdf) {
  const now = createdAt || new Date();

  return {
    'Дата/время создания': now,
    'Номер заявки ПР': requestNo,
    'Дата заявки ПР': now,
    'КП №': kpRow['КП №'] || '',
    'Дата КП': kpRow['Дата КП'] || '',
    'Заказчик': kpRow['Заказчик'] || '',
    'Адрес заказчика': kpRow['Адрес заказчика'] || '',
    '№ договора': kpRow['№ договора'] || '',
    'Счет на оплату': invoiceNo || '',
    'Менеджер': kpRow['Менеджер'] || '',
    'Телефон': kpRow['Телефон'] || '',
    'Плановая дата отгрузки': plannedShipDate || '',
    'Итого к оплате, руб': kpRow['Итого к оплате, руб'] || '',
    'КП PDF URL (Drive)': _prodPickKpField_(kpRow, ['PDF URL (Drive)', 'КП PDF URL (Drive)']) || '',
    'КП PDF Download URL': _prodPickKpField_(kpRow, ['PDF Download URL', 'КП PDF Download URL']) || '',
    'КП Drive File ID': _prodPickKpField_(kpRow, ['Drive File ID', 'КП Drive File ID']) || '',
    'Строка в журнал КП': kpRow.__row || '',
    'Позиции (JSON)': _prodPickKpField_(kpRow, ['Позиции (JSON)']) || '',
    'Заявка PDF URL (Drive)': savedPdf ? (savedPdf.fileUrl || '') : '',
    'Заявка PDF Download URL': savedPdf ? (savedPdf.downloadUrl || '') : '',
    'Заявка Drive File ID': savedPdf ? (savedPdf.fileId || '') : ''
  };
}

function appendProductionRequestToLog_(sh, rowObj) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const aliases = _getProdReqHeaderAliases_();
  const valueMap = {};

  Object.keys(rowObj || {}).forEach(function (k) {
    valueMap[_canonicalProdHeader_(k)] = rowObj[k];
  });

  const row = headers.map(function (h) {
    const canon = aliases[_prodNormHeader_(h)] || _prodNormHeader_(h);
    return (canon in valueMap) ? valueMap[canon] : '';
  });

  sh.appendRow(row);
  const rowNum = sh.getLastRow();

  const hdrNorm = headers.map(_canonicalProdHeader_);

  const moneyCols = [
    'Итого к оплате, руб'
  ].map(_canonicalProdHeader_);

  for (var c = 1; c <= hdrNorm.length; c++) {
    if (moneyCols.indexOf(hdrNorm[c - 1]) >= 0) {
      try { sh.getRange(rowNum, c).setNumberFormat('#,##0.00'); } catch (e) {}
    }
  }

  const dateCols = [
    'Дата/время создания',
    'Дата заявки ПР'
  ].map(_canonicalProdHeader_);

  for (var d = 1; d <= hdrNorm.length; d++) {
    if (dateCols.indexOf(hdrNorm[d - 1]) >= 0) {
      try { sh.getRange(rowNum, d).setNumberFormat('dd.MM.yyyy HH:mm:ss'); } catch (e) {}
    }
  }

  return rowNum;
}

function findProdRequestDuplicate_(prodLogSheet, requestNo, kpDriveFileId) {
  const lastRow = prodLogSheet.getLastRow();
  if (lastRow < 2) return null;

  const cReq = _prodGetLogColByCanonical_(prodLogSheet, 'Номер заявки ПР');
  const cId = _prodGetLogColByCanonical_(prodLogSheet, 'КП Drive File ID');
  if (!cReq || !cId) return null;

  const n = lastRow - 1;
  const reqVals = prodLogSheet.getRange(2, cReq, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  const idVals = prodLogSheet.getRange(2, cId, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });

  for (var i = n - 1; i >= 0; i--) {
    if (reqVals[i] === requestNo && idVals[i] === kpDriveFileId) {
      return { row: i + 2 };
    }
  }
  return null;
}

function _prodGetLogColByCanonical_(sheet, canonicalHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const canon = headers.map(_canonicalProdHeader_);
  return canon.indexOf(_canonicalProdHeader_(canonicalHeader)) + 1;
}

/* ========================================================================== */
/* KPI LOG                                                                    */
/* ========================================================================== */

function getKpLogRowAsObject_(kpLogSheet, row) {
  const lastCol = kpLogSheet.getLastColumn();
  const headers = kpLogSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const values = kpLogSheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const display = kpLogSheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

  const obj = {};
  for (let i = 0; i < headers.length; i++) {
    const key = String(headers[i] || '').trim();
    if (!key) continue;
    obj[key] = values[i];
    obj[key + '__display'] = display[i];
  }

  obj.__row = row;
  return obj;
}

function buildProdRequestPreviewText_(kpRow, requestNo, invoiceNo, plannedShipDate) {
  const kpDateOnly = _prodFormatDateOnly_(kpRow['Дата КП'] || '');

  return [
    'Будет создана заявка на производство:',
    '',
    'Номер заявки: ' + requestNo,
    'Заказчик: ' + String(kpRow['Заказчик'] || ''),
    'Адрес: ' + String(kpRow['Адрес заказчика'] || ''),
    'Документы основание: Договора №: ' + String(kpRow['№ договора'] || '') +
      '; Коммерческое предложение № ' + String(kpRow['КП №'] || '') + ' от ' + kpDateOnly +
      '; Счет на оплату: ' + invoiceNo,
    'Менеджер / телефон: ' + String(kpRow['Менеджер'] || '') + ' / ' + String(kpRow['Телефон'] || ''),
    'Плановая дата отгрузки: ' + plannedShipDate,
    '',
    'После создания:',
    '• текущая запись получит статус "Отправлена в производство";',
    '• дубли (тот же КП № + Заказчик) будут помечены "Аннулирована".'
  ].join('\n');
}

function setKpLogStatusByRow_(kpLogSheet, row, statusValue) {
  const statusCol = getKpLogStatusColumn_(kpLogSheet);
  if (!statusCol) throw new Error('Не найдена колонка "Статус" в "Журнал КП".');
  kpLogSheet.getRange(row, statusCol).setValue(statusValue);
}

function getKpLogStatusColumn_(kpLogSheet) {
  const headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  const norm = headers.map(function (h) { return String(h || '').trim().toLowerCase(); });
  return norm.indexOf('статус') + 1;
}

function findAnotherSentProductionByKpAndCustomer_(kpLogSheet, kpNo, customer, currentDriveFileId) {
  const lastRow = kpLogSheet.getLastRow();
  if (lastRow < 2) return null;

  const headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  const norm = headers.map(function (h) { return String(h || '').trim().toLowerCase(); });

  const cKp = norm.indexOf('кп №') + 1;
  const cCust = norm.indexOf('заказчик') + 1;
  const cStatus = norm.indexOf('статус') + 1;
  const cDrive = (function () {
    const idx1 = norm.indexOf('drive file id');
    if (idx1 >= 0) return idx1 + 1;
    const idx2 = norm.indexOf('кп drive file id');
    if (idx2 >= 0) return idx2 + 1;
    return 0;
  })();

  if (!cKp || !cCust || !cStatus || !cDrive) return null;

  const n = lastRow - 1;
  const kpVals = kpLogSheet.getRange(2, cKp, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  const custVals = kpLogSheet.getRange(2, cCust, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  const statusVals = kpLogSheet.getRange(2, cStatus, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  const idVals = kpLogSheet.getRange(2, cDrive, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });

  for (let i = 0; i < n; i++) {
    if (kpVals[i] !== kpNo) continue;
    if (custVals[i] !== customer) continue;
    if (idVals[i] === currentDriveFileId) continue;
    if (statusVals[i] === 'Отправлена в производство') {
      return { row: i + 2, driveFileId: idVals[i] };
    }
  }
  return null;
}

function cancelDuplicateKpRows_(kpLogSheet, opts) {
  const kpNo = String(opts.kpNo || '').trim();
  const customer = String(opts.customer || '').trim();
  const exceptDriveFileId = String(opts.exceptDriveFileId || '').trim();
  const keepRow = Number(opts.keepRow || 0);

  const lastRow = kpLogSheet.getLastRow();
  if (lastRow < 2) return 0;

  const headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  const norm = headers.map(function (h) { return String(h || '').trim().toLowerCase(); });

  const cKp = norm.indexOf('кп №') + 1;
  const cCust = norm.indexOf('заказчик') + 1;
  const cStatus = norm.indexOf('статус') + 1;
  const cDrive = (function () {
    const idx1 = norm.indexOf('drive file id');
    if (idx1 >= 0) return idx1 + 1;
    const idx2 = norm.indexOf('кп drive file id');
    if (idx2 >= 0) return idx2 + 1;
    return 0;
  })();

  if (!cKp || !cCust || !cStatus || !cDrive) return 0;

  const n = lastRow - 1;
  const vals = kpLogSheet.getRange(2, 1, n, kpLogSheet.getLastColumn()).getDisplayValues();

  let changed = 0;
  for (let i = 0; i < n; i++) {
    const row = i + 2;
    if (row === keepRow) continue;

    const vKp = String(vals[i][cKp - 1] || '').trim();
    const vCust = String(vals[i][cCust - 1] || '').trim();
    const vDrive = String(vals[i][cDrive - 1] || '').trim();
    const vStatus = String(vals[i][cStatus - 1] || '').trim();

    if (vKp !== kpNo) continue;
    if (vCust !== customer) continue;
    if (!vDrive || vDrive === exceptDriveFileId) continue;
    if (vStatus === 'Отправлена в производство') continue;

    kpLogSheet.getRange(row, cStatus).setValue('Аннулирована');
    changed++;
  }

  return changed;
}

/* ========================================================================== */
/* HELPERS                                                                    */
/* ========================================================================== */

function _prodPickKpField_(obj, keys) {
  const kk = keys || [];
  for (var i = 0; i < kk.length; i++) {
    if (obj && obj.hasOwnProperty(kk[i]) && obj[kk[i]] !== '' && obj[kk[i]] !== null && obj[kk[i]] !== undefined) {
      return obj[kk[i]];
    }
  }
  return '';
}

function _prodPromptRequired_(title, message, defaultValue) {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;

  const value = String(resp.getResponseText() || defaultValue || '').trim();
  if (!value) {
    ui.alert('Поле обязательно для заполнения.');
    return null;
  }
  return value;
}

function _prodNormHeader_(s) {
  return String(s || '')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _prodToNumber_(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const s = String(v || '').replace(/\s+/g, '').replace(',', '.');
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function _prodFormatDateOnly_(v) {
  if (!v && v !== 0) return '';
  try {
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy');
    }
  } catch (e) {}

  const s = String(v || '').trim();
  if (!s) return '';

  const m = s.match(/^(\d{1,2}\.\d{1,2}\.\d{4})/);
  if (m && m[1]) return m[1];

  const d = new Date(s);
  if (!isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }

  return s;
}

function _prodSafeFilePart_(s) {
  return String(s || '')
    .replace(/[\\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
}

function _prodEsc_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}