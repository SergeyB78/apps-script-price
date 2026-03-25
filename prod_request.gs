/**
 * prod_request.gs
 *
 * Версия на основе рабочего сценария с обычными prompt-окнами:
 * - без HTML-формы ввода
 * - дата плановой отгрузки по умолчанию = сегодня + 45 дней
 * - PDF заявки создаётся через временный лист, который потом удаляется
 * - Журнал заявок на производство заполняется ПО ЗАГОЛОВКАМ
 * - после формирования показывается диалог со ссылкой "скачать PDF"
 * - PDF: левое поле 2 см, правое поле 1 см
 * - таблица растянута по ширине
 * - картинки вставляются только из Google Drive (по ссылке/ID)
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
  COL_WIDTHS: [120, 120, 120, 440, 55, 65, 210], // A:G
  DEFAULT_ROW_HEIGHT: 22,
  TITLE_FONT_SIZE: 16,
  BODY_FONT_SIZE: 10,
  HEADER_FONT_SIZE: 11,

  DATA_ROW_HEIGHT: 92,
  IMG_SIZE: 78,
  IMG_X_OFFSET: 8,
  IMG_Y_OFFSET: 6,
  MAX_IMAGES_TOTAL: 60,

  PDF_PORTRAIT: false, // landscape
  PDF_MARGINS_CM: {
    top: 0.5,
    bottom: 0.5,
    left: 2.0,
    right: 1.0
  }
};

/* ========================================================================== */
/* ENTRY POINTS                                                               */
/* ========================================================================== */

function createProductionOrderFromSelectedKp() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  var kpLog = ensureKpLogSchema_(ss);
  var prodLog = ensureProductionRequestLogSchema_(ss);

  var activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getName() !== kpLog.getName()) {
    ui.alert(
      'Для создания заявки на производство перейдите на лист "' + kpLog.getName() + '" и выделите нужную строку.'
    );
    return;
  }

  var row = activeSheet.getActiveRange() ? activeSheet.getActiveRange().getRow() : 0;
  if (row < 2) {
    ui.alert('Выберите строку в "Журнал КП" (не шапку).');
    return;
  }

  var kpRow = getKpLogRowAsObject_(kpLog, row);

  var kpDriveFileId = String(_prodPickKpField_(kpRow, ['Drive File ID', 'КП Drive File ID']) || '').trim();
  var kpNo = String(kpRow['КП №'] || '').trim();
  var customer = String(kpRow['Заказчик'] || '').trim();
  var status = String(kpRow['Статус'] || '').trim() || 'Новая';

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

  var existingSent = findAnotherSentProductionByKpAndCustomer_(kpLog, kpNo, customer, kpDriveFileId);
  if (existingSent) {
    ui.alert(
      'Заявка на производство уже создана по дублю этого КП.\n\n' +
      'Строка журнала КП: ' + existingSent.row + '\n' +
      'Статус: Отправлена в производство\n\n' +
      'Для текущей строки создание заявки запрещено.'
    );
    return;
  }

  var previewText = buildProdRequestPreviewText_(kpRow, '', '', '');
  var confirm = ui.alert(
    'Создать заявку на производство',
    previewText + '\n\nПродолжить?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var requestNoPrompt = ui.prompt(
    'Номер заявки на производство',
    'Введите номер заявки на производство:',
    ui.ButtonSet.OK_CANCEL
  );
  if (requestNoPrompt.getSelectedButton() !== ui.Button.OK) return;
  var requestNo = String(requestNoPrompt.getResponseText() || '').trim();
  if (!requestNo) {
    ui.alert('Создание отменено: номер заявки не введён.');
    return;
  }

  var invoicePrompt = ui.prompt(
    'Счет на оплату',
    'Введите номер счета на оплату:',
    ui.ButtonSet.OK_CANCEL
  );
  if (invoicePrompt.getSelectedButton() !== ui.Button.OK) return;
  var invoiceNo = String(invoicePrompt.getResponseText() || '').trim();
  if (!invoiceNo) {
    ui.alert('Создание отменено: номер счета не введён.');
    return;
  }

  var defaultPlanDate = _prodFormatDateShort_(_prodAddDays_(new Date(), 45));
  var planPrompt = ui.prompt(
    'Плановая дата отгрузки',
    'Введите плановую дату отгрузки.\n\nПо умолчанию: ' + defaultPlanDate + '\n(можно оставить поле пустым и нажать OK)',
    ui.ButtonSet.OK_CANCEL
  );
  if (planPrompt.getSelectedButton() !== ui.Button.OK) return;
  var plannedShipDate = String(planPrompt.getResponseText() || '').trim();
  if (!plannedShipDate) plannedShipDate = defaultPlanDate;

  var dupProd = findProdRequestDuplicate_(prodLog, requestNo, kpDriveFileId);
  if (dupProd) {
    ui.alert(
      'Такая заявка уже есть в "Журнал заявок на производство".\n\n' +
      'Строка: ' + dupProd.row + '\n' +
      'Номер заявки: ' + requestNo
    );
    return;
  }

  var now = new Date();
  var reqData = buildProductionRequestData_(ss, kpRow, requestNo, invoiceNo, plannedShipDate, now);

  var tempSheet = null;
  var saved = null;

  try {
    tempSheet = buildProductionRequestTempSheet_(ss, reqData);

    SpreadsheetApp.flush();
    Utilities.sleep(400);

    saved = exportProductionRequestPdf_(ss, tempSheet, reqData.fileName);

    var prodRow = buildProdRequestLogRow_(kpRow, requestNo, invoiceNo, plannedShipDate, now, saved);
    var prodRowNum = appendProductionRequestToLog_(prodLog, prodRow);

    setKpLogStatusByRow_(kpLog, row, 'Отправлена в производство');

    var cancelledCount = cancelDuplicateKpRows_(kpLog, {
      kpNo: kpNo,
      customer: customer,
      exceptDriveFileId: kpDriveFileId,
      keepRow: row
    });

    try { ss.setActiveSheet(kpLog); } catch (e) {}

    showProdRequestLinksDialog_(saved.fileUrl, saved.downloadUrl, {
      prodRowNum: prodRowNum,
      requestNo: requestNo,
      cancelledCount: cancelledCount
    });

  } catch (err) {
    ui.alert('Ошибка при формировании заявки:\n\n' + ((err && err.message) ? err.message : String(err)));
  } finally {
    try {
      if (tempSheet && ss.getSheetByName(tempSheet.getName())) {
        ss.deleteSheet(tempSheet);
      }
    } catch (e) {}
  }
}

function createProductionRequestFromSelectedKp() {
  return createProductionOrderFromSelectedKp();
}

/* ========================================================================== */
/* DATA BUILD                                                                 */
/* ========================================================================== */

function buildProductionRequestData_(ss, kpRow, requestNo, invoiceNo, plannedShipDate, createdAt) {
  var requestDateTime = createdAt || new Date();
  var requestDateOnly = _prodFormatDateOnly_(requestDateTime);

  var kpNo = String(kpRow['КП №'] || '').trim();
  var kpDateRaw = kpRow['Дата КП'] || '';
  var kpDateOnly = _prodFormatDateOnly_(kpDateRaw);

  var customer = String(kpRow['Заказчик'] || '').trim();
  var customerAddr = String(kpRow['Адрес заказчика'] || '').trim();
  var contractNo = String(kpRow['№ договора'] || '').trim();
  var manager = String(kpRow['Менеджер'] || '').trim();
  var phone = String(kpRow['Телефон'] || '').trim();

  var docsBase =
    'Договора №: ' + contractNo +
    '; Коммерческое предложение № ' + kpNo + ' от ' + kpDateOnly +
    '; Счет на оплату: ' + invoiceNo;

  var items = _prodParseItemsJson_(_prodPickKpField_(kpRow, ['Позиции (JSON)']) || '[]');
  var catalogIndex = _prodBuildCatalogIndex_(ss);
  var enrichedItems = _prodEnrichItemsFromCatalog_(items, catalogIndex);

  var fileName =
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
    var arr = JSON.parse(jsonString || '[]');
    if (Array.isArray(arr)) return arr;
  } catch (e) {}
  return [];
}

function _prodEnrichItemsFromCatalog_(items, catalogIndex) {
  return (items || []).map(function (item) {
    var uid = String(item.uid || '').trim();
    var art = String(item.art || '').trim();

    var src = null;
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
  var sheetName = PROD_REQ_CFG.TEMP_SHEET_PREFIX + new Date().getTime();
  var sh = ss.insertSheet(sheetName);

  _prodResetSheet_(sh);
  _prodEnsureSheetSize_(sh, Math.max(60, data.items.length * 3 + 20), PROD_REQ_CFG.FORM_COLS);
  _prodSetupSheetLayout_(sh);

  sh.getRange(1, 1, 1, 7).merge()
    .setValue(data.title)
    .setFontWeight('bold')
    .setFontSize(PROD_REQ_CFG.TITLE_FONT_SIZE)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 28);

  var r = 3;
  _prodSetMetaRow_(sh, r++, 'Заказчик:', data.customer, 22);
  _prodSetMetaRow_(sh, r++, 'Адрес заказчика:', data.customerAddr, 22);
  _prodSetMetaRow_(sh, r++, 'Документы основание:', data.docsBase, 38);
  _prodSetMetaRow_(sh, r++, 'Менеджер / телефон:', data.managerPhone, 22);
  _prodSetMetaRow_(sh, r++, 'Плановая дата отгрузки:', data.plannedShipDate, 22);

  r += 1;

  var headerRow = r;
  var headers = ['Артикул', 'Вид 1', 'Вид 2', 'Наименование / размеры', 'Ед.', 'Кол-во', 'Примечание'];

  sh.getRange(headerRow, 1, 1, 7).setValues([headers]);
  sh.getRange(headerRow, 1, 1, 7)
    .setFontWeight('bold')
    .setFontSize(PROD_REQ_CFG.HEADER_FONT_SIZE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true);
  sh.setRowHeight(headerRow, 30);

  var imagesInserted = 0;
  var row = headerRow + 1;

  data.items.forEach(function (item) {
    sh.setRowHeight(row, PROD_REQ_CFG.DATA_ROW_HEIGHT);

    sh.getRange(row, 1).setValue(item.art || '');
    sh.getRange(row, 4).setValue(item.name || '');
    sh.getRange(row, 5).setValue(item.unit || '');
    sh.getRange(row, 6).setValue(item.qty || '');
    sh.getRange(row, 7).setValue(item.note || '');

    sh.getRange(row, 1, 1, 7)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true);

    sh.getRange(row, 1).setHorizontalAlignment('left');
    sh.getRange(row, 4).setHorizontalAlignment('left');
    sh.getRange(row, 5).setHorizontalAlignment('center');
    sh.getRange(row, 6).setHorizontalAlignment('center');
    sh.getRange(row, 7).setHorizontalAlignment('left');

    if (imagesInserted < PROD_REQ_CFG.MAX_IMAGES_TOTAL) {
      if (_prodInsertDriveImage_(sh, row, 2, item.view1)) imagesInserted++;
    }
    if (imagesInserted < PROD_REQ_CFG.MAX_IMAGES_TOTAL) {
      if (_prodInsertDriveImage_(sh, row, 3, item.view2)) imagesInserted++;
    }

    row++;
  });

  var lastDataRow = Math.max(headerRow, row - 1);
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

function _prodInsertDriveImage_(sh, row, col, url) {
  var fileId = _prodExtractDriveFileId_(url);
  if (!fileId) return false;

  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    var img = sh.insertImage(blob, col, row, PROD_REQ_CFG.IMG_X_OFFSET, PROD_REQ_CFG.IMG_Y_OFFSET);
    try { img.setWidth(PROD_REQ_CFG.IMG_SIZE); } catch (e) {}
    try { img.setHeight(PROD_REQ_CFG.IMG_SIZE); } catch (e) {}
    return true;
  } catch (e) {
    return false;
  }
}

function _prodResetSheet_(sh) {
  var maxR = Math.max(sh.getMaxRows(), 1);
  var maxC = Math.max(sh.getMaxColumns(), 1);

  try {
    sh.getImages().forEach(function (img) { img.remove(); });
  } catch (e) {}

  try {
    var f = sh.getFilter();
    if (f) f.remove();
  } catch (e) {}

  var rng = sh.getRange(1, 1, maxR, maxC);
  try { rng.breakApart(); } catch (e) {}

  sh.clear();
  try {
    sh.showRows(1, maxR);
    sh.showColumns(1, maxC);
  } catch (e) {}

  sh.setRowHeights(1, maxR, PROD_REQ_CFG.DEFAULT_ROW_HEIGHT);
}

function _prodEnsureSheetSize_(sh, minRows, minCols) {
  var curRows = sh.getMaxRows();
  var curCols = sh.getMaxColumns();
  if (curRows < minRows) sh.insertRowsAfter(curRows, minRows - curRows);
  if (curCols < minCols) sh.insertColumnsAfter(curCols, minCols - curCols);
}

function _prodSetupSheetLayout_(sh) {
  for (var c = 1; c <= PROD_REQ_CFG.COL_WIDTHS.length; c++) {
    sh.setColumnWidth(c, PROD_REQ_CFG.COL_WIDTHS[c - 1]);
  }
}

/* ========================================================================== */
/* CATALOG / IMAGE URLS                                                       */
/* ========================================================================== */

function _prodBuildCatalogIndex_(ss) {
  var out = { byUid: {}, byArt: {} };

  [ss.getSheetByName(PROD_REQ_CFG.SHEET_DB), ss.getSheetByName(PROD_REQ_CFG.SHEET_PRICE)]
    .filter(Boolean)
    .forEach(function (sh) {
      _prodReadCatalogRows_(sh).forEach(function (r) {
        if (r.uid && !out.byUid[r.uid]) out.byUid[r.uid] = r;
        if (r.art && !out.byArt[r.art]) out.byArt[r.art] = r;
      });
    });

  return out;
}

function _prodReadCatalogRows_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  var header1 = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var header2 = lastRow >= 2 ? sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] : [];
  var headerRow = 1;

  var map = _prodBuildHeaderMap_(header1);
  if (!map.uid && !map.art && !map.name) {
    map = _prodBuildHeaderMap_(header2);
    headerRow = 2;
  }

  if (!map.art && !map.name) return [];

  var dataStart = headerRow + 1;
  var numRows = lastRow - dataStart + 1;
  if (numRows <= 0) return [];

  var values = sh.getRange(dataStart, 1, numRows, lastCol).getValues();
  var display = sh.getRange(dataStart, 1, numRows, lastCol).getDisplayValues();
  var formulas = sh.getRange(dataStart, 1, numRows, lastCol).getFormulas();

  var richV1 = null;
  var richV2 = null;
  try {
    if (map.view1) richV1 = sh.getRange(dataStart, map.view1, numRows, 1).getRichTextValues();
  } catch (e) {}
  try {
    if (map.view2) richV2 = sh.getRange(dataStart, map.view2, numRows, 1).getRichTextValues();
  } catch (e) {}

  var out = [];

  for (var i = 0; i < numRows; i++) {
    var row = values[i];
    var rowDisp = display[i];
    var rowFormula = formulas[i];

    var uid = map.uid ? String(row[map.uid - 1] || '').trim() : '';
    var art = map.art ? String(row[map.art - 1] || '').trim() : '';
    var name = map.name ? String(row[map.name - 1] || '').trim() : '';
    var unit = map.unit ? String(row[map.unit - 1] || '').trim() : '';

    if (!uid && !art && !name) continue;

    var v1rich = (richV1 && richV1[i]) ? richV1[i][0] : null;
    var v2rich = (richV2 && richV2[i]) ? richV2[i][0] : null;

    var v1 = map.view1 ? _prodExtractImageUrl_(v1rich, rowFormula[map.view1 - 1], rowDisp[map.view1 - 1], row[map.view1 - 1]) : '';
    var v2 = map.view2 ? _prodExtractImageUrl_(v2rich, rowFormula[map.view2 - 1], rowDisp[map.view2 - 1], row[map.view2 - 1]) : '';

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
  var norm = (headers || []).map(_prodNormHeader_);
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
  var vv = (variants || []).map(_prodNormHeader_);
  for (var i = 0; i < normHeaders.length; i++) {
    if (vv.indexOf(normHeaders[i]) >= 0) return i + 1;
  }
  return 0;
}

function _prodFindColByContainsAll_(normHeaders, parts) {
  var pp = (parts || []).map(_prodNormHeader_);
  for (var i = 0; i < normHeaders.length; i++) {
    var h = normHeaders[i] || '';
    if (pp.every(function (p) { return h.indexOf(p) >= 0; })) return i + 1;
  }
  return 0;
}

function _prodExtractImageUrl_(rich, formula, displayValue, rawValue) {
  try {
    if (rich) {
      var direct = rich.getLinkUrl && rich.getLinkUrl();
      if (direct) return direct;

      var runs = rich.getRuns && rich.getRuns();
      if (runs && runs.length) {
        for (var i = 0; i < runs.length; i++) {
          var u = runs[i].getLinkUrl && runs[i].getLinkUrl();
          if (u) return u;
        }
      }
    }
  } catch (e) {}

  var f = String(formula || '');
  var m = f.match(/IMAGE\(\s*"([^"]+)"/i);
  if (m && m[1]) return m[1];

  m = f.match(/HYPERLINK\(\s*"([^"]+)"/i);
  if (m && m[1]) return m[1];

  var candidates = [displayValue, rawValue];
  for (var j = 0; j < candidates.length; j++) {
    var s = String(candidates[j] || '').trim();
    if (/^https?:\/\//i.test(s)) return s;
  }

  return '';
}

function _prodExtractDriveFileId_(url) {
  var s = String(url || '').trim();
  if (!s) return '';

  var m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/i);
  if (m && m[1]) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/i);
  if (m && m[1]) return m[1];

  m = s.match(/[-\w]{25,}/);
  if (m && m[0] && s.indexOf('drive.google.com') >= 0) return m[0];

  return '';
}

/* ========================================================================== */
/* PDF EXPORT                                                                 */
/* ========================================================================== */

function exportProductionRequestPdf_(ss, sheet, fileName) {
  if (!PROD_REQ_CFG.PDF_FOLDER_ID) {
    throw new Error('Не задан ID папки для PDF.');
  }

  var blob = _prodExportSheetToPdfBlob_(ss, sheet, fileName);
  return _prodSavePdfToDriveFolder_(blob, PROD_REQ_CFG.PDF_FOLDER_ID);
}

function _prodExportSheetToPdfBlob_(ss, sheet, fileName) {
  var ssId = ss.getId();
  var gid = sheet.getSheetId();
  var m = PROD_REQ_CFG.PDF_MARGINS_CM;

  var url =
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
      'top_margin=' + _prodCmToIn_(m.top),
      'bottom_margin=' + _prodCmToIn_(m.bottom),
      'left_margin=' + _prodCmToIn_(m.left),
      'right_margin=' + _prodCmToIn_(m.right)
    ].join('&');

  var token = ScriptApp.getOAuthToken();
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Не удалось сформировать PDF заявки. Код: ' + code + '. Ответ: ' + resp.getContentText());
  }

  return resp.getBlob().setName(fileName + '.pdf');
}

function _prodSavePdfToDriveFolder_(blob, folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(blob);

  return {
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(file.getId())
  };
}

function showProdRequestLinksDialog_(fileUrl, downloadUrl, meta) {
  var rowNum = meta && meta.prodRowNum ? String(meta.prodRowNum) : '';
  var requestNo = meta && meta.requestNo ? String(meta.requestNo) : '';
  var cancelledCount = meta && typeof meta.cancelledCount !== 'undefined' ? String(meta.cancelledCount) : '0';

  var html = HtmlService.createHtmlOutput(
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
/* PROD LOG                                                                   */
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
  var map = {};

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
  var aliases = _getProdReqHeaderAliases_();
  var norm = _prodNormHeader_(header);
  return aliases[norm] || norm;
}

function ensureProductionRequestLogSchema_(ss) {
  var name = PROD_REQ_CFG.SHEET_PROD_LOG;
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  var need = _getProdReqHeadersCanonical_();
  var hasHeaderRow = sh.getLastRow() >= 1;
  var current = hasHeaderRow
    ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0]
    : [];

  var currentNorm = current.map(_canonicalProdHeader_);
  var needNorm = need.map(_canonicalProdHeader_);

  if (!hasHeaderRow || current.join('').trim() === '') {
    sh.getRange(1, 1, 1, need.length).setValues([need]);
    return sh;
  }

  need.forEach(function (h, i) {
    if (currentNorm.indexOf(needNorm[i]) < 0) {
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(h);
    }
  });

  return sh;
}

function buildProdRequestLogRow_(kpRow, requestNo, invoiceNo, plannedShipDate, createdAt, savedPdf) {
  var now = createdAt || new Date();

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
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  var aliases = _getProdReqHeaderAliases_();
  var valueMap = {};

  Object.keys(rowObj || {}).forEach(function (k) {
    valueMap[_canonicalProdHeader_(k)] = rowObj[k];
  });

  var row = headers.map(function (h) {
    var canon = aliases[_prodNormHeader_(h)] || _prodNormHeader_(h);
    return (canon in valueMap) ? valueMap[canon] : '';
  });

  sh.appendRow(row);
  var rowNum = sh.getLastRow();

  var hdrNorm = headers.map(_canonicalProdHeader_);
  var moneyCols = ['Итого к оплате, руб'].map(_canonicalProdHeader_);
  for (var c = 1; c <= hdrNorm.length; c++) {
    if (moneyCols.indexOf(hdrNorm[c - 1]) >= 0) {
      try { sh.getRange(rowNum, c).setNumberFormat('#,##0.00'); } catch (e) {}
    }
  }

  var dateCols = ['Дата/время создания', 'Дата заявки ПР'].map(_canonicalProdHeader_);
  for (var d = 1; d <= hdrNorm.length; d++) {
    if (dateCols.indexOf(hdrNorm[d - 1]) >= 0) {
      try { sh.getRange(rowNum, d).setNumberFormat('dd.MM.yyyy HH:mm:ss'); } catch (e) {}
    }
  }

  return rowNum;
}

function findProdRequestDuplicate_(prodLogSheet, requestNo, kpDriveFileId) {
  var lastRow = prodLogSheet.getLastRow();
  if (lastRow < 2) return null;

  var cReq = _prodGetLogColByCanonical_(prodLogSheet, 'Номер заявки ПР');
  var cId = _prodGetLogColByCanonical_(prodLogSheet, 'КП Drive File ID');
  if (!cReq || !cId) return null;

  var n = lastRow - 1;
  var reqVals = prodLogSheet.getRange(2, cReq, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  var idVals = prodLogSheet.getRange(2, cId, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });

  for (var i = n - 1; i >= 0; i--) {
    if (reqVals[i] === requestNo && idVals[i] === kpDriveFileId) {
      return { row: i + 2 };
    }
  }
  return null;
}

function _prodGetLogColByCanonical_(sheet, canonicalHeader) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var canon = headers.map(_canonicalProdHeader_);
  return canon.indexOf(_canonicalProdHeader_(canonicalHeader)) + 1;
}

/* ========================================================================== */
/* KPI LOG                                                                    */
/* ========================================================================== */

function getKpLogRowAsObject_(kpLogSheet, row) {
  var lastCol = kpLogSheet.getLastColumn();
  var headers = kpLogSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var values = kpLogSheet.getRange(row, 1, 1, lastCol).getValues()[0];
  var display = kpLogSheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').trim();
    if (!key) continue;
    obj[key] = values[i];
    obj[key + '__display'] = display[i];
  }

  obj.__row = row;
  return obj;
}

function findAnotherSentProductionByKpAndCustomer_(kpLogSheet, kpNo, customer, currentDriveFileId) {
  var lastRow = kpLogSheet.getLastRow();
  if (lastRow < 2) return null;

  var headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  var norm = headers.map(function (h) { return String(h || '').trim().toLowerCase(); });

  var cKp = norm.indexOf('кп №') + 1;
  var cCust = norm.indexOf('заказчик') + 1;
  var cStatus = norm.indexOf('статус') + 1;
  var cDrive = (function () {
    var idx1 = norm.indexOf('drive file id');
    if (idx1 >= 0) return idx1 + 1;
    var idx2 = norm.indexOf('кп drive file id');
    if (idx2 >= 0) return idx2 + 1;
    return 0;
  })();

  if (!cKp || !cCust || !cStatus || !cDrive) return null;

  var n = lastRow - 1;
  var kpVals = kpLogSheet.getRange(2, cKp, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  var custVals = kpLogSheet.getRange(2, cCust, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  var statusVals = kpLogSheet.getRange(2, cStatus, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
  var idVals = kpLogSheet.getRange(2, cDrive, n, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });

  for (var i = 0; i < n; i++) {
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
  var kpNo = String(opts.kpNo || '').trim();
  var customer = String(opts.customer || '').trim();
  var exceptDriveFileId = String(opts.exceptDriveFileId || '').trim();
  var keepRow = Number(opts.keepRow || 0);

  var lastRow = kpLogSheet.getLastRow();
  if (lastRow < 2) return 0;

  var headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  var norm = headers.map(function (h) { return String(h || '').trim().toLowerCase(); });

  var cKp = norm.indexOf('кп №') + 1;
  var cCust = norm.indexOf('заказчик') + 1;
  var cStatus = norm.indexOf('статус') + 1;
  var cDrive = (function () {
    var idx1 = norm.indexOf('drive file id');
    if (idx1 >= 0) return idx1 + 1;
    var idx2 = norm.indexOf('кп drive file id');
    if (idx2 >= 0) return idx2 + 1;
    return 0;
  })();

  if (!cKp || !cCust || !cStatus || !cDrive) return 0;

  var n = lastRow - 1;
  var vals = kpLogSheet.getRange(2, 1, n, kpLogSheet.getLastColumn()).getDisplayValues();

  var changed = 0;
  for (var i = 0; i < n; i++) {
    var row = i + 2;
    if (row === keepRow) continue;

    var vKp = String(vals[i][cKp - 1] || '').trim();
    var vCust = String(vals[i][cCust - 1] || '').trim();
    var vDrive = String(vals[i][cDrive - 1] || '').trim();
    var vStatus = String(vals[i][cStatus - 1] || '').trim();

    if (vKp !== kpNo) continue;
    if (vCust !== customer) continue;
    if (!vDrive || vDrive === exceptDriveFileId) continue;
    if (vStatus === 'Отправлена в производство') continue;

    kpLogSheet.getRange(row, cStatus).setValue('Аннулирована');
    changed++;
  }

  return changed;
}

function setKpLogStatusByRow_(kpLogSheet, row, statusValue) {
  var statusCol = getKpLogStatusColumn_(kpLogSheet);
  if (!statusCol) throw new Error('Не найдена колонка "Статус" в "Журнал КП".');
  kpLogSheet.getRange(row, statusCol).setValue(statusValue);
}

function getKpLogStatusColumn_(kpLogSheet) {
  var headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  var norm = headers.map(function (h) { return String(h || '').trim().toLowerCase(); });
  return norm.indexOf('статус') + 1;
}

/* ========================================================================== */
/* PREVIEW / DIALOG                                                           */
/* ========================================================================== */

function buildProdRequestPreviewText_(kpRow, requestNo, invoiceNo, plannedShipDate) {
  var kpDateOnly = _prodFormatDateOnly_(kpRow['Дата КП'] || '');

  return [
    'Будет создана заявка на производство по записи "Журнал КП":',
    '',
    'Строка журнала КП: ' + String(kpRow.__row || ''),
    'КП №: ' + String(kpRow['КП №'] || ''),
    'Дата КП: ' + kpDateOnly,
    'Заказчик: ' + String(kpRow['Заказчик'] || ''),
    'Адрес: ' + String(kpRow['Адрес заказчика'] || ''),
    '',
    'Документы основание:',
    'Договора №: ' + String(kpRow['№ договора'] || '') +
      '; Коммерческое предложение № ' + String(kpRow['КП №'] || '') +
      ' от ' + kpDateOnly +
      (invoiceNo ? '; Счет на оплату: ' + invoiceNo : ''),
    '',
    'Менеджер / телефон: ' + String(kpRow['Менеджер'] || '') + ' / ' + String(kpRow['Телефон'] || ''),
    (plannedShipDate ? 'Плановая дата отгрузки: ' + plannedShipDate : '')
  ].join('\n');
}

/* ========================================================================== */
/* HELPERS                                                                    */
/* ========================================================================== */

function _prodPickKpField_(obj, keys) {
  var kk = keys || [];
  for (var i = 0; i < kk.length; i++) {
    if (obj && obj.hasOwnProperty(kk[i]) && obj[kk[i]] !== '' && obj[kk[i]] !== null && obj[kk[i]] !== undefined) {
      return obj[kk[i]];
    }
  }
  return '';
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
  var s = String(v || '').replace(/\s+/g, '').replace(',', '.');
  var n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function _prodFormatDateOnly_(v) {
  if (!v && v !== 0) return '';
  try {
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy');
    }
  } catch (e) {}

  var s = String(v || '').trim();
  if (!s) return '';

  var m = s.match(/^(\d{1,2}\.\d{1,2}\.\d{4})/);
  if (m && m[1]) return m[1];

  var d = new Date(s);
  if (!isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }

  return s;
}

function _prodFormatDateShort_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yy');
}

function _prodAddDays_(dateObj, days) {
  var d = new Date(dateObj.getTime());
  d.setDate(d.getDate() + Number(days || 0));
  return d;
}

function _prodCmToIn_(cm) {
  return Number(cm || 0) / 2.54;
}

function _prodSafeFilePart_(s) {
  return String(s || '')
    .replace(/[\\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
}

function _prodEscAttr_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function _prodEsc_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}