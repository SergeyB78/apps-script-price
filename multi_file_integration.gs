/**
 * multi_file_integration.gs
 *
 * Первый этап архитектуры: "один мастер-файл + отдельный рабочий файл менеджера".
 *
 * Что уже реализовано:
 * 1) Не трогаем существующий onOpen() из main.gs.
 * 2) Добавляем отдельное меню "Интеграция" через installable open-trigger.
 * 3) В рабочем файле менеджера умеем подтягивать из мастер-файла справочники:
 *    - БД Оборудования
 *    - Links_UID
 *    - Справочник
 *    После синхронизации можем автоматически пересобрать Прайс.
 * 4) В мастер-файле умеем собирать сводные журналы из файлов менеджеров:
 *    - Сводный Журнал КП
 *    - Сводный Журнал заявок на производство
 *
 * Базовый сценарий внедрения:
 *
 * В каждом рабочем файле менеджера:
 * - выполнить installMultiFileOpenTrigger_()
 * - открыть "Интеграция" -> "Открыть настройки интеграции"
 * - указать APP_ROLE = manager и MASTER_FILE_ID
 * - выполнить "Обновить базы из мастер-файла"
 *
 * В мастер-файле:
 * - выполнить installMultiFileOpenTrigger_()
 * - открыть "Интеграция" -> "Открыть настройки интеграции"
 * - указать APP_ROLE = master
 * - открыть "Реестр файлов менеджеров" и заполнить Spreadsheet ID файлов менеджеров
 * - выполнить "Собрать оба сводных журнала"
 */

var MF_CFG_DEFAULTS = {
  SETTINGS_SHEET: 'Интеграция',
  MANAGER_REGISTRY_SHEET: 'Реестр файлов менеджеров',

  APP_ROLE: 'manager', // manager | master
  MASTER_FILE_ID: '',

  MASTER_DB_SHEET: 'БД Оборудования',
  MASTER_LINKS_SHEET: 'Links_UID',
  MASTER_MANAGERS_SHEET: 'Справочник',

  LOCAL_DB_SHEET: 'БД Оборудования',
  LOCAL_LINKS_SHEET: 'Links_UID',
  LOCAL_MANAGERS_SHEET: 'Справочник',

  LOCAL_KP_LOG_SHEET: 'Журнал КП',
  LOCAL_PROD_LOG_SHEET: 'Журнал заявок на производство',

  MASTER_SUMMARY_KP_SHEET: 'Сводный Журнал КП',
  MASTER_SUMMARY_PROD_SHEET: 'Сводный Журнал заявок на производство',

  AUTO_REBUILD_PRICE_AFTER_SYNC: 'yes'
};

function installMultiFileOpenTrigger_() {
  var ss = SpreadsheetApp.getActive();
  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function (t) {
    try {
      if (t.getHandlerFunction && t.getHandlerFunction() === 'onOpenMultiFile_') {
        ScriptApp.deleteTrigger(t);
      }
    } catch (e) {}
  });

  ScriptApp.newTrigger('onOpenMultiFile_')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  SpreadsheetApp.getUi().alert(
    'Готово',
    'Установлен отдельный open-trigger для меню "Интеграция".\n\n' +
    'Чтобы увидеть меню сразу, выполните функцию onOpenMultiFile_() один раз вручную или просто закройте и заново откройте файл.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpenMultiFile_() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Интеграция');

  menu.addItem('Открыть настройки интеграции', 'openIntegrationSettings_')
    .addItem('Открыть реестр файлов менеджеров', 'openManagerFilesRegistry_')
    .addSeparator()
    .addItem('Обновить базы из мастер-файла', 'syncReferenceSheetsFromMaster_')
    .addSeparator()
    .addItem('Собрать сводный Журнал КП', 'rebuildMasterKpSummary_')
    .addItem('Собрать сводный Журнал заявок', 'rebuildMasterProdSummary_')
    .addItem('Собрать оба сводных журнала', 'rebuildAllMasterSummaries_')
    .addToUi();
}

/* ========================= Settings / Registry ========================= */

function openIntegrationSettings_() {
  var sh = ensureIntegrationSettingsSheet_();
  SpreadsheetApp.getActive().setActiveSheet(sh);
}

function ensureIntegrationSettingsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(MF_CFG_DEFAULTS.SETTINGS_SHEET);
  if (!sh) sh = ss.insertSheet(MF_CFG_DEFAULTS.SETTINGS_SHEET);

  if (sh.getLastRow() < 2) {
    var rows = [
      ['Параметр', 'Значение', 'Комментарий'],
      ['APP_ROLE', MF_CFG_DEFAULTS.APP_ROLE, 'manager или master'],
      ['MASTER_FILE_ID', MF_CFG_DEFAULTS.MASTER_FILE_ID, 'Spreadsheet ID мастер-файла'],
      ['MASTER_DB_SHEET', MF_CFG_DEFAULTS.MASTER_DB_SHEET, 'Лист БД в мастер-файле'],
      ['MASTER_LINKS_SHEET', MF_CFG_DEFAULTS.MASTER_LINKS_SHEET, 'Лист Links_UID в мастер-файле'],
      ['MASTER_MANAGERS_SHEET', MF_CFG_DEFAULTS.MASTER_MANAGERS_SHEET, 'Лист Справочник в мастер-файле'],
      ['LOCAL_DB_SHEET', MF_CFG_DEFAULTS.LOCAL_DB_SHEET, 'Локальный лист БД'],
      ['LOCAL_LINKS_SHEET', MF_CFG_DEFAULTS.LOCAL_LINKS_SHEET, 'Локальный лист Links_UID'],
      ['LOCAL_MANAGERS_SHEET', MF_CFG_DEFAULTS.LOCAL_MANAGERS_SHEET, 'Локальный лист Справочник'],
      ['LOCAL_KP_LOG_SHEET', MF_CFG_DEFAULTS.LOCAL_KP_LOG_SHEET, 'Локальный Журнал КП'],
      ['LOCAL_PROD_LOG_SHEET', MF_CFG_DEFAULTS.LOCAL_PROD_LOG_SHEET, 'Локальный Журнал заявок на производство'],
      ['MASTER_SUMMARY_KP_SHEET', MF_CFG_DEFAULTS.MASTER_SUMMARY_KP_SHEET, 'Сводный Журнал КП в мастер-файле'],
      ['MASTER_SUMMARY_PROD_SHEET', MF_CFG_DEFAULTS.MASTER_SUMMARY_PROD_SHEET, 'Сводный Журнал заявок в мастер-файле'],
      ['MANAGER_REGISTRY_SHEET', MF_CFG_DEFAULTS.MANAGER_REGISTRY_SHEET, 'Реестр файлов менеджеров (только мастер)'],
      ['AUTO_REBUILD_PRICE_AFTER_SYNC', MF_CFG_DEFAULTS.AUTO_REBUILD_PRICE_AFTER_SYNC, 'yes / no']
    ];
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.setFrozenRows(1);
    try {
      sh.getRange('A1:C1').setFontWeight('bold');
      sh.autoResizeColumns(1, 3);
    } catch (e) {}
  }

  return sh;
}

function getIntegrationSettings_() {
  var sh = ensureIntegrationSettingsSheet_();
  var lastRow = sh.getLastRow();
  var values = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, 2).getDisplayValues() : [];

  var cfg = {};
  Object.keys(MF_CFG_DEFAULTS).forEach(function (k) {
    cfg[k] = MF_CFG_DEFAULTS[k];
  });

  values.forEach(function (row) {
    var key = String(row[0] || '').trim();
    if (!key) return;
    cfg[key] = String(row[1] || '').trim();
  });

  return cfg;
}

function openManagerFilesRegistry_() {
  var cfg = getIntegrationSettings_();
  var sh = ensureManagerFilesRegistrySheet_(cfg);
  SpreadsheetApp.getActive().setActiveSheet(sh);
}

function ensureManagerFilesRegistrySheet_(cfg) {
  cfg = cfg || getIntegrationSettings_();

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(cfg.MANAGER_REGISTRY_SHEET || MF_CFG_DEFAULTS.MANAGER_REGISTRY_SHEET);
  if (!sh) sh = ss.insertSheet(cfg.MANAGER_REGISTRY_SHEET || MF_CFG_DEFAULTS.MANAGER_REGISTRY_SHEET);

  if (sh.getLastRow() < 2) {
    var rows = [
      ['Менеджер', 'Spreadsheet ID', 'Активен', 'Комментарий'],
      ['Любимова Анна', '', 'Да', 'Заполните ID рабочего файла менеджера'],
      ['Филлипов Сергей', '', 'Да', ''],
      ['Макеев Иван', '', 'Нет', '']
    ];
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.setFrozenRows(1);
    try {
      sh.getRange('A1:D1').setFontWeight('bold');
      sh.autoResizeColumns(1, 4);
    } catch (e) {}
  }

  return sh;
}

function getManagerFileRegistryRows_() {
  var cfg = getIntegrationSettings_();
  var sh = ensureManagerFilesRegistrySheet_(cfg);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  return values
    .map(function (r, idx) {
      return {
        row: idx + 2,
        managerName: String(r[0] || '').trim(),
        fileId: String(r[1] || '').trim(),
        isActive: /^да|yes|true|1$/i.test(String(r[2] || '').trim()),
        comment: String(r[3] || '').trim()
      };
    })
    .filter(function (x) {
      return x.managerName && x.fileId && x.isActive;
    });
}

/* ========================= Manager side: reference sync ========================= */

function syncReferenceSheetsFromMaster_() {
  var cfg = getIntegrationSettings_();
  var ui = SpreadsheetApp.getUi();

  if (String(cfg.APP_ROLE || '').toLowerCase() !== 'manager') {
    ui.alert(
      'Операция недоступна',
      'Функция "Обновить базы из мастер-файла" предназначена для рабочего файла менеджера.\n\n' +
      'Текущее значение APP_ROLE = ' + cfg.APP_ROLE,
      ui.ButtonSet.OK
    );
    return;
  }

  if (!cfg.MASTER_FILE_ID) {
    ui.alert(
      'Не заполнены настройки',
      'Укажите MASTER_FILE_ID на листе "Интеграция".',
      ui.ButtonSet.OK
    );
    return;
  }

  var ss = SpreadsheetApp.getActive();
  var master;
  try {
    master = SpreadsheetApp.openById(cfg.MASTER_FILE_ID);
  } catch (e) {
    ui.alert('Не удалось открыть мастер-файл:\n\n' + (e && e.message ? e.message : e));
    return;
  }

  var pairs = [
    { source: cfg.MASTER_DB_SHEET, target: cfg.LOCAL_DB_SHEET },
    { source: cfg.MASTER_LINKS_SHEET, target: cfg.LOCAL_LINKS_SHEET },
    { source: cfg.MASTER_MANAGERS_SHEET, target: cfg.LOCAL_MANAGERS_SHEET }
  ];

  var synced = [];
  pairs.forEach(function (pair) {
    syncOneSheetFromMaster_(master, ss, pair.source, pair.target);
    synced.push(pair.target);
  });

  var shouldRebuildPrice = /^yes|да|true|1$/i.test(String(cfg.AUTO_REBUILD_PRICE_AFTER_SYNC || '').trim());
  var priceRebuilt = false;
  if (shouldRebuildPrice && typeof buildPriceSheetWithOutlines === 'function') {
    try {
      buildPriceSheetWithOutlines();
      priceRebuilt = true;
    } catch (e) {
      ui.alert(
        'Справочники обновлены, но пересборка "Прайс" завершилась ошибкой:\n\n' +
        (e && e.message ? e.message : e)
      );
      return;
    }
  }

  ss.toast('Синхронизация завершена', 'Интеграция', 5);
  ui.alert(
    'Готово',
    'Обновлены листы:\n- ' + synced.join('\n- ') +
    (priceRebuilt ? '\n\nПрайс также пересобран локально.' : ''),
    ui.ButtonSet.OK
  );
}

function syncOneSheetFromMaster_(masterSs, localSs, sourceSheetName, targetSheetName) {
  var sourceSh = masterSs.getSheetByName(sourceSheetName);
  if (!sourceSh) throw new Error('В мастер-файле не найден лист "' + sourceSheetName + '".');

  var targetSh = localSs.getSheetByName(targetSheetName);
  if (!targetSh) targetSh = localSs.insertSheet(targetSheetName);

  resetSheetForSync_(targetSh);

  var lastRow = Math.max(sourceSh.getLastRow(), 1);
  var lastCol = Math.max(sourceSh.getLastColumn(), 1);

  ensureSheetSizeForSync_(targetSh, Math.max(sourceSh.getMaxRows(), lastRow), Math.max(sourceSh.getMaxColumns(), lastCol));

  var srcRange = sourceSh.getRange(1, 1, lastRow, lastCol);
  var dstRange = targetSh.getRange(1, 1, lastRow, lastCol);

  srcRange.copyTo(dstRange);

  copySheetDimensions_(sourceSh, targetSh, lastRow, lastCol);

  try {
    if (sourceSh.getFrozenRows() > 0) targetSh.setFrozenRows(sourceSh.getFrozenRows());
    if (sourceSh.getFrozenColumns() > 0) targetSh.setFrozenColumns(sourceSh.getFrozenColumns());
  } catch (e) {}
}

function resetSheetForSync_(sh) {
  try {
    sh.getImages().forEach(function (img) { img.remove(); });
  } catch (e) {}

  try {
    var f = sh.getFilter();
    if (f) f.remove();
  } catch (e) {}

  var maxR = Math.max(sh.getMaxRows(), 1);
  var maxC = Math.max(sh.getMaxColumns(), 1);
  var rng = sh.getRange(1, 1, maxR, maxC);

  try { rng.breakApart(); } catch (e) {}
  try { rng.clearDataValidations(); } catch (e) {}

  sh.clear();
  try { sh.clearConditionalFormatRules(); } catch (e) {}
}

function ensureSheetSizeForSync_(sh, minRows, minCols) {
  var curRows = sh.getMaxRows();
  var curCols = sh.getMaxColumns();

  if (curRows < minRows) sh.insertRowsAfter(curRows, minRows - curRows);
  if (curCols < minCols) sh.insertColumnsAfter(curCols, minCols - curCols);
}

function copySheetDimensions_(srcSh, dstSh, lastRow, lastCol) {
  for (var c = 1; c <= lastCol; c++) {
    try { dstSh.setColumnWidth(c, srcSh.getColumnWidth(c)); } catch (e) {}
    try {
      if (srcSh.isColumnHiddenByUser(c)) dstSh.hideColumns(c);
      else dstSh.showColumns(c);
    } catch (e) {}
  }

  for (var r = 1; r <= lastRow; r++) {
    try { dstSh.setRowHeight(r, srcSh.getRowHeight(r)); } catch (e) {}
    try {
      if (srcSh.isRowHiddenByUser(r)) dstSh.hideRows(r);
      else dstSh.showRows(r);
    } catch (e) {}
  }
}

/* ========================= Master side: summaries ========================= */

function rebuildAllMasterSummaries_() {
  rebuildMasterKpSummary_(true);
  rebuildMasterProdSummary_(true);
  SpreadsheetApp.getUi().alert('Готово', 'Оба сводных журнала пересобраны.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function rebuildMasterKpSummary_(silent) {
  var cfg = getIntegrationSettings_();
  ensureMasterRole_(cfg);

  var targetHeaders = ['Менеджер файла', 'Spreadsheet ID', 'Строка источника']
    .concat(_getKpLogHeadersCanonical_());

  rebuildSummarySheetFromManagerFiles_({
    targetSheetName: cfg.MASTER_SUMMARY_KP_SHEET,
    sourceSheetName: cfg.LOCAL_KP_LOG_SHEET,
    targetHeaders: targetHeaders
  });

  if (!silent) {
    SpreadsheetApp.getUi().alert('Готово', 'Сводный Журнал КП пересобран.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function rebuildMasterProdSummary_(silent) {
  var cfg = getIntegrationSettings_();
  ensureMasterRole_(cfg);

  var canonical = (typeof _getProdReqHeadersCanonical_ === 'function')
    ? _getProdReqHeadersCanonical_()
    : [
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

  var targetHeaders = ['Менеджер файла', 'Spreadsheet ID', 'Строка источника']
    .concat(canonical);

  rebuildSummarySheetFromManagerFiles_({
    targetSheetName: cfg.MASTER_SUMMARY_PROD_SHEET,
    sourceSheetName: cfg.LOCAL_PROD_LOG_SHEET,
    targetHeaders: targetHeaders
  });

  if (!silent) {
    SpreadsheetApp.getUi().alert('Готово', 'Сводный Журнал заявок пересобран.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ensureMasterRole_(cfg) {
  cfg = cfg || getIntegrationSettings_();
  if (String(cfg.APP_ROLE || '').toLowerCase() !== 'master') {
    throw new Error(
      'Функция доступна только в мастер-файле.\n' +
      'Текущее значение APP_ROLE = ' + cfg.APP_ROLE
    );
  }
}

function rebuildSummarySheetFromManagerFiles_(options) {
  var ss = SpreadsheetApp.getActive();
  var registry = getManagerFileRegistryRows_();

  var targetSheetName = options.targetSheetName;
  var sourceSheetName = options.sourceSheetName;
  var targetHeaders = options.targetHeaders || [];

  var targetSh = ss.getSheetByName(targetSheetName);
  if (!targetSh) targetSh = ss.insertSheet(targetSheetName);

  resetSheetForSync_(targetSh);
  ensureSheetSizeForSync_(targetSh, Math.max(1000, registry.length * 50 + 20), Math.max(targetHeaders.length, 1));

  targetSh.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
  targetSh.setFrozenRows(1);

  var rowsOut = [];

  registry.forEach(function (entry) {
    var sourceSs;
    try {
      sourceSs = SpreadsheetApp.openById(entry.fileId);
    } catch (e) {
      rowsOut.push(buildErrorRowForSummary_(targetHeaders, entry, 'Не удалось открыть файл: ' + (e && e.message ? e.message : e)));
      return;
    }

    var sourceSh = sourceSs.getSheetByName(sourceSheetName);
    if (!sourceSh) {
      rowsOut.push(buildErrorRowForSummary_(targetHeaders, entry, 'Не найден лист "' + sourceSheetName + '"'));
      return;
    }

    var data = readSheetAsObjects_(sourceSh);
    data.rows.forEach(function (obj) {
      var out = targetHeaders.map(function (h) {
        if (h === 'Менеджер файла') return entry.managerName;
        if (h === 'Spreadsheet ID') return entry.fileId;
        if (h === 'Строка источника') return obj.__sourceRow || '';
        return Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : '';
      });
      rowsOut.push(out);
    });
  });

  if (rowsOut.length) {
    targetSh.getRange(2, 1, rowsOut.length, targetHeaders.length).setValues(rowsOut);
  }

  try { targetSh.autoResizeColumns(1, targetHeaders.length); } catch (e) {}
  try {
    var lastRow = targetSh.getLastRow();
    if (lastRow >= 2) {
      targetSh.getRange(1, 1, lastRow, targetHeaders.length).createFilter();
    }
  } catch (e) {}
}

function buildErrorRowForSummary_(targetHeaders, entry, errorText) {
  return targetHeaders.map(function (h) {
    if (h === 'Менеджер файла') return entry.managerName;
    if (h === 'Spreadsheet ID') return entry.fileId;
    if (h === 'Строка источника') return '';
    if (h === 'Статус') return 'Ошибка синхронизации';
    if (h === 'Комментарий') return errorText;
    return '';
  });
}

function readSheetAsObjects_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };

  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  if (lastRow < 2) return { headers: headers, rows: [] };

  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rows = [];

  values.forEach(function (row, idx) {
    var isEmpty = row.every(function (cell) {
      return cell === '' || cell === null;
    });
    if (isEmpty) return;

    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var key = String(headers[c] || '').trim();
      if (!key) continue;
      obj[key] = row[c];
    }
    obj.__sourceRow = idx + 2;
    rows.push(obj);
  });

  return { headers: headers, rows: rows };
}
