/**
 * prod_request.gs — создание заявки на производство из выбранной строки "Журнал КП"
 *
 * Логика:
 * - Менеджер стоит на строке в "Журнал КП"
 * - Запускает функцию из меню (подключите в ваш существующий menu builder)
 * - Вводит номер заявки вручную
 * - В "Журнал заявок на производство" пишется запись
 * - В "Журнал КП" выбранной записи -> статус "Отправлена в производство"
 * - Для дублей (тот же КП № + Заказчик, но другой Drive File ID) -> статус "Аннулирована"
 */

function createProductionRequestFromSelectedKp() {
  const ss = SpreadsheetApp.getActive();

  const kpLog = ensureKpLogSchema_(ss);
  const prodLog = ensureProductionRequestLogSchema_(ss);

  const activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getName() !== kpLog.getName()) {
    SpreadsheetApp.getUi().alert(
      'Для создания заявки на производство перейдите на лист "' + kpLog.getName() + '" и выделите нужную строку.'
    );
    return;
  }

  const row = activeSheet.getActiveRange() ? activeSheet.getActiveRange().getRow() : 0;
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Выберите строку в "Журнал КП" (не шапку).');
    return;
  }

  const kpRow = getKpLogRowAsObject_(kpLog, row);

  // Обязательные поля для заявки на производство
  const driveFileId = String(kpRow['Drive File ID'] || '').trim();
  const kpNo = String(kpRow['КП №'] || '').trim();
  const customer = String(kpRow['Заказчик'] || '').trim();
  const status = String(kpRow['Статус'] || '').trim() || 'Новая';

  if (!driveFileId) {
    SpreadsheetApp.getUi().alert(
      'Нельзя создать заявку на производство.\n\nВ выбранной строке "Журнал КП" не заполнен "Drive File ID".'
    );
    return;
  }

  if (!kpNo) {
    SpreadsheetApp.getUi().alert('Нельзя создать заявку на производство: не заполнен "КП №" в журнале.');
    return;
  }

  if (!customer) {
    SpreadsheetApp.getUi().alert('Нельзя создать заявку на производство: не заполнен "Заказчик" в журнале.');
    return;
  }

  if (status === 'Отправлена в производство') {
    SpreadsheetApp.getUi().alert('По этой записи уже создана заявка на производство.');
    return;
  }

  if (status === 'Аннулирована') {
    SpreadsheetApp.getUi().alert('Выбранная запись имеет статус "Аннулирована". Создание заявки запрещено.');
    return;
  }

  // Проверка, что по данному КП+заказчику уже не отправлена другая запись в производство
  const existingSent = findAnotherSentProductionByKpAndCustomer_(kpLog, kpNo, customer, driveFileId);
  if (existingSent) {
    SpreadsheetApp.getUi().alert(
      'Заявка на производство уже создана по дублю этого КП.\n\n' +
      'Строка журнала КП: ' + existingSent.row + '\n' +
      'Статус: Отправлена в производство\n\n' +
      'Для текущей строки создание заявки запрещено.'
    );
    return;
  }

  const previewText = buildProdRequestPreviewText_(kpRow);

  // Подтверждение
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Создать заявку на производство',
    previewText + '\n\nПродолжить?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Ввод номера заявки
  const prompt = ui.prompt(
    'Номер заявки на производство',
    'Введите номер заявки (можно вручную):',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  const requestNo = String(prompt.getResponseText() || '').trim();
  if (!requestNo) {
    ui.alert('Создание отменено: номер заявки не введён.');
    return;
  }

  // Повторная защита от дубля в журнале заявок (по номеру заявки + Drive File ID)
  const dupProd = findProdRequestDuplicate_(prodLog, requestNo, driveFileId);
  if (dupProd) {
    ui.alert(
      'Такая заявка уже есть в "Журнал заявок на производство".\n\n' +
      'Строка: ' + dupProd.row + '\n' +
      'Номер заявки: ' + requestNo
    );
    return;
  }

  // Пишем в журнал заявок на производство
  const prodRow = buildProdRequestLogRow_(kpRow, requestNo);
  const prodRowNum = appendProductionRequestToLog_(prodLog, prodRow);

  // Статусы в Журнале КП:
  // - текущая строка = "Отправлена в производство"
  setKpLogStatusByRow_(kpLog, row, 'Отправлена в производство');

  // - дубли (тот же КП № + Заказчик, другой Drive File ID) = "Аннулирована"
  const cancelledCount = cancelDuplicateKpRows_(kpLog, {
    kpNo: kpNo,
    customer: customer,
    exceptDriveFileId: driveFileId,
    keepRow: row
  });

  ui.alert(
    'Заявка на производство создана.\n\n' +
    'Строка в журнале заявок: ' + prodRowNum + '\n' +
    'Номер заявки: ' + requestNo + '\n' +
    'Статус текущей записи КП: Отправлена в производство\n' +
    'Аннулировано дублей: ' + cancelledCount
  );
}

/* ========================= Schema: Журнал заявок на производство ========================= */

function _prodReqSheetName_() {
  return (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.PROD_REQUEST_LOG)
    ? CFG.SHEETS.PROD_REQUEST_LOG
    : 'Журнал заявок на производство';
}

function _prodReqHeaders_() {
  if (typeof CFG !== 'undefined' && CFG.SCHEMAS && Array.isArray(CFG.SCHEMAS.PROD_REQUEST_LOG) && CFG.SCHEMAS.PROD_REQUEST_LOG.length) {
    return CFG.SCHEMAS.PROD_REQUEST_LOG.slice();
  }
  return [
    'Дата/время создания',
    'Номер заявки на производство',
    'КП №',
    'Дата КП',
    'Заказчик',
    'Адрес заказчика',
    'Менеджер',
    'Телефон',
    'Итого к оплате, руб',
    'PDF URL (Drive)',
    'Drive File ID',
    'Строка в Журнал КП',
    'Позиции (JSON)',
    'Комментарий'
  ];
}

function ensureProductionRequestLogSchema_(ss) {
  const name = _prodReqSheetName_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const need = _prodReqHeaders_();
  const lastCol = Math.max(sh.getLastColumn(), 1);

  const current = sh.getLastRow() >= 1
    ? sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    : [];

  const currentNorm = current.map(_normProdHeader_);
  const needNorm = need.map(_normProdHeader_);

  if (current.join('').trim() === '') {
    sh.getRange(1, 1, 1, need.length).setValues([need]);
    return sh;
  }

  // Добавляем недостающие справа
  need.forEach((h, i) => {
    if (currentNorm.indexOf(needNorm[i]) < 0) {
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(h);
    }
  });

  return sh;
}

function _normProdHeader_(s) {
  return String(s || '')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/* ========================= Чтение строки Журнал КП ========================= */

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

function buildProdRequestPreviewText_(kpRow) {
  return [
    'Будет создана заявка на производство по записи "Журнал КП":',
    '',
    'Строка журнала КП: ' + (kpRow.__row || ''),
    'КП №: ' + String(kpRow['КП №'] || ''),
    'Дата КП: ' + String(kpRow['Дата КП__display'] || kpRow['Дата КП'] || ''),
    'Заказчик: ' + String(kpRow['Заказчик'] || ''),
    'Адрес: ' + String(kpRow['Адрес заказчика'] || ''),
    'Менеджер: ' + String(kpRow['Менеджер'] || ''),
    'Телефон: ' + String(kpRow['Телефон'] || ''),
    'Итого к оплате: ' + String(kpRow['Итого к оплате, руб__display'] || kpRow['Итого к оплате, руб'] || ''),
    'Drive File ID: ' + String(kpRow['Drive File ID'] || ''),
    '',
    'После создания:',
    '• текущая запись получит статус "Отправлена в производство";',
    '• дубли (тот же КП № + Заказчик) будут помечены "Аннулирована".'
  ].join('\n');
}

/* ========================= Запись в журнал заявок ========================= */

function buildProdRequestLogRow_(kpRow, requestNo) {
  return {
    'Дата/время создания': new Date(),
    'Номер заявки на производство': requestNo,
    'КП №': kpRow['КП №'] || '',
    'Дата КП': kpRow['Дата КП'] || '',
    'Заказчик': kpRow['Заказчик'] || '',
    'Адрес заказчика': kpRow['Адрес заказчика'] || '',
    'Менеджер': kpRow['Менеджер'] || '',
    'Телефон': kpRow['Телефон'] || '',
    'Итого к оплате, руб': kpRow['Итого к оплате, руб'] || '',
    'PDF URL (Drive)': kpRow['PDF URL (Drive)'] || '',
    'Drive File ID': kpRow['Drive File ID'] || '',
    'Строка в Журнал КП': kpRow.__row || '',
    'Позиции (JSON)': kpRow['Позиции (JSON)'] || '',
    'Комментарий': ''
  };
}

function appendProductionRequestToLog_(sh, rowObj) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const row = headers.map(h => rowObj.hasOwnProperty(h) ? rowObj[h] : '');
  sh.appendRow(row);

  const rowNum = sh.getLastRow();

  // Форматы
  const hdrNorm = headers.map(_normProdHeader_);
  const cMoney = hdrNorm.indexOf(_normProdHeader_('Итого к оплате, руб')) + 1;
  if (cMoney) {
    try { sh.getRange(rowNum, cMoney).setNumberFormat('#,##0.00'); } catch (e) {}
  }

  return rowNum;
}

function findProdRequestDuplicate_(prodLogSheet, requestNo, driveFileId) {
  const lastRow = prodLogSheet.getLastRow();
  if (lastRow < 2) return null;

  const lastCol = prodLogSheet.getLastColumn();
  const headers = prodLogSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(_normProdHeader_);

  const cReq = headers.indexOf(_normProdHeader_('Номер заявки на производство')) + 1;
  const cId = headers.indexOf(_normProdHeader_('Drive File ID')) + 1;
  if (!cReq || !cId) return null;

  const n = lastRow - 1;
  const reqVals = prodLogSheet.getRange(2, cReq, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const idVals = prodLogSheet.getRange(2, cId, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());

  for (let i = n - 1; i >= 0; i--) {
    if (reqVals[i] === requestNo && idVals[i] === driveFileId) {
      return { row: i + 2 };
    }
  }
  return null;
}

/* ========================= Статусы в Журнал КП ========================= */

function setKpLogStatusByRow_(kpLogSheet, row, statusValue) {
  const statusCol = getKpLogStatusColumn_(kpLogSheet);
  if (!statusCol) throw new Error('Не найдена колонка "Статус" в "Журнал КП".');
  kpLogSheet.getRange(row, statusCol).setValue(statusValue);
}

function getKpLogStatusColumn_(kpLogSheet) {
  const headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  const norm = headers.map(h => String(h || '').trim().toLowerCase());
  return norm.indexOf('статус') + 1;
}

function findAnotherSentProductionByKpAndCustomer_(kpLogSheet, kpNo, customer, currentDriveFileId) {
  const lastRow = kpLogSheet.getLastRow();
  if (lastRow < 2) return null;

  const headers = kpLogSheet.getRange(1, 1, 1, kpLogSheet.getLastColumn()).getDisplayValues()[0];
  const norm = headers.map(h => String(h || '').trim().toLowerCase());

  const cKp = norm.indexOf('кп №') + 1;
  const cCust = norm.indexOf('заказчик') + 1;
  const cStatus = norm.indexOf('статус') + 1;
  const cDrive = norm.indexOf('drive file id') + 1;

  if (!cKp || !cCust || !cStatus || !cDrive) return null;

  const n = lastRow - 1;
  const kpVals = kpLogSheet.getRange(2, cKp, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const custVals = kpLogSheet.getRange(2, cCust, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const statusVals = kpLogSheet.getRange(2, cStatus, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const idVals = kpLogSheet.getRange(2, cDrive, n, 1).getDisplayValues().map(r => String(r[0] || '').trim());

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
  const norm = headers.map(h => String(h || '').trim().toLowerCase());

  const cKp = norm.indexOf('кп №') + 1;
  const cCust = norm.indexOf('заказчик') + 1;
  const cStatus = norm.indexOf('статус') + 1;
  const cDrive = norm.indexOf('drive file id') + 1;

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

    // Не трогаем уже отправленные (но обычно таких не будет из-за предварительной проверки)
    if (vStatus === 'Отправлена в производство') continue;

    kpLogSheet.getRange(row, cStatus).setValue('Аннулирована');
    changed++;
  }

  return changed;
}