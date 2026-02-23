/*************** kp_pdf_export.gs ***************/
/**
 * Модуль PDF:
 * - validateRequiredKpFields_
 * - exportSheetToPdfBlob_
 * - savePdfToDriveFolder_
 * - поиск/скрытие строк блоков (findRowByExactText_, hide/show rows helpers)
 * - showLinksDialog_
 */

/* ===================== VALIDATION ===================== */

function validateRequiredKpFields_(sh) {
  const requiredLabels = [
    'Наименование Заказчика',
    'Адрес Заказчика',
    'Менеджер',
    'Коммерческое предложение №',
    'Дата КП',
  ];

  const missing = [];

  // 1) Проверка обязательных полей (значение берём из колонки D по подписи)
  for (const label of requiredLabels) {
    const v = findValueByLabelInColD_(sh, label); // утилита из kp_utils.gs
    if (isEmptyValue_(v)) missing.push(label);
  }

  // 2) Проверка корзины: должно быть > 0 позиций
  // Используем существующую функцию, которая уже читает корзину КП. :contentReference[oaicite:1]{index=1}
  const cartLabel = 'Количество позиций в корзине (> 0)';
  try {
    if (typeof extractCartAsJson_ === 'function') {
      const cart = extractCartAsJson_(sh);
      const cnt = (cart && Array.isArray(cart.items)) ? cart.items.length : 0;
      if (cnt <= 0) missing.push(cartLabel);
    } else {
      // если вдруг кто-то удалил kp_log.gs / функцию — считаем, что корзина невалидна
      missing.push(cartLabel);
    }
  } catch (e) {
    missing.push(cartLabel);
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

/* ===================== UI ===================== */

function showLinksDialog_(fileUrl, downloadUrl) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;font-size:14px;">
      <p>Скачать (попадёт в “Загрузки” браузера): <a href="${downloadUrl}" target="_blank">скачать PDF</a></p>
    </div>`
  ).setWidth(420).setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(html, 'КП → PDF');
}
