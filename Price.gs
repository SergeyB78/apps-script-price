/**
 * Price.gs — формирование листа "Прайс" из "БД Оборудования"
 * Поведение сохранено как в исходной версии:
 * - пересоздаём лист "Прайс" (удаляем и создаём заново)
 * - перемещаем "Прайс" на 2-е место
 * - группирующие строки "Тип / Серия" в колонке B, merge B:I
 * - outline (+/-) и сворачиваем группы
 * - UID скрываем
 * - копируем ширины колонок и форматы из БД
 *
 * onOpen() находится в main.gs (НЕ трогаем).
 */

const PRICE = {
  SOURCE_SHEET: () => CFG.SHEETS.DB,
  PRICE_SHEET: () => CFG.SHEETS.PRICE,
  ACTIVE_STATUS: () => CFG.DB_VALUES.ACTIVE_STATUS,

  COL_UID: () => CFG.DB_HEADERS.UID,
  COL_TYPE: () => CFG.DB_HEADERS.EQUIP_TYPE,
  COL_SERIES: () => CFG.DB_HEADERS.SERIES,
  COL_ART: () => CFG.DB_HEADERS.SKU,
  COL_VIEW1: () => CFG.DB_HEADERS.IMG1,
  COL_VIEW2: () => CFG.DB_HEADERS.IMG2,
  COL_NAME: () => CFG.DB_HEADERS.NAME,
  COL_UNIT: () => CFG.DB_HEADERS.UNIT,
  COL_PRICE: () => CFG.DB_HEADERS.COST,
  COL_NOTE: () => CFG.DB_HEADERS.NOTE,
  COL_STATUS: () => CFG.DB_HEADERS.STATUS,

  ICON: "🖼️",
};

/**
 * ВАЖНО: оставляем это имя, потому что у вас оно где-то уже используется
 * (меню/кнопка/назначенный скрипт).
 */
function buildPriceSheetWithOutlines() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    buildPrice_();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** алиас на всякий случай */
function buildPrice() {
  buildPriceSheetWithOutlines();
}

function buildPrice_() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(PRICE.SOURCE_SHEET());
  if (!src) throw new Error(`Не найден лист "${PRICE.SOURCE_SHEET()}".`);

  const values = src.getDataRange().getValues();
  if (values.length < 2) return;

  const header = values[0].map(v => String(v || "").trim());
  const idx = makeIndex_(header);

  const required = [
    PRICE.COL_UID(), PRICE.COL_TYPE(), PRICE.COL_SERIES(), PRICE.COL_ART(),
    PRICE.COL_VIEW1(), PRICE.COL_VIEW2(), PRICE.COL_NAME(), PRICE.COL_UNIT(),
    PRICE.COL_PRICE(), PRICE.COL_NOTE(), PRICE.COL_STATUS()
  ];
  for (const c of required) {
    if (idx[c] === undefined) {
      throw new Error(`На листе "${PRICE.SOURCE_SHEET()}" не найдена обязательная колонка: "${c}".`);
    }
  }

  const lastRow = src.getLastRow();
  const view1RT = src.getRange(2, idx[PRICE.COL_VIEW1()] + 1, lastRow - 1, 1).getRichTextValues();
  const view2RT = src.getRange(2, idx[PRICE.COL_VIEW2()] + 1, lastRow - 1, 1).getRichTextValues();

  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status = String(row[idx[PRICE.COL_STATUS()]] || "").trim();
    if (status !== PRICE.ACTIVE_STATUS()) continue;

    const uid = String(row[idx[PRICE.COL_UID()]] || "").trim();
    if (!uid) continue;

    const type = String(row[idx[PRICE.COL_TYPE()]] || "").trim();
    const series = String(row[idx[PRICE.COL_SERIES()]] || "").trim();
    const hasSeries = series ? 1 : 0;

    const v1Url = extractUrl_(view1RT[r - 1]?.[0], row[idx[PRICE.COL_VIEW1()]]);
    const v2Url = extractUrl_(view2RT[r - 1]?.[0], row[idx[PRICE.COL_VIEW2()]]);

    items.push({
      uid, type, series, hasSeries,
      art: String(row[idx[PRICE.COL_ART()]] || "").trim(),
      view1Url: v1Url,
      view2Url: v2Url,
      name: row[idx[PRICE.COL_NAME()]] || "",
      unit: row[idx[PRICE.COL_UNIT()]] || "",
      price: row[idx[PRICE.COL_PRICE()]] || "",
      note: row[idx[PRICE.COL_NOTE()]] || ""
    });
  }

  items.sort((a, b) => {
    const t = a.type.localeCompare(b.type, "ru");
    if (t) return t;
    if (a.hasSeries !== b.hasSeries) return b.hasSeries - a.hasSeries;
    const s = a.series.localeCompare(b.series, "ru");
    if (s) return s;
    return a.art.localeCompare(b.art, "ru");
  });

  // пересоздаём лист "Прайс"
  let sh = ss.getSheetByName(PRICE.PRICE_SHEET());
  if (!sh) {
    sh = ss.insertSheet(PRICE.PRICE_SHEET());
  } else {
    const oldIdx = sh.getIndex();
    ss.deleteSheet(sh);
    sh = ss.insertSheet(PRICE.PRICE_SHEET(), Math.max(0, oldIdx - 1));
  }

  // делаем "Прайс" 2-м листом и активируем
  ss.setActiveSheet(sh);
  try { ss.moveActiveSheet(2); } catch (e) {}
  ss.setActiveSheet(sh);

  // Шапка прайса (без колонок Вид оборудования/Серия)
  const outHeader = [[
    "UID",
    "Артикул",
    "Вид 1",
    "Вид 2",
    "Наименование изделия/ размеры",
    "Ед. изм.",
    "Стоимость оборудования",
    "Примечание",
    "Кол-во"
  ]];

  sh.getRange(1, 1, 1, outHeader[0].length).setValues(outHeader);
  sh.setFrozenRows(1);

  // Данные + блоки
  const out = [];
  const rowMeta = [];
  let prevKey = null;

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const seriesLabel = it.series ? it.series : "(без серии)";
    const key = `${it.type} / ${seriesLabel}`;

    if (key !== prevKey) {
      out.push(["", key, "", "", "", "", "", "", ""]);
      rowMeta.push({ isGroup: true });
      prevKey = key;
    }

    out.push([it.uid, it.art, "", "", it.name, it.unit, it.price, it.note, ""]);
    rowMeta.push({ isGroup: false, itemIndex: i });
  }

  if (out.length) sh.getRange(2, 1, out.length, outHeader[0].length).setValues(out);

  // Маппинг БД -> Прайс (форматы + ширины)
  const mapCols = [
    { srcCol: idx[PRICE.COL_UID()] + 1,   dstCol: 1 }, // UID
    { srcCol: idx[PRICE.COL_ART()] + 1,   dstCol: 2 }, // Артикул
    { srcCol: idx[PRICE.COL_VIEW1()] + 1, dstCol: 3 }, // Вид 1
    { srcCol: idx[PRICE.COL_VIEW2()] + 1, dstCol: 4 }, // Вид 2
    { srcCol: idx[PRICE.COL_NAME()] + 1,  dstCol: 5 }, // Наименование
    { srcCol: idx[PRICE.COL_UNIT()] + 1,  dstCol: 6 }, // Ед.
    { srcCol: idx[PRICE.COL_PRICE()] + 1, dstCol: 7 }, // Стоимость
    { srcCol: idx[PRICE.COL_NOTE()] + 1,  dstCol: 8 }, // Примечание
  ];

  // ширины
  for (const m of mapCols) {
    sh.setColumnWidth(m.dstCol, src.getColumnWidth(m.srcCol));
  }
  sh.setColumnWidth(9, src.getColumnWidth(idx[PRICE.COL_UNIT()] + 1)); // Кол-во

  // форматы (шапка + строка 2 как шаблон)
  const dataRows = out.length;
  if (dataRows > 0) {
    for (const m of mapCols) {
      src.getRange(1, m.srcCol, 1, 1).copyTo(
        sh.getRange(1, m.dstCol, 1, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
    }
    src.getRange(1, idx[PRICE.COL_NOTE()] + 1, 1, 1).copyTo(
      sh.getRange(1, 9, 1, 1),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

    for (const m of mapCols) {
      src.getRange(2, m.srcCol, 1, 1).copyTo(
        sh.getRange(2, m.dstCol, dataRows, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
    }
    sh.getRange(2, 9, dataRows, 1).setNumberFormat("0.##").setHorizontalAlignment("center");
  }

  // 🖼️ иконки-ссылки
  if (dataRows > 0) {
    const rt1 = [];
    const rt2 = [];
    for (let r = 0; r < rowMeta.length; r++) {
      const m = rowMeta[r];
      if (m.isGroup) {
        rt1.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
        rt2.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
      } else {
        const it = items[m.itemIndex];
        rt1.push([makeIconLink_(it.view1Url)]);
        rt2.push([makeIconLink_(it.view2Url)]);
      }
    }
    sh.getRange(2, 3, dataRows, 1).setRichTextValues(rt1);
    sh.getRange(2, 4, dataRows, 1).setRichTextValues(rt2);
  }

  // блок-строки: merge B:I + стиль
  const startRow = 2;
  const groupRows = [];
  for (let i = 0; i < rowMeta.length; i++) if (rowMeta[i].isGroup) groupRows.push(startRow + i);

  if (groupRows.length) {
    for (const rr of groupRows) sh.getRange(rr, 2, 1, 8).merge(); // B..I
    sh.getRangeList(groupRows.map(rr => `B${rr}:I${rr}`))
      .setFontWeight("bold")
      .setBackground("#EDEDED")
      .setHorizontalAlignment("left")
      .setBorder(true, true, true, true, false, false);
  }

  // outline (+/-) и сразу свернуть
  if (dataRows > 0) {
    let cursor = 0;
    while (cursor < rowMeta.length) {
      if (!rowMeta[cursor].isGroup) { cursor++; continue; }

      const startItemsIdx = cursor + 1;
      if (startItemsIdx >= rowMeta.length) break;
      if (rowMeta[startItemsIdx].isGroup) { cursor++; continue; }

      let endItemsIdx = startItemsIdx;
      while (endItemsIdx + 1 < rowMeta.length && !rowMeta[endItemsIdx + 1].isGroup) endItemsIdx++;

      const startItemsRow = startRow + startItemsIdx;
      const numRows = endItemsIdx - startItemsIdx + 1;

      sh.getRange(startItemsRow, 1, numRows, 1).shiftRowGroupDepth(1);
      try {
        const rg = sh.getRowGroup(startItemsRow, 1);
        if (rg) rg.collapse();
      } catch (e) {}

      cursor = endItemsIdx + 1;
    }
  }

  // UID скрыть
  sh.hideColumn(sh.getRange("A:A"));

  ss.setActiveSheet(sh);
}

function makeIndex_(hdr) {
  const m = {};
  hdr.forEach((h, i) => { if (h) m[h] = i; });
  return m;
}

function extractUrl_(rich, fallbackCellValue) {
  try {
    if (rich) {
      const u = rich.getLinkUrl && rich.getLinkUrl();
      if (u) return u;

      const runs = rich.getRuns && rich.getRuns();
      if (runs && runs.length) {
        for (const run of runs) {
          const ru = run.getLinkUrl && run.getLinkUrl();
          if (ru) return ru;
        }
      }
    }
  } catch (e) {}

  const s = String(fallbackCellValue || "").trim();
  const m = s.match(/https?:\/\/[^\s")]+/i);
  return m ? m[0] : "";
}

function makeIconLink_(url) {
  const u = String(url || "").trim();
  if (!u) return SpreadsheetApp.newRichTextValue().setText("").build();
  return SpreadsheetApp.newRichTextValue().setText(PRICE.ICON).setLinkUrl(u).build();
}
