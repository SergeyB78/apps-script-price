/**
 * settings.gs — единая конфигурация проекта.
 *
 * Принцип:
 * 1) Все "магические" строки/имена/параметры — сюда.
 * 2) В рабочем коде постепенно заменяем литералы на CFG.*.
 * 3) Пока что ID (Drive) храним в коде. Позже перенесём в Script Properties.
 */

const CFG = Object.freeze({
  /** Имена листов */
  SHEETS: Object.freeze({
    DB: 'БД Оборудования',
    PRICE: 'Прайс',
    KP: 'КП',
    KP_LOG: 'Журнал КП',
  }),

  /** Заголовки (колонки) в БД */
  DB_HEADERS: Object.freeze({
    UID: 'UID',
    EQUIP_TYPE: 'Вид оборудования',
    SERIES: 'Серия',
    SKU: 'Артикул',
    IMG1: 'Вид 1',
    IMG2: 'Вид 2',
    NAME: 'Наименование изделия/ размеры',
    UNIT: 'Ед. изм.',
    COST: 'Стоимость оборудования',
    STATUS: 'Статус позиции',
    NOTE: 'Примечание',
  }),

  /** Значения, которые встречаются в БД */
  DB_VALUES: Object.freeze({
    ACTIVE_STATUS: 'Активная позиция',
  }),

  /** ID (Пока в коде. Потом лучше перенесём в Script Properties.) */
  IDS: Object.freeze({
    HEADER_FILE_ID: '1E-6KvJH6CPFkEHgG0nLX_NE7rYlm67bw',
    DRIVE_FOLDER_ID: '1o8lqVv3DlUe4e3bMpWKvNZncf5r_2xDD',
  }),

  /** Ключи Script Properties (на будущее, когда уберём ID из GitHub) */
  PROPS: Object.freeze({
    DRIVE_FOLDER_ID: 'DRIVE_FOLDER_ID',
    HEADER_FILE_ID: 'HEADER_FILE_ID',
  }),

  /** Настройки КП (используются в КП.gs) */
  KP: Object.freeze({
    COL_END: 11, // K
    DEFAULT_COL_WIDTH: 100,
    DEFAULT_ROW_HEIGHT: 21,
    FONT_SIZE: 20,

    // ВНИМАНИЕ: позже лучше заменить на NamedRanges, чтобы строки не "плыли".
    PARAM_ROWS: Object.freeze({
      DISCOUNT: 19, // D19
      INSTALL: 20,  // D20
      DELIVERY: 21, // D21
    }),

    NUMBER_FORMATS: Object.freeze({
      MONEY: '#,##0.00',
      INT: '#,##0',
      PCT: '0.00',
      TEXT: '@',
    }),
  }),
});

/**
 * Заготовка на будущее:
 * Получить Script Property.
 * required=true -> кидает понятную ошибку, если свойства нет.
 */
function getScriptProp_(key, required = true) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if ((v === null || v === undefined || String(v).trim() === '') && required) {
    throw new Error(
      `Не задан Script Property: "${key}". ` +
      `Задай в Apps Script → Project Settings → Script properties.`
    );
  }
  return v;
}
