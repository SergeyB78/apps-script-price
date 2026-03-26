function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Прайс")
    // НЕ трогаем вашу рабочую функцию — просто вызываем её из меню
    .addItem("Сформировать прайс", "buildPriceSheetWithOutlines")
    .addSeparator()
    .addItem("Сформировать КП", "buildKP")
    .addItem("Экспорт КП в PDF", "exportKpPdfAndLog")
    .addItem('Восстановить КП из журнала', 'restoreKPFromJournal')
    .addSeparator()
    .addItem(
      "Создать заявку на производство (по строке Журнала КП)",
      "createProductionRequestFromSelectedKp"
    )
    .addToUi();
}