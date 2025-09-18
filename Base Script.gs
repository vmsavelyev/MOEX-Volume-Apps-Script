//Оптимизированный скрипт который паралелить заполнение таблицы

function fillMissingValuesBatchParallel() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Объемы');
  if (!sheet) {
    Logger.log("Лист 'Объемы' не найден");
    return;
  }

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  if (lastColumn < 2 || lastRow < 2) {
    Logger.log("Недостаточно данных");
    return;
  }

  // Даты из первой строки, начиная со второго столбца
  const datesTableFormat = sheet.getRange(1, 2, 1, lastColumn - 1).getValues()[0];

  // Тикеры из первого столбца, начиная со второй строки
  const tickersRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const tickers = tickersRange.getValues().flat().filter(t => t && t.toString().trim() !== '');

  // Функция конвертации даты в формат YYYY-MM-DD
  function convertDateFormat(dateValue) {
    if (!dateValue) return null;
    if (Object.prototype.toString.call(dateValue) === '[object Date]') {
      if (isNaN(dateValue.getTime())) return null;
      const dd = ('0' + dateValue.getDate()).slice(-2);
      const mm = ('0' + (dateValue.getMonth() + 1)).slice(-2);
      const yyyy = dateValue.getFullYear();
      return `${yyyy}-${mm}-${dd}`;
    }
    if (typeof dateValue === 'string') {
      const parts = dateValue.split('.');
      if (parts.length !== 3) return null;
      return `${parts[2]}-${parts[1]}-${parts[0]}`;
    }
    return null;
  }

  const datesApiFormat = datesTableFormat.map(d => convertDateFormat(d));
  const fromDate = datesApiFormat[0];
  const tillDate = datesApiFormat[datesApiFormat.length - 1];

  const scriptProperties = PropertiesService.getScriptProperties();
  let currentIndex = Number(scriptProperties.getProperty('currentIndex')) || 0;

  const batchSize = 100;
  const endIndex = Math.min(currentIndex + batchSize, tickers.length);

  // Формируем массив запросов для fetchAll
  const requests = [];
  for (let i = currentIndex; i < endIndex; i++) {
    const ticker = tickers[i];
    const url = `https://iss.moex.com/iss/history/engines/stock/markets/shares/boards/TQBR/securities/${ticker}.json?from=${fromDate}&till=${tillDate}`;
    requests.push({ url: url, muteHttpExceptions: true });
  }

  // Выполняем параллельные запросы
  const responses = UrlFetchApp.fetchAll(requests);

  for (let i = 0; i < responses.length; i++) {
    const response = responses[i];
    const tickerIndex = currentIndex + i;
    const ticker = tickers[tickerIndex];

    if (response.getResponseCode() !== 200) {
      Logger.log(`Ошибка HTTP ${response.getResponseCode()} для тикера ${ticker}`);
      continue;
    }

    const json = JSON.parse(response.getContentText());
    const columns = json.history.columns;
    const data = json.history.data;

    if (!data || data.length === 0) {
      Logger.log(`Нет данных для тикера ${ticker}`);
      continue;
    }

    const valueIndex = columns.indexOf('VALUE');
    const dateIndex = columns.indexOf('TRADEDATE');
    if (valueIndex === -1 || dateIndex === -1) {
      Logger.log(`Не найден столбец VALUE или TRADEDATE для тикера ${ticker}`);
      continue;
    }

    // Создаём словарь дата -> value
    const valuesByDate = {};
    data.forEach(row => {
      valuesByDate[row[dateIndex]] = row[valueIndex];
    });

    const row = tickerIndex + 2; // строка в таблице для текущего тикера

    // Получаем текущие значения всей строки (по всем датам)
    const currentRowValues = sheet.getRange(row, 2, 1, datesApiFormat.length).getValues()[0];

    // Формируем массив значений для записи, не затирая уже заполненные ячейки
    const rowValues = datesApiFormat.map((dateKey, idx) => {
      if (currentRowValues[idx] !== '' && currentRowValues[idx] !== null) {
        return currentRowValues[idx];
      }
      const val = valuesByDate[dateKey];
      return (val !== undefined && val !== null) ? val : '';
    });

    // Записываем всю строку за один вызов
    sheet.getRange(row, 2, 1, datesApiFormat.length).setValues([rowValues]);
  }

  // Обновляем индекс и создаём триггер для следующей пачки
  if (endIndex >= tickers.length) {
    Logger.log('Все тикеры обработаны.');
    scriptProperties.deleteProperty('currentIndex');
  } else {
    scriptProperties.setProperty('currentIndex', endIndex.toString());
    deleteTimeDrivenTriggers();
    ScriptApp.newTrigger('fillMissingValuesBatchParallel')
      .timeBased()
      .after(5 * 1000)
      .create();
    Logger.log(`Обработано тикеров: ${endIndex}. Следующий запуск через 5 секунд.`);
  }
}

function deleteTimeDrivenTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fillMissingValuesBatchParallel') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}