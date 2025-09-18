//Расчет дельты но только при нормально сортировки строк 1) Тикет 2) Дельта 3) Тикет 4) Дельта

function fillMissingValuesBatchParallel() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Объемы');
  if (!sheet) {
    Logger.log("Лист 'Объемы' не найден");
    return;
  }

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  if (lastColumn < 3 || lastRow < 2) { // даты теперь с 3-го столбца
    Logger.log("Недостаточно данных");
    return;
  }

  // Даты из первой строки, начиная с 3-го столбца
  const datesTableFormat = sheet.getRange(1, 3, 1, lastColumn - 2).getValues()[0];

  // Тикеры из первого столбца, начиная со второй строки
  const tickersRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const tickers = tickersRange.getValues().flat().filter(t => t && t.toString().trim() !== '');

  // Признаки строк (Объем или Дельта) из столбца 2, начиная со второй строки
  const rowTypesRange = sheet.getRange(2, 2, lastRow - 1, 1);
  const rowTypes = rowTypesRange.getValues().flat();

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
    const rowType = rowTypes[tickerIndex]; // тип строки: Объем или Дельта

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
    const currentRowValues = sheet.getRange(row, 3, 1, datesApiFormat.length).getValues()[0];

    let rowValues;

    if (rowType.toString().toLowerCase() === 'объем') {
      // Для строк "Объем" — записываем значения VALUE из API, не затирая заполненные ячейки
      rowValues = datesApiFormat.map((dateKey, idx) => {
        if (currentRowValues[idx] !== '' && currentRowValues[idx] !== null) {
          return currentRowValues[idx];
        }
        const val = valuesByDate[dateKey];
        return (val !== undefined && val !== null) ? val : '';
      });
    } else if (rowType.toString().toLowerCase() === 'дельта') {
      // Для строк "Дельта" — считаем % изменение по данным из строки "Объем" (предыдущая строка)
      if (tickerIndex === 0) {
        // Если это первая строка, нет предыдущей, записываем пустые значения
        rowValues = new Array(datesApiFormat.length).fill('');
      } else {
        // Получаем значения объема из предыдущей строки
        const volumeRowValues = sheet.getRange(row - 1, 3, 1, datesApiFormat.length).getValues()[0];
        rowValues = [];

        for (let idx = 0; idx < datesApiFormat.length; idx++) {
          if (idx === 0) {
            // Для первой даты дельта не считается
            rowValues.push('');
            continue;
          }
          const prevVal = parseFloat(volumeRowValues[idx - 1]);
          const currVal = parseFloat(volumeRowValues[idx]);

          if (isNaN(prevVal) || prevVal === 0 || isNaN(currVal)) {
            rowValues.push('');
          } else {
            const delta = ((currVal - prevVal) / prevVal) * 100;
            rowValues.push(delta / 100);  // делим на 100 для корректного отображения процентов
          }
        }
      }
    } else {
      // Если признак строки не "Объем" и не "Дельта", просто пропускаем
      continue;
    }

    // Записываем всю строку за один вызов
    sheet.getRange(row, 3, 1, datesApiFormat.length).setValues([rowValues]);

    // Устанавливаем формат ячеек в зависимости от типа строки
    if (rowType.toString().toLowerCase() === 'дельта') {
      // Формат процентов
      sheet.getRange(row, 3, 1, datesApiFormat.length).setNumberFormat('0.00%');
    } else if (rowType.toString().toLowerCase() === 'объем') {
      // Формат Russian Ruble с локалью
      sheet.getRange(row, 3, 1, datesApiFormat.length).setNumberFormat('#,##0.00[$ ₽] ');
    }
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
