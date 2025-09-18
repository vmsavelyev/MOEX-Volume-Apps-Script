//Скрипт который заполняет данные по датам в первой строке начиная со второго столба + который собирает данные по всем пикетам. В первом столбец

function fillMissingValuesOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Объемы');
  if (!sheet) {
    Logger.log("Лист 'Объемы' не найден");
    return;
  }

  // Получаем даты из первой строки, начиная со второго столбца
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 2) {
    Logger.log("В первой строке нет дат");
    return;
  }
  const datesTableFormat = sheet.getRange(1, 2, 1, lastColumn - 1).getValues()[0]; // массив дат в формате dd.MM.yyyy

  // Преобразуем даты из dd.MM.yyyy в yyyy-MM-dd для API
  function convertDateFormat(dateValue) {
  if (!dateValue) return null; // пустое значение

  // Если это объект Date, преобразуем в строку dd.MM.yyyy
  if (Object.prototype.toString.call(dateValue) === '[object Date]') {
    if (isNaN(dateValue.getTime())) return null; // некорректная дата
    const dd = ('0' + dateValue.getDate()).slice(-2);
    const mm = ('0' + (dateValue.getMonth() + 1)).slice(-2);
    const yyyy = dateValue.getFullYear();
    return `${yyyy}-${mm}-${dd}`; // формат API
  }

  // Если строка, ожидаем формат dd.MM.yyyy
  if (typeof dateValue === 'string') {
    const parts = dateValue.split('.');
    if (parts.length !== 3) return null;
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }

  // Если другое — возвращаем null
  return null;
}
  const datesApiFormat = datesTableFormat.map(d => convertDateFormat(d));

  // Получаем тикеры из первого столбца, начиная со второй строки
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Нет тикеров в первом столбце");
    return;
  }
  const tickersRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const tickers = tickersRange.getValues().flat().filter(t => t && t.toString().trim() !== '');

  // Функция для получения VALUE за дату и тикер
  function fetchValue(ticker, dateStr) {
    const url = `https://iss.moex.com/iss/history/engines/stock/markets/shares/boards/TQBR/securities/${ticker}.json?from=${dateStr}&till=${dateStr}`;
    try {
      const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      if (response.getResponseCode() !== 200) {
        Logger.log(`Ошибка HTTP ${response.getResponseCode()} для тикера ${ticker} и даты ${dateStr}`);
        return null;
      }
      const json = JSON.parse(response.getContentText());
      const columns = json.history.columns;
      const data = json.history.data;
      if (!data || data.length === 0) {
        return null;
      }
      const valueIndex = columns.indexOf('VALUE');
      if (valueIndex === -1) {
        Logger.log(`Не найден столбец VALUE для тикера ${ticker}`);
        return null;
      }
      return data[0][valueIndex];
    } catch (e) {
      Logger.log(`Ошибка при запросе данных для тикера ${ticker} за ${dateStr}: ${e.message}`);
      return null;
    }
  }

  // Проходим по каждому тикеру и каждой дате, заполняем только если ячейка пустая
  tickers.forEach((ticker, rowIndex) => {
    const row = rowIndex + 2; // строки начинаются со второй

    for (let i = 0; i < datesApiFormat.length; i++) {
      const col = i + 2; // даты начинаются со второго столбца

      // Проверяем, пустая ли ячейка
      const cellValue = sheet.getRange(row, col).getValue();
      if (cellValue !== '' && cellValue !== null) {
        // Если есть данные, пропускаем
        continue;
      }

      // Если пусто — получаем данные и записываем
      const dateApi = datesApiFormat[i];
      if (!dateApi) {
        Logger.log(`Некорректный формат даты: ${datesTableFormat[i]}`);
        continue;
      }

      const value = fetchValue(ticker, dateApi);
      if (value !== null) {
        sheet.getRange(row, col).setValue(value);
      } else {
        sheet.getRange(row, col).setValue('');
      }

      Utilities.sleep(150); // Пауза для API
    }
  });

  Logger.log('Заполнение пропущенных значений VALUE завершено');
}