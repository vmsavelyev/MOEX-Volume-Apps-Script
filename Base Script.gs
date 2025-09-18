//Собираем объем объем по одной бумаге - Сбер

function addSberVolumeForYesterday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Объемы');
  if (!sheet) {
    Logger.log("Лист 'Объем' не найден");
    return;
  }

  // Генерируем даты 2025 года в формате dd.MM.yyyy и yyyy-MM-dd
  const datesTableFormat = [];
  const datesApiFormat = [];
  for (let d = new Date(2025, 0, 1); d <= new Date(2025, 0, 30); d.setDate(d.getDate() + 1)) {
    const dd = ('0' + d.getDate()).slice(-2);
    const mm = ('0' + (d.getMonth() + 1)).slice(-2);
    const yyyy = d.getFullYear();
    datesTableFormat.push(`${dd}.${mm}.${yyyy}`);
    datesApiFormat.push(`${yyyy}-${mm}-${dd}`);
  }

  // Проверяем первую строку на наличие дат, если нет - пишем
  const lastColumn = sheet.getLastColumn();
  const headerRange = sheet.getRange(1, 2, 1, lastColumn - 1);
  const existingDates = lastColumn > 1 ? headerRange.getValues()[0] : [];

  // Записываем даты 2025 в первую строку, начиная со второго столбца
  datesTableFormat.forEach((dateStr, idx) => {
    if (existingDates[idx] !== dateStr) {
      sheet.getRange(1, idx + 2).setValue(dateStr);
    }
  });

  // Проверяем, есть ли тикер SBER в первом столбце (начиная со второй строки)
  const lastRow = sheet.getLastRow();
  let sberRow = -1;
  if (lastRow >= 2) {
    const tickers = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    sberRow = tickers.findIndex(t => t === 'SBER');
    if (sberRow !== -1) {
      sberRow += 2; // Индекс в массиве начинается с 0, а в таблице со строки 2
    }
  }
  if (sberRow === -1) {
    // Добавляем SBER в первый столбец под последней заполненной строкой
    sberRow = lastRow >= 2 ? lastRow + 1 : 2;
    sheet.getRange(sberRow, 1).setValue('SBER');
  }

    // Функция для получения VALUE за дату
    function fetchValue(dateStr) {
      const url = `https://iss.moex.com/iss/history/engines/stock/markets/shares/boards/TQBR/securities/SBER.json?from=${dateStr}&till=${dateStr}`;
      try {
        const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
        if (response.getResponseCode() !== 200) {
          Logger.log(`Ошибка HTTP ${response.getResponseCode()} для даты ${dateStr}`);
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
          Logger.log("Не найден столбец VALUE");
          return null;
        }
        return data[0][valueIndex];
      } catch (e) {
        Logger.log(`Ошибка при запросе данных за ${dateStr}: ${e.message}`);
        return null;
      }
    }
    
  // Записываем значения VALUE по датам в строку SBER, начиная со второго столбца
  for (let i = 0; i < datesApiFormat.length; i++) {
    const dateApi = datesApiFormat[i];
    const col = i + 2; // колонки с датами начинаются со второго столбца
    const value = fetchValue(dateApi);
    if (value !== null) {
      sheet.getRange(sberRow, col).setValue(value);
    } else {
      sheet.getRange(sberRow, col).setValue('');
    }
    Utilities.sleep(150); // Пауза, чтобы не перегружать API
  }

  Logger.log('Заполнение объёмов SBER за 2025 год завершено');
}