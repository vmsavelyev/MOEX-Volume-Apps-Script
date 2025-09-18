1) Добавил изменение цены в % в рамках дня. См строку Изменение. Данные берутся напрямую с API MOEX-а, никаких дополнительных вычислений не производится. Визуально - вроде бы бьется цена. 

2) В скрипте есть правило которое не позволяло заполнять данные по текущему дню, если он еще не закончен. Так вот, окончанием дня считалось 20:00 по МсК. Теперь будет 23:55.


function fillMissingValuesBatchParallel() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Объемы');
  if (!sheet) {
    Logger.log("Лист 'Объемы' не найден");
    return;
  }

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  if (lastColumn < 4 || lastRow < 2) { // теперь минимум 4 столбца
    Logger.log("Недостаточно данных");
    return;
  }

  // Даты из первой строки, начиная с 4-го столбца (т.к. столбец 3 - Средняя)
  const datesTableFormat = sheet.getRange(1, 4, 1, lastColumn - 3).getValues()[0];

  // Тикеры из первого столбца, начиная со второй строки
  const tickersRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const tickers = tickersRange.getValues().flat();

  // Признаки строк (Объем, Дельта, Изменение) из столбца 2, начиная со второй строки
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

  // Функция проверки, завершился ли торговый день для данной даты
  function isDateFinished(dateStr) {
    // dateStr в формате 'YYYY-MM-DD'
    const now = new Date();

    // Московское время (UTC+3)
    const nowMSK = new Date(now.toLocaleString("en-US", {timeZone: "Europe/Moscow"}));

    const dateParts = dateStr.split('-');
    if (dateParts.length !== 3) return false;

    const year = Number(dateParts[0]);
    const month = Number(dateParts[1]) - 1; // в JS месяцы 0-11
    const day = Number(dateParts[2]);

    const date = new Date(year, month, day);

    // Если дата в будущем — возвращаем false
    if (date > nowMSK) return false;

    // Если дата — сегодня, проверяем время (например, считаем день завершённым после 20:00 МСК)
    if (date.toDateString() === nowMSK.toDateString()) {
    if (nowMSK.getHours() < 23 || (nowMSK.getHours() === 23 && nowMSK.getMinutes() < 55)) {
    return false; // ещё не завершился торговый день
      }
    }

    // Иначе — день завершён
    return true;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  let currentIndex = Number(scriptProperties.getProperty('currentIndex')) || 0;

  const batchSize = 100;
  const endIndex = Math.min(currentIndex + batchSize, tickers.length);

  // Создаём словарь тикер -> индекс строки с объемом
  const volumeTickerToRow = {};
  for (let i = 0; i < tickers.length; i++) {
    if (rowTypes[i].toString().toLowerCase() === 'объем') {
      volumeTickerToRow[tickers[i]] = i + 2; // строка в таблице, начиная с 2-й
    }
  }

  // Получаем список выходных и праздничных дней из API isdayoff.ru для текущего года
  const currentYear = new Date().getFullYear();
  const russianHolidaysAndWeekends = getRussianWorkCalendar(currentYear);

  // Функция проверки, является ли дата выходным или праздником
  const isHolidayOrWeekend = (dateStr) => russianHolidaysAndWeekends.includes(dateStr);

  // Формируем массив запросов для fetchAll
  const requests = [];
  for (let i = currentIndex; i < endIndex; i++) {
    const ticker = tickers[i];
    let url;
    if (ticker.toUpperCase() === 'LQDT') {
      url = `https://iss.moex.com/iss/history/engines/stock/markets/shares/securities/${ticker}.json?from=${fromDate}&till=${tillDate}`;
    } else {
      url = `https://iss.moex.com/iss/history/engines/stock/markets/shares/boards/TQBR/securities/${ticker}.json?from=${fromDate}&till=${tillDate}`;
    }
    requests.push({ url: url, muteHttpExceptions: true });
  }

  // Выполняем параллельные запросы
  const responses = UrlFetchApp.fetchAll(requests);

  for (let i = 0; i < responses.length; i++) {
    const response = responses[i];
    const tickerIndex = currentIndex + i;
    const ticker = tickers[tickerIndex];
    const rowType = rowTypes[tickerIndex].toString().toLowerCase();

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

    const row = tickerIndex + 2; // строка в таблице для текущего тикера

    // Получаем текущие значения всей строки (по всем датам)
    // Данные начинаются с 4-го столбца (т.к. 3-й - Средняя)
    const currentRowValues = sheet.getRange(row, 4, 1, datesApiFormat.length).getValues()[0];

    let rowValues;

    if (rowType === 'объем') {
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

      // Для строк "Объем" — записываем значения VALUE из API, не затирая заполненные ячейки
      // и заполняем пропуски значением последнего рабочего дня
      let lastKnownValue = null;
      rowValues = datesApiFormat.map((dateKey, idx) => {
        if (!isDateFinished(dateKey)) {
          // День не завершён — оставляем ячейку как есть или пустой
          return currentRowValues[idx] !== '' && currentRowValues[idx] !== null ? currentRowValues[idx] : '';
        }
        if (currentRowValues[idx] !== '' && currentRowValues[idx] !== null) {
          lastKnownValue = currentRowValues[idx];
          return currentRowValues[idx];
        }
        const val = valuesByDate[dateKey];
        if (val !== undefined && val !== null) {
          lastKnownValue = val;
          return val;
        }
        // Если данных нет, возвращаем последнее известное значение (для выходных)
        return lastKnownValue !== null ? lastKnownValue : '';
      });
    } else if (rowType === 'дельта') {
      const volumeRow = volumeTickerToRow[ticker];
      if (!volumeRow) {
        Logger.log(`Не найдена строка Объем для тикера ${ticker}, пропускаем дельту`);
        continue;
      }
      const volumeValues = sheet.getRange(volumeRow, 4, 1, datesApiFormat.length).getValues()[0];
      rowValues = [];

      for (let idx = 0; idx < datesApiFormat.length; idx++) {
        if (!isDateFinished(datesApiFormat[idx])) {
          // День не завершён — оставляем пустым или текущим значением
          rowValues.push(currentRowValues[idx] !== '' && currentRowValues[idx] !== null ? currentRowValues[idx] : '');
          continue;
        }
        if (idx === 0) {
          rowValues.push('');
          continue;
        }
        const prevVal = parseFloat(volumeValues[idx - 1]);
        const currVal = parseFloat(volumeValues[idx]);

        if (isNaN(prevVal) || prevVal === 0 || isNaN(currVal)) {
          rowValues.push('');
        } else {
          const delta = ((currVal - prevVal) / prevVal);
          rowValues.push(delta);
        }
      }
    } else if (rowType === 'изменение') {
      const trendIndex = columns.indexOf('TRENDCLSPR');
      const dateIndex = columns.indexOf('TRADEDATE');

      if (trendIndex === -1 || dateIndex === -1) {
        Logger.log(`Не найден столбец TRENDCLSPR или TRADEDATE для тикера ${ticker}`);
        continue;
      }

      // Создаём словарь дата -> trend
      const trendByDate = {};
      data.forEach(row => {
        trendByDate[row[dateIndex]] = row[trendIndex];
      });

      rowValues = datesApiFormat.map((dateKey, idx) => {
        if (!isDateFinished(dateKey)) {
          // День не завершён — оставляем ячейку как есть или пустой
          return currentRowValues[idx] !== '' && currentRowValues[idx] !== null ? currentRowValues[idx] : '';
        }
        const trendVal = trendByDate[dateKey];
        if (trendVal === undefined || trendVal === null || isNaN(trendVal)) {
          return '';
        }
        return trendVal / 100; // предположим, что TRENDCLSPR в процентах, делим на 100 для формата
      });
    } else {
      // Если признак строки не подходит, пропускаем
      continue;
    }

    // Записываем всю строку за один вызов
    sheet.getRange(row, 4, 1, datesApiFormat.length).setValues([rowValues]);

    // Рассчитываем среднее значение по строке, игнорируя пустые и null
    const numericValues = rowValues.filter(v => typeof v === 'number' && !isNaN(v));
    const avg = numericValues.length > 0 ? numericValues.reduce((a, b) => a + b, 0) / numericValues.length : '';

    // Записываем среднее значение в столбец 3 ("Средняя")
    sheet.getRange(row, 3).setValue(avg);

    // Устанавливаем формат ячеек в зависимости от типа строки
    if (rowType === 'дельта' || rowType === 'изменение') {
      sheet.getRange(row, 4, 1, datesApiFormat.length).setNumberFormat('0.00%');
      sheet.getRange(row, 3).setNumberFormat('0.00%'); // форматируем среднее как процент
    } else if (rowType === 'объем') {
      sheet.getRange(row, 4, 1, datesApiFormat.length).setNumberFormat('#,##0.00[$ ₽]');
      sheet.getRange(row, 3).setNumberFormat('#,##0.00[$ ₽]'); // форматируем среднее как объем
    }
  }

  // Окрашиваем столбцы с выходными и праздничными днями согласно производственному календарю
  const weekendColor = '#fff2cc';
  const dataStartRow = 2;
  const dataEndRow = lastRow;

  datesApiFormat.forEach((dateStr, idx) => {
    if (!dateStr) return;
    if (isHolidayOrWeekend(dateStr)) {
      sheet.getRange(dataStartRow, idx + 4, dataEndRow - 1, 1).setBackground(weekendColor);
    } else {
      // Очистить фон, если нужно
      // sheet.getRange(dataStartRow, idx + 4, dataEndRow - 1, 1).setBackground(null);
    }
  });

  // Обновляем индекс и создаём триггер для следующей пачки или на следующий день в 4:00 МСК
  if (endIndex >= tickers.length) {
    Logger.log('Все тикеры обработаны.');
    scriptProperties.deleteProperty('currentIndex');

    // Удаляем старые триггеры
    deleteTimeDrivenTriggers();

    // Создаём триггер на следующий день в 4:00 утра по московскому времени
    const now = new Date();

    // Получаем текущую дату и время в МСК
    const nowMSK = new Date(now.toLocaleString("en-US", {timeZone: "Europe/Moscow"}));

    // Создаём дату следующего дня в 4:00 МСК
    const nextDayMSK = new Date(
      nowMSK.getFullYear(),
      nowMSK.getMonth(),
      nowMSK.getDate() + 1,
      4, 0, 0, 0
    );

    // Переводим nextDayMSK в локальное время сервера (по часовому поясу скрипта)
    // Для этого берём разницу между nowMSK и now и корректируем дату
    const offset = now.getTime() - nowMSK.getTime();
    const triggerTime = new Date(nextDayMSK.getTime() + offset);

    ScriptApp.newTrigger('fillMissingValuesBatchParallel')
      .timeBased()
      .at(triggerTime)
      .create();

    Logger.log(`Создан триггер на следующий день в 4:00 утра МСК: ${triggerTime}`);
  } else {
    // Текущая логика с запуском через 5 секунд для следующей пачки
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

// Функция для получения официальных выходных и праздничных дней России по производственному календарю с isdayoff.ru
function getRussianWorkCalendar(year) {
  const weekendsAndHolidays = [];
  for (let month = 1; month <= 12; month++) {
    const url = `https://isdayoff.ru/api/getdata?year=${year}&month=${month}&cc=ru`;
    try {
      const response = UrlFetchApp.fetch(url);
      if (response.getResponseCode() !== 200) {
        Logger.log(`Ошибка при получении данных за ${year}-${month}`);
        continue;
      }
      const data = response.getContentText(); // строка из 0 и 1 по дням месяца
      for (let day = 1; day <= data.length; day++) {
        if (data.charAt(day - 1) === '1') {
          const mm = month < 10 ? '0' + month : month;
          const dd = day < 10 ? '0' + day : day;
          weekendsAndHolidays.push(`${year}-${mm}-${dd}`);
        }
      }
    } catch (e) {
      Logger.log(`Ошибка запроса isdayoff.ru за ${year}-${month}: ${e}`);
    }
  }
  return weekendsAndHolidays;
}
