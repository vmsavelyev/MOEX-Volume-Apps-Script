//Собираем объем объем по одной бумаге - Сбер

function addSberVolumeForYesterday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Объемы');
  if (!sheet) {
    Logger.log("Лист 'Объем' не найден");
    return;
  }

  // Вчерашняя дата
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);

  const yyyy = yesterday.getFullYear();
  const mm = ('0' + (yesterday.getMonth() + 1)).slice(-2);
  const dd = ('0' + yesterday.getDate()).slice(-2);
  const dateStrApi = `${yyyy}-${mm}-${dd}`;
  const dateStrTable = Utilities.formatDate(yesterday, "GMT+3", "dd.MM.yyyy");

  // Проверяем, есть ли уже столбец с этой датой
  const lastColumn = sheet.getLastColumn();
  for (let col = 2; col <= lastColumn; col++) {
    const header = sheet.getRange(1, col).getValue();
    if (header === dateStrTable) {
      Logger.log("Данные за эту дату уже добавлены");
      return;
    }
  }

  const url = `https://iss.moex.com/iss/history/engines/stock/markets/shares/boards/TQBR/securities/SBER.json?from=${dateStrApi}&till=${dateStrApi}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const columns = json.history.columns;
    const data = json.history.data;

    if (!data || data.length === 0) {
      Logger.log(`Данные за ${dateStrApi} отсутствуют`);
      return;
    }

    const dateIndex = columns.indexOf('TRADEDATE');
    const volumeIndex = columns.indexOf('VALUE');

    if (dateIndex === -1 || volumeIndex === -1) {
      Logger.log("Не найдены колонки TRADEDATE или VALUE");
      return;
    }

    // Обычно возвращается одна запись за день
    const record = data[0];
    const volume = record[volumeIndex];

    // Записываем дату в заголовок нового столбца
    sheet.getRange(1, lastColumn + 1).setValue(dateStrTable);

    // Записываем объём во вторую строку под датой
    sheet.getRange(2, lastColumn + 1).setValue(volume);

    Logger.log(`Добавлен объём за ${dateStrTable}: ${volume}`);

