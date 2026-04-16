const xlsx = require("xlsx");

/**
 задача: распарсить ods эксель файл, в нем есть несколько страниц.
 на странице каждой есть:
 - таблица VTC B3:Q18
 - таблица VE  B22:Q37
 ячейки каждой пары таблицах, например D8 и D27 соответствуют друг другу
 нужно из всех страниц собрать одну новую таблицу VTC, сравнив VE на каждой из страниц и выбрать такое где значение макисмально
 результат сохранить в тот же файл новой страницей и так чтобы она не учитывалась при анализе если повторно запустить
 */

// Настройки
const FILE_NAME = "./interpolated.ods"; // Имя файла (он же источник и результат)
const RESULT_SHEET_NAME = "Optimal_VTC"; // Имя листа, куда запишем результат

// МИНИМАЛЬНАЯ РАЗНИЦА: насколько VE должно быть больше текущего максимума, чтобы мы взяли новый VTC
// Например, 1.0 означает, что если старое VE=80.0, новое должно быть >= 81.0
const MIN_VE_DIFF = 1.0;

// Размерность таблиц (16x16)
const ROWS = 16;
const COLS = 16;

// Смещения и координаты таблиц
const START_COL = 1;      // Таблицы начинаются с колонки B (индекс 1)
const VTC_START_ROW = 2;  // VTC начинается со строки 3 (индекс 2)
const VE_START_ROW = 21;  // VE начинается со строки 22 (индекс 21)

// Координаты осей для VTC (чтобы перенести их в результат)
const X_AXIS_ROW = 1;     // Ось X лежит в строке 2 (индекс 1), столбцы B:Q
const Y_AXIS_COL = 0;     // Ось Y лежит в столбце A (индекс 0), строки 3:18

function processOds() {
  console.log(`Чтение файла: ${FILE_NAME}...`);
  const workbook = xlsx.readFile(FILE_NAME);

  const maxVe = Array.from({ length: ROWS }, () => Array(COLS).fill(-Infinity));
  const finalVtc = Array.from({ length: ROWS }, () => Array(COLS).fill(0));

  let processedSheetsCount = 0;

  // Массивы для сохранения осей X и Y
  let xAxis = [];
  let yAxis = [];

  // Проходим по всем страницам в файле
  for (const sheetName of workbook.SheetNames) {
    if (sheetName === RESULT_SHEET_NAME) {
      console.log(`Пропуск итогового листа: ${sheetName}`);
      continue;
    }

    console.log(`Анализ листа: ${sheetName}`);
    processedSheetsCount++;
    const sheet = workbook.Sheets[sheetName];

    // Читаем оси с первой попавшейся страницы (предполагаем, что они везде одинаковые)
    if (xAxis.length === 0) {
      for (let c = 0; c < COLS; c++) {
        const cell = sheet[xlsx.utils.encode_cell({ r: X_AXIS_ROW, c: START_COL + c })];
        xAxis.push(cell ? parseFloat(cell.v) || 0 : 0);
      }
      for (let r = 0; r < ROWS; r++) {
        const cell = sheet[xlsx.utils.encode_cell({ r: VTC_START_ROW + r, c: Y_AXIS_COL })];
        yAxis.push(cell ? parseFloat(cell.v) || 0 : 0);
      }
    }

    for (let r = 0; r < ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        const vtcAddress = xlsx.utils.encode_cell({ r: VTC_START_ROW + r, c: START_COL + c });
        const veAddress = xlsx.utils.encode_cell({ r: VE_START_ROW + r, c: START_COL + c });

        const vtcCell = sheet[vtcAddress];
        const veCell = sheet[veAddress];

        const vtcValue = vtcCell ? parseFloat(vtcCell.v) : 0;
        const veValue = veCell ? parseFloat(veCell.v) : -Infinity;

        // Проверяем: текущее VE валидно и разница с предыдущим максимумом >= заданному порогу MIN_VE_DIFF
        // (При первой записи maxVe равен -Infinity, поэтому первое значение запишется 100%)
        if (!isNaN(veValue) && (veValue - maxVe[r][c]) >= MIN_VE_DIFF) {
          maxVe[r][c] = veValue;
          finalVtc[r][c] = isNaN(vtcValue) ? 0 : vtcValue;
        }
      }
    }
  }

  if (processedSheetsCount === 0) {
    console.log('Нет страниц для анализа!');
    return;
  }

  console.log('Генерация итоговой таблицы с осями...');

  // Формируем данные для новой страницы
  const outputData = [];

  // Строка 1: пустая
  outputData.push([]);

  // Строка 2: пустая ячейка A2, затем значения оси X (B2:Q2)
  outputData.push([null, ...xAxis]);

  // Строки 3-18: значение оси Y, затем значения VTC (A3:Q18)
  for (let r = 0; r < ROWS; r++) {
    const rowData = [yAxis[r], ...finalVtc[r]];
    outputData.push(rowData);
  }

  const newSheet = xlsx.utils.aoa_to_sheet(outputData);

  // Записываем или обновляем лист
  workbook.Sheets[RESULT_SHEET_NAME] = newSheet;

  if (!workbook.SheetNames.includes(RESULT_SHEET_NAME)) {
    workbook.SheetNames.push(RESULT_SHEET_NAME);
  }

  console.log(`Сохранение данных в ${FILE_NAME}...`);
  xlsx.writeFile(workbook, FILE_NAME);
  console.log(`Готово! Результат с осями и учетом порога (${MIN_VE_DIFF}) сохранен на лист "${RESULT_SHEET_NAME}"`);
}

processOds();