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
const FILE_NAME = "./vtc-calc.ods"; // Имя файла (он же источник и результат)
const RESULT_SHEET_NAME = "Optimal_VTC"; // Имя листа, куда запишем результат

// Размерность таблиц (16x16)
const ROWS = 16;
const COLS = 16;
const START_COL = 1; // Колонка B (индекс 1)
const VTC_START_ROW = 2; // Строка 3 (индекс 2)
const VE_START_ROW = 21; // Строка 22 (индекс 21)

function processOds() {
  console.log(`Чтение файла: ${FILE_NAME}...`);
  const workbook = xlsx.readFile(FILE_NAME);

  const maxVe = Array.from({ length: ROWS }, () => Array(COLS).fill(-Infinity));
  const finalVtc = Array.from({ length: ROWS }, () => Array(COLS).fill(0));

  let processedSheetsCount = 0;

  // Проходим по всем страницам в файле
  for (const sheetName of workbook.SheetNames) {
    // Пропускаем страницу с результатами от предыдущего запуска
    if (sheetName === RESULT_SHEET_NAME) {
      console.log(`Пропуск итогового листа: ${sheetName}`);
      continue;
    }

    console.log(`Анализ листа: ${sheetName}`);
    processedSheetsCount++;
    const sheet = workbook.Sheets[sheetName];

    for (let r = 0; r < ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        const vtcAddress = xlsx.utils.encode_cell({
          r: VTC_START_ROW + r,
          c: START_COL + c,
        });
        const veAddress = xlsx.utils.encode_cell({
          r: VE_START_ROW + r,
          c: START_COL + c,
        });

        const vtcCell = sheet[vtcAddress];
        const veCell = sheet[veAddress];

        const vtcValue = vtcCell ? parseFloat(vtcCell.v) : 0;
        const veValue = veCell ? parseFloat(veCell.v) : -Infinity;

        if (!isNaN(veValue) && veValue > maxVe[r][c]) {
          maxVe[r][c] = veValue;
          finalVtc[r][c] = isNaN(vtcValue) ? 0 : vtcValue;
        }
      }
    }
  }

  if (processedSheetsCount === 0) {
    console.log("Нет страниц для анализа!");
    return;
  }

  console.log("Генерация итоговой таблицы...");

  const outputData = [
    [], // Строка 1
    [], // Строка 2
  ];

  for (let r = 0; r < ROWS; r++) {
    const rowData = [null, ...finalVtc[r]];
    outputData.push(rowData);
  }

  const newSheet = xlsx.utils.aoa_to_sheet(outputData);

  // Записываем или обновляем лист в текущей книге
  workbook.Sheets[RESULT_SHEET_NAME] = newSheet;

  // Если такого листа еще не было, добавляем его в список страниц
  if (!workbook.SheetNames.includes(RESULT_SHEET_NAME)) {
    workbook.SheetNames.push(RESULT_SHEET_NAME);
  }

  console.log(`Сохранение данных в ${FILE_NAME}...`);
  // Перезаписываем исходный файл
  xlsx.writeFile(workbook, FILE_NAME);
  console.log(`Готово! Результат сохранен на страницу "${RESULT_SHEET_NAME}"`);
}

processOds();
