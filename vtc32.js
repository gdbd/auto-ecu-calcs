const xlsx = require('xlsx');

// --- Настройки ---
const FILE_NAME = './vtc-calc-2.ods';           // Исходный файл
const RESULT_SHEET_NAME = 'Optimal_VTC'; // Имя листа для результата
const MIN_VE_DIFF = 1.0;                 // Порог чувствительности

// Размерность таблиц (32x24)
const ROWS = 24;
const COLS = 32;

// Координаты для VTC (B21:AG44)
const START_COL = 1;         // Столбец B (индекс 1)
const VTC_START_ROW = 20;    // Строка 21 (индекс 20)

// Координаты для VE (B49:AG72)
const VE_START_ROW = 48;     // Строка 49 (индекс 48)

// Координаты Осей именно для таблицы VTC (B21)
const X_AXIS_ROW = 19;       // Строка 20 (индекс 19) — над таблицей VTC
const Y_AXIS_COL = 0;        // Столбец A (индекс 0) — слева от таблицы VTC

function processOds() {
    console.log(`Чтение файла: ${FILE_NAME}...`);
    const workbook = xlsx.readFile(FILE_NAME);

    const maxVe = Array.from({ length: ROWS }, () => Array(COLS).fill(-Infinity));
    const finalVtc = Array.from({ length: ROWS }, () => Array(COLS).fill(0));

    let xAxis = [];
    let yAxis = [];
    let processedSheetsCount = 0;

    for (const sheetName of workbook.SheetNames) {
        if (sheetName === RESULT_SHEET_NAME) continue;

        console.log(`Анализ листа: ${sheetName}`);
        processedSheetsCount++;
        const sheet = workbook.Sheets[sheetName];

        // Считываем оси именно для верхней таблицы (VTC)
        if (xAxis.length === 0) {
            for (let c = 0; c < COLS; c++) {
                const cell = sheet[xlsx.utils.encode_cell({ r: X_AXIS_ROW, c: START_COL + c })];
                xAxis.push(cell ? cell.v : '');
            }
            for (let r = 0; r < ROWS; r++) {
                const cell = sheet[xlsx.utils.encode_cell({ r: VTC_START_ROW + r, c: Y_AXIS_COL })];
                yAxis.push(cell ? cell.v : '');
            }
        }

        // Сравнение VE и выбор лучшего VTC
        for (let r = 0; r < ROWS; r++) {
            for (let c = 0; c < COLS; c++) {
                const vtcAddr = xlsx.utils.encode_cell({ r: VTC_START_ROW + r, c: START_COL + c });
                const veAddr = xlsx.utils.encode_cell({ r: VE_START_ROW + r, c: START_COL + c });

                const vtcVal = sheet[vtcAddr] ? parseFloat(sheet[vtcAddr].v) : 0;
                const veVal = sheet[veAddr] ? parseFloat(sheet[veAddr].v) : -Infinity;

                if (!isNaN(veVal) && (veVal - maxVe[r][c]) >= MIN_VE_DIFF) {
                    maxVe[r][c] = veVal;
                    finalVtc[r][c] = isNaN(vtcVal) ? 0 : vtcVal;
                }
            }
        }
    }

    if (processedSheetsCount === 0) return console.log('Нет данных для анализа');

    console.log('Сборка итогового листа...');

    // Создаем массив данных нужного размера, чтобы все ячейки встали на свои места
    // Нам нужно как минимум до 44 строки для VTC
    const outputAOA = [];
    for (let i = 0; i <= VTC_START_ROW + ROWS; i++) {
        outputAOA.push(new Array(START_COL + COLS).fill(null));
    }

    // 1. Записываем ось X (в строку 20)
    xAxis.forEach((val, i) => {
        outputAOA[X_AXIS_ROW][START_COL + i] = val;
    });

    // 2. Записываем ось Y (в столбец A) и данные VTC (в B21:AG44)
    for (let r = 0; r < ROWS; r++) {
        outputAOA[VTC_START_ROW + r][Y_AXIS_COL] = yAxis[r];
        for (let c = 0; c < COLS; c++) {
            outputAOA[VTC_START_ROW + r][START_COL + c] = finalVtc[r][c];
        }
    }

    const newSheet = xlsx.utils.aoa_to_sheet(outputAOA);

    // Обновляем книгу
    workbook.Sheets[RESULT_SHEET_NAME] = newSheet;
    if (!workbook.SheetNames.includes(RESULT_SHEET_NAME)) {
        workbook.SheetNames.push(RESULT_SHEET_NAME);
    }

    xlsx.writeFile(workbook, FILE_NAME);
    console.log(`Готово! Результат в "${RESULT_SHEET_NAME}": оси в строке 20 и колонке A.`);
}

processOds();