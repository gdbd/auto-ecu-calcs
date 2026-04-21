const xlsx = require('xlsx');

// --- Настройки ---
const FILE_NAME = './vtc-calc-2.ods';
const RESULT_SHEET_NAME = 'Optimal_VTC';
const MIN_VE_DIFF = 1.0;

// --- Координаты 32x24 (Большая таблица) ---
const L_ROWS = 24;
const L_COLS = 32;
const L_START_COL = 1;      // Столбец B
const L_VTC_START_ROW = 20; // Строка 21
const L_VE_START_ROW = 48;  // Строка 49
const L_X_ROW = 19;         // Ось X: Строка 20
const L_Y_COL = 0;          // Ось Y: Столбец A

// --- Координаты 16x16 (Маленькая таблица) ---
const S_ROWS = 16;
const S_COLS = 16;
const S_START_COL = 1;      // Столбец B
const S_START_ROW = 2;      // Строка 3
const S_X_ROW = 1;          // Ось X: Строка 2
const S_Y_COL = 0;          // Ось Y: Столбец A

// ================= ФУНКЦИИ ИНТЕРПОЛЯЦИИ =================

function getBoundingIndices(val, arr) {
    const len = arr.length;
    if (len < 2) return [0, 0];
    const ascending = arr[len - 1] > arr[0];

    // Clamping (зажим), если выходим за пределы осей
    if (ascending) {
        if (val <= arr[0]) return [0, 0];
        if (val >= arr[len - 1]) return [len - 1, len - 1];
    } else {
        if (val >= arr[0]) return [0, 0];
        if (val <= arr[len - 1]) return [len - 1, len - 1];
    }

    for (let i = 0; i < len - 1; i++) {
        if ((ascending && val >= arr[i] && val <= arr[i + 1]) ||
            (!ascending && val <= arr[i] && val >= arr[i + 1])) {
            return [i, i + 1];
        }
    }
    return [0, 0];
}

function interpolateBilinear(targetX, targetY, xAxis, yAxis, grid) {
    const [ix1, ix2] = getBoundingIndices(targetX, xAxis);
    const [iy1, iy2] = getBoundingIndices(targetY, yAxis);

    const x1 = xAxis[ix1], x2 = xAxis[ix2];
    const y1 = yAxis[iy1], y2 = yAxis[iy2];
    const q11 = grid[iy1][ix1], q21 = grid[iy1][ix2];
    const q12 = grid[iy2][ix1], q22 = grid[iy2][ix2];

    if (ix1 === ix2 && iy1 === iy2) return q11;
    if (ix1 === ix2) return q11 + (targetY - y1) * (q12 - q11) / (y2 - y1);
    if (iy1 === iy2) return q11 + (targetX - x1) * (q21 - q11) / (x2 - x1);

    const r1 = q11 + (targetX - x1) * (q21 - q11) / (x2 - x1);
    const r2 = q12 + (targetX - x1) * (q22 - q12) / (x2 - x1);
    return r1 + (targetY - y1) * (r2 - r1) / (y2 - y1);
}

// ================= ОСНОВНОЙ ПРОЦЕСС =================

function processOds() {
    console.log(`Чтение файла: ${FILE_NAME}...`);
    const workbook = xlsx.readFile(FILE_NAME);

    const maxVe = Array.from({ length: L_ROWS }, () => Array(L_COLS).fill(-Infinity));
    const finalVtcLarge = Array.from({ length: L_ROWS }, () => Array(L_COLS).fill(0));

    let largeX = [], largeY = [];
    let smallX = [], smallY = [];
    let processedSheetsCount = 0;

    for (const sheetName of workbook.SheetNames) {
        if (sheetName === RESULT_SHEET_NAME) continue;

        console.log(`Анализ листа: ${sheetName}`);
        processedSheetsCount++;
        const sheet = workbook.Sheets[sheetName];

        // 1. Считываем оси для обеих размерностей с первой страницы
        if (largeX.length === 0) {
            // Оси 32x24
            for (let c = 0; c < L_COLS; c++) {
                const cell = sheet[xlsx.utils.encode_cell({ r: L_X_ROW, c: L_START_COL + c })];
                largeX.push(cell ? parseFloat(cell.v) || 0 : 0);
            }
            for (let r = 0; r < L_ROWS; r++) {
                const cell = sheet[xlsx.utils.encode_cell({ r: L_VTC_START_ROW + r, c: L_Y_COL })];
                largeY.push(cell ? parseFloat(cell.v) || 0 : 0);
            }
            // Оси 16x16
            for (let c = 0; c < S_COLS; c++) {
                const cell = sheet[xlsx.utils.encode_cell({ r: S_X_ROW, c: S_START_COL + c })];
                smallX.push(cell ? parseFloat(cell.v) || 0 : 0);
            }
            for (let r = 0; r < S_ROWS; r++) {
                const cell = sheet[xlsx.utils.encode_cell({ r: S_START_ROW + r, c: S_Y_COL })];
                smallY.push(cell ? parseFloat(cell.v) || 0 : 0);
            }
        }

        // 2. Сравнение VE и выбор лучшего VTC (матрица 32x24)
        for (let r = 0; r < L_ROWS; r++) {
            for (let c = 0; c < L_COLS; c++) {
                const vtcAddr = xlsx.utils.encode_cell({ r: L_VTC_START_ROW + r, c: L_START_COL + c });
                const veAddr = xlsx.utils.encode_cell({ r: L_VE_START_ROW + r, c: L_START_COL + c });

                const vtcVal = sheet[vtcAddr] ? parseFloat(sheet[vtcAddr].v) : 0;
                const veVal = sheet[veAddr] ? parseFloat(sheet[veAddr].v) : -Infinity;

                if (!isNaN(veVal) && (veVal - maxVe[r][c]) >= MIN_VE_DIFF) {
                    maxVe[r][c] = veVal;
                    finalVtcLarge[r][c] = isNaN(vtcVal) ? 0 : vtcVal;
                }
            }
        }
    }

    if (processedSheetsCount === 0) return console.log('Нет данных для анализа');

    // 3. Выполняем интерполяцию (сжатие из 32x24 в 16x16)
    console.log('Сжатие лучшей матрицы в формат 16x16...');
    const finalVtcSmall = Array.from({ length: S_ROWS }, () => Array(S_COLS).fill(0));

    for (let r = 0; r < S_ROWS; r++) {
        for (let c = 0; c < S_COLS; c++) {
            let val = interpolateBilinear(smallX[c], smallY[r], largeX, largeY, finalVtcLarge);
            finalVtcSmall[r][c] = Math.round(val * 100) / 100; // Округление до 2 знаков
        }
    }

    console.log('Сборка итогового листа...');

    // 4. Подготавливаем чистый массив для записи (от строки 1 до 44 минимум)
    const outputAOA = [];
    for (let i = 0; i <= L_VTC_START_ROW + L_ROWS; i++) {
        outputAOA.push(new Array(L_START_COL + L_COLS).fill(null));
    }

    // --- Запись 16x16 ---
    // Оси X и Y (B2:Q2, A3:A18)
    smallX.forEach((val, i) => outputAOA[S_X_ROW][S_START_COL + i] = val);
    smallY.forEach((val, i) => outputAOA[S_START_ROW + i][S_Y_COL] = val);
    // Данные (B3:Q18)
    for (let r = 0; r < S_ROWS; r++) {
        for (let c = 0; c < S_COLS; c++) {
            outputAOA[S_START_ROW + r][S_START_COL + c] = finalVtcSmall[r][c];
        }
    }

    // --- Запись 32x24 ---
    // Оси X и Y (B20:AG20, A21:A44)
    largeX.forEach((val, i) => outputAOA[L_X_ROW][L_START_COL + i] = val);
    largeY.forEach((val, i) => outputAOA[L_VTC_START_ROW + i][L_Y_COL] = val);
    // Данные (B21:AG44)
    for (let r = 0; r < L_ROWS; r++) {
        for (let c = 0; c < L_COLS; c++) {
            outputAOA[L_VTC_START_ROW + r][L_START_COL + c] = finalVtcLarge[r][c];
        }
    }

    const newSheet = xlsx.utils.aoa_to_sheet(outputAOA);
    workbook.Sheets[RESULT_SHEET_NAME] = newSheet;

    if (!workbook.SheetNames.includes(RESULT_SHEET_NAME)) {
        workbook.SheetNames.push(RESULT_SHEET_NAME);
    }

    xlsx.writeFile(workbook, FILE_NAME);
    console.log(`Готово! В "${RESULT_SHEET_NAME}" сохранена оптимальная карта 32x24 и её копия 16x16.`);
}

processOds();