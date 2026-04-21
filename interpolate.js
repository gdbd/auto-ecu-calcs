const xlsx = require('xlsx');

/*
 расширяем задачу: теперь в той же структуре файла есть VE размерности 32*24 по адресам B42:AG65. подписи осей : по X B41:AG41 по Y A42:A65

нужно сделать интерполяцию в VE размером 16*16 (по прежним адресам B22:Q37)

оси X B21:Q21 Y A22:A37 . нужно учитывать значения координат осей.

файл сохранить в такой же новый

далее ничего не делать, прежний алгоритм будет отдельным файлом как раньше запускаться
 */

// Настройки
const INPUT_FILE = './vtc-calc.ods';           // Исходный файл с картами
const OUTPUT_FILE = './interpolated.ods';   // Файл с готовым результатом

// --- Настройки большой таблицы (исходная 32x24) ---
const LARGE_COLS = 32;
const LARGE_ROWS = 24;
const LARGE_X_ROW = 40; // Строка 41 (индекс 40), столбцы B:AG (1..32)
const LARGE_Y_COL = 0;  // Столбец A (индекс 0), строки 42:65 (41..64)
const LARGE_DATA_START_COL = 1;
const LARGE_DATA_START_ROW = 41;

// --- Настройки маленькой таблицы (целевая 16x16) ---
const SMALL_COLS = 16;
const SMALL_ROWS = 16;
const SMALL_X_ROW = 20; // Строка 21 (индекс 20), столбцы B:Q (1..16)
const SMALL_Y_COL = 0;  // Столбец A (индекс 0), строки 22:37 (21..36)
const SMALL_DATA_START_COL = 1;
const SMALL_DATA_START_ROW = 21;

/**
 * Читает одномерный массив (ось) из листа
 */
function readAxis(sheet, startCol, startRow, length, isHorizontal) {
    const axis = [];
    for (let i = 0; i < length; i++) {
        const c = isHorizontal ? startCol + i : startCol;
        const r = isHorizontal ? startRow : startRow + i;
        const cell = sheet[xlsx.utils.encode_cell({ c, r })];
        axis.push(cell ? parseFloat(cell.v) || 0 : 0);
    }
    return axis;
}

/**
 * Читает двумерный массив (таблицу) из листа
 */
function readGrid(sheet, startCol, startRow, cols, rows) {
    const grid = [];
    for (let r = 0; r < rows; r++) {
        const rowData = [];
        for (let c = 0; c < cols; c++) {
            const cell = sheet[xlsx.utils.encode_cell({ c: startCol + c, r: startRow + r })];
            rowData.push(cell ? parseFloat(cell.v) || 0 : 0);
        }
        grid.push(rowData);
    }
    return grid;
}

/**
 * Ищет два ближайших индекса в массиве оси для искомого значения.
 * Поддерживает как возрастающие, так и убывающие оси.
 */
function getBoundingIndices(val, arr) {
    if (arr.length < 2) return [0, 0];
    const ascending = arr[arr.length - 1] > arr[0];

    // Ограничиваем (clamping) по краям, если выходим за пределы
    if (ascending) {
        if (val <= arr[0]) return [0, 0];
        if (val >= arr[arr.length - 1]) return [arr.length - 1, arr.length - 1];
    } else {
        if (val >= arr[0]) return [0, 0];
        if (val <= arr[arr.length - 1]) return [arr.length - 1, arr.length - 1];
    }

    // Ищем отрезок, в который попадает значение
    for (let i = 0; i < arr.length - 1; i++) {
        const a = arr[i], b = arr[i + 1];
        if ((ascending && val >= a && val <= b) || (!ascending && val <= a && val >= b)) {
            return [i, i + 1];
        }
    }
    return [0, 0]; // Резервный возврат
}

/**
 * Билинейная интерполяция
 */
function interpolateBilinear(targetX, targetY, xAxis, yAxis, grid) {
    const [ix1, ix2] = getBoundingIndices(targetX, xAxis);
    const [iy1, iy2] = getBoundingIndices(targetY, yAxis);

    const x1 = xAxis[ix1], x2 = xAxis[ix2];
    const y1 = yAxis[iy1], y2 = yAxis[iy2];

    const q11 = grid[iy1][ix1]; // Нижний левый (условно)
    const q21 = grid[iy1][ix2]; // Нижний правый
    const q12 = grid[iy2][ix1]; // Верхний левый
    const q22 = grid[iy2][ix2]; // Верхний правый

    // Если точное совпадение по обеим осям (попали ровно в узел)
    if (ix1 === ix2 && iy1 === iy2) return q11;

    // Линейная интерполяция, если совпала только X
    if (ix1 === ix2) {
        return q11 + (targetY - y1) * (q12 - q11) / (y2 - y1);
    }

    // Линейная интерполяция, если совпала только Y
    if (iy1 === iy2) {
        return q11 + (targetX - x1) * (q21 - q11) / (x2 - x1);
    }

    // Полноценная билинейная интерполяция (2D)
    const r1 = q11 + (targetX - x1) * (q21 - q11) / (x2 - x1);
    const r2 = q12 + (targetX - x1) * (q22 - q12) / (x2 - x1);
    const finalValue = r1 + (targetY - y1) * (r2 - r1) / (y2 - y1);

    return finalValue;
}

function processInterpolation() {
    console.log(`Открываем файл: ${INPUT_FILE}...`);
    const workbook = xlsx.readFile(INPUT_FILE);

    for (const sheetName of workbook.SheetNames) {
        console.log(`Интерполяция таблицы на листе: ${sheetName}`);
        const sheet = workbook.Sheets[sheetName];

        // 1. Считываем большие данные
        const largeX = readAxis(sheet, LARGE_DATA_START_COL, LARGE_X_ROW, LARGE_COLS, true);
        const largeY = readAxis(sheet, LARGE_Y_COL, LARGE_DATA_START_ROW, LARGE_ROWS, false);
        const largeGrid = readGrid(sheet, LARGE_DATA_START_COL, LARGE_DATA_START_ROW, LARGE_COLS, LARGE_ROWS);

        // 2. Считываем оси целевой маленькой таблицы
        const smallX = readAxis(sheet, SMALL_DATA_START_COL, SMALL_X_ROW, SMALL_COLS, true);
        const smallY = readAxis(sheet, SMALL_Y_COL, SMALL_DATA_START_ROW, SMALL_ROWS, false);

        // 3. Проходим по маленькой таблице 16x16 и интерполируем значения
        for (let r = 0; r < SMALL_ROWS; r++) {
            const targetY = smallY[r];

            for (let c = 0; c < SMALL_COLS; c++) {
                const targetX = smallX[c];

                // Рассчитываем значение
                let interpolatedVal = interpolateBilinear(targetX, targetY, largeX, largeY, largeGrid);

                // Округляем до 2 знаков после запятой для красоты (по желанию можно убрать)
                interpolatedVal = Math.round(interpolatedVal * 100) / 100;

                // Записываем результат обратно в лист (в диапазон B22:Q37)
                const cellAddress = xlsx.utils.encode_cell({
                    c: SMALL_DATA_START_COL + c,
                    r: SMALL_DATA_START_ROW + r
                });

                sheet[cellAddress] = { t: 'n', v: interpolatedVal };
            }
        }
    }

    console.log(`Сохранение результата в новый файл: ${OUTPUT_FILE}...`);
    xlsx.writeFile(workbook, OUTPUT_FILE);
    console.log('Успешно завершено!');
}

processInterpolation();