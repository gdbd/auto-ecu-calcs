const xlsx = require('xlsx');

/*
сделай еще один скрипт - экстраполяция ячеек на вкладке current из B3:Q18 (оси B2:Q2 A3:A18) в B21:AG44 (оси B48:AG48 A21:A44)
*/
// --- Конфигурация ---
const FILE_NAME = './vtc-calc-2.ods';
const SHEET_TARGET = 'current';
const OUTPUT_FILE = 'extrapolated.ods';

// Исходная сетка (16x16)
const SRC = {
    xRow: 1,             // B2:Q2 (индекс строки 1)
    yCol: 0,             // A3:A18 (индекс колонки 0)
    dataStartRow: 2,     // B3 (индекс строки 2)
    dataStartCol: 1,     // B3 (индекс колонки 1)
    sizeX: 16,
    sizeY: 16
};

// Целевая сетка (32x24)
const DST = {
    xRow: 47,            // B48:AG48 (индекс строки 47)
    yCol: 0,             // A21:A44 (индекс колонки 0)
    dataStartRow: 20,    // B21 (индекс строки 20)
    dataStartCol: 1,     // B21 (индекс колонки 1)
    sizeX: 32,
    sizeY: 24
};

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

function getBoundingIndices(val, arr) {
    const len = arr.length;
    if (len < 2) return [0, 0];
    const ascending = arr[len - 1] > arr[0];

    // Экстраполяция (Clamping): если значение за пределами оси
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

function interpolate(targetX, targetY, xAxis, yAxis, grid) {
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

function run() {
    console.log(`Загрузка файла ${FILE_NAME}...`);
    const workbook = xlsx.readFile(FILE_NAME);
    const sheet = workbook.Sheets[SHEET_TARGET];

    if (!sheet) {
        console.error(`Лист "${SHEET_TARGET}" не найден!`);
        return;
    }

    // Читаем исходные данные (16x16)
    const srcX = readAxis(sheet, SRC.dataStartCol, SRC.xRow, SRC.sizeX, true);
    const srcY = readAxis(sheet, SRC.yCol, SRC.dataStartRow, SRC.sizeY, false);
    const srcGrid = readGrid(sheet, SRC.dataStartCol, SRC.dataStartRow, SRC.sizeX, SRC.sizeY);

    // Читаем целевые оси (32x24)
    const dstX = readAxis(sheet, DST.dataStartCol, DST.xRow, DST.sizeX, true);
    const dstY = readAxis(sheet, DST.yCol, DST.dataStartRow, DST.sizeY, false);

    console.log("Расчет экстраполяции...");

    for (let r = 0; r < DST.sizeY; r++) {
        for (let c = 0; c < DST.sizeX; c++) {
            const val = interpolate(dstX[c], dstY[r], srcX, srcY, srcGrid);

            const cellAddr = xlsx.utils.encode_cell({
                c: DST.dataStartCol + c,
                r: DST.dataStartRow + r
            });

            sheet[cellAddr] = { t: 'n', v: Math.round(val * 100) / 100 };
        }
    }

    console.log(`Сохранение в ${OUTPUT_FILE}...`);
    xlsx.writeFile(workbook, OUTPUT_FILE);
    console.log("Готово.");
}

run();