const fs = require('node:fs');
const path = require('node:path');
const XLSX = require('xlsx');
const {
  FileValidationError,
  SUPPORTED_EXTENSIONS,
  isRowMeaningful,
  normalizeCell,
  trimTrailingEmptyCells
} = require('./common');

function ensureSupportedFile(filePath) {
  const extension = path.extname(filePath).toLowerCase();

  if (!SUPPORTED_EXTENSIONS.has(extension)) {
    throw new FileValidationError('FILE_TYPE', '文件类型错误，请重新导入');
  }
}

function readWorkbookRows(filePath, { blankrows = false } = {}) {
  ensureSupportedFile(filePath);

  if (!fs.existsSync(filePath) || fs.statSync(filePath).size === 0) {
    throw new FileValidationError('FILE_READ', '文件为空或不可读，请重新导入');
  }

  try {
    const workbook = XLSX.readFile(filePath, {
      cellDates: false,
      dense: true,
      raw: false
    });
    const firstSheetName = workbook.SheetNames[0];

    if (!firstSheetName) {
      throw new FileValidationError('FILE_READ', '文件为空或不可读，请重新导入');
    }

    const sheet = workbook.Sheets[firstSheetName];
    return XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows,
      defval: ''
    });
  } catch (error) {
    if (error instanceof FileValidationError) {
      throw error;
    }

    throw new FileValidationError('FILE_READ', '文件为空或不可读，请重新导入');
  }
}

function readRows(filePath) {
  const rows = readWorkbookRows(filePath, { blankrows: false });

  if (!Array.isArray(rows) || rows.length === 0 || !rows.some(isRowMeaningful)) {
    throw new FileValidationError('FILE_READ', '文件为空或不可读，请重新导入');
  }

  return rows;
}

function readRowsWithMetadata(filePath, expectedHeaders = []) {
  const rawRows = readWorkbookRows(filePath, { blankrows: true });
  const normalizedExpectedHeaders = Array.isArray(expectedHeaders)
    ? expectedHeaders.map((header) => normalizeCell(header)).filter((header) => header !== '')
    : [];
  const meaningfulRows = rawRows
    .map((row, index) => ({
      rowNumber: index + 1,
      cells: trimTrailingEmptyCells(Array.isArray(row) ? row : [])
    }))
    .filter((row) => isRowMeaningful(row.cells));

  if (!meaningfulRows.length) {
    throw new FileValidationError('FILE_READ', '文件为空或不可读，请重新导入');
  }

  if (!normalizedExpectedHeaders.length) {
    return {
      rows: meaningfulRows.map((row) => row.cells),
      rowNumbers: meaningfulRows.map((row) => row.rowNumber)
    };
  }

  const expectedHeaderCount = normalizedExpectedHeaders.length;
  let matchedRowIndex = -1;
  let matchedColumnIndex = -1;

  meaningfulRows.some((row, rowIndex) => {
    const maximumStartIndex = row.cells.length - expectedHeaderCount;

    for (let startIndex = 0; startIndex <= maximumStartIndex; startIndex += 1) {
      const candidateHeaders = row.cells
        .slice(startIndex, startIndex + expectedHeaderCount)
        .map((cell) => normalizeCell(cell));

      if (candidateHeaders.every((cell, index) => cell === normalizedExpectedHeaders[index])) {
        matchedRowIndex = rowIndex;
        matchedColumnIndex = startIndex;
        return true;
      }
    }

    return false;
  });

  if (matchedRowIndex < 0 || matchedColumnIndex < 0) {
    throw new FileValidationError(
      'FILE_READ',
      '当前导入文件未匹配到所选模板的表头，请确认模板或原始网银账单是否正确'
    );
  }

  const rows = [];
  const rowNumbers = [];
  const summaryLabels = ['总收入笔数', '总收入金额', '总支出笔数', '总支出金额'];

  for (const [index, row] of meaningfulRows.slice(matchedRowIndex).entries()) {
    const normalizedCells = row.cells.slice(matchedColumnIndex, matchedColumnIndex + expectedHeaderCount);

    while (normalizedCells.length < expectedHeaderCount) {
      normalizedCells.push('');
    }

    if (
      index > 0 &&
      summaryLabels.some((label) => normalizeCell(normalizedCells[0]).includes(label))
    ) {
      break;
    }

    if (index > 0 && !isRowMeaningful(normalizedCells)) {
      continue;
    }

    rows.push(normalizedCells);
    rowNumbers.push(row.rowNumber);
  }

  return {
    rows,
    rowNumbers
  };
}

function extractHeaders(filePath) {
  const rows = readRows(filePath);
  const headerRow = rows[0];

  if (!isRowMeaningful(headerRow)) {
    throw new FileValidationError('FILE_READ', '文件为空或不可读，请重新导入');
  }

  const lastMeaningfulIndex = headerRow.reduce((index, cell, currentIndex) => {
    return normalizeCell(cell) !== '' ? currentIndex : index;
  }, -1);

  if (lastMeaningfulIndex < 0) {
    throw new FileValidationError('FILE_READ', '文件为空或不可读，请重新导入');
  }

  return headerRow.slice(0, lastMeaningfulIndex + 1).map((cell) => normalizeCell(cell));
}

function loadEnumValues(enumFilePath) {
  const rows = readRows(enumFilePath);
  const firstRow = rows[0] || [];
  const shouldSkipFirstRow =
    firstRow.filter((cell) => normalizeCell(cell) !== '').length === 1 &&
    ['common字段', '映射字段', '枚举值'].includes(normalizeCell(firstRow[0]).toLowerCase());
  const values = [];
  const seen = new Set();

  rows.forEach((row, rowIndex) => {
    if (rowIndex === 0 && shouldSkipFirstRow) {
      return;
    }

    const value = normalizeCell(row[0]);

    if (!value || seen.has(value)) {
      return;
    }

    seen.add(value);
    values.push(value);
  });

  return values;
}

function extractEnumValuesFromImportedFile(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  const fileName = path.basename(filePath);

  if (extension !== '.xlsx' || !fileName.includes('枚举')) {
    throw new FileValidationError('FILE_TYPE', '请导入文件名带有“枚举”的xlsx文件');
  }

  const values = loadEnumValues(filePath);

  if (!values.length) {
    throw new FileValidationError('FILE_READ', '枚举表为空或不可读，请重新导入');
  }

  return values;
}

module.exports = {
  ensureSupportedFile,
  extractEnumValuesFromImportedFile,
  extractHeaders,
  loadEnumValues,
  readRows,
  readRowsWithMetadata
};
