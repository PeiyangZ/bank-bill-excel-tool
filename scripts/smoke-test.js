const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const XLSX = require('xlsx');
const { AppDatabase } = require('../src/backend/database');
const {
  buildMappedRows,
  extractHeaders,
  loadEnumValues,
  writeBalanceWorkbook,
  writeWorkbookRows
} = require('../src/backend/file-service');

function makeWorkbook(filePath, rows) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, filePath);
}

function run() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'bank-bill-tool-'));
  const projectRoot = path.resolve(__dirname, '..');
  const bundledEnumPath = path.join(projectRoot, 'COMMON枚举.xlsx');
  const dbPath = path.join(root, 'app.sqlite');
  const templatePath = path.join(root, 'template.xlsx');
  const dataPath = path.join(root, 'input.xlsx');
  const detailOutputPath = path.join(root, '2026-03-09', 'detail', 'template-Balance-2026-03-09.xlsx');
  const balanceTemplatePath = path.join(root, 'balance-template.xlsx');
  const balanceOutputPath = path.join(root, '2026-03-09', 'balance', 'template-Balance-2026-03-09.xlsx');

  makeWorkbook(templatePath, [['原字段A', '原字段B', '原字段C', '原字段D', '原字段E', '原字段F'], ['值1', '值2', '值3', '值4', '值5', '值6']]);
  makeWorkbook(dataPath, [['原字段A', '原字段B', '原字段C', '原字段D', '原字段E', '原字段F', '原字段G'], ['1234.56', '789.01', '2026-03-09', '20260310', 'NET_001', 88, '456.78']]);
  makeWorkbook(balanceTemplatePath, [['银行名称', '所在地', '币种', '银行账号', '账单日期', '期初余额', '期初可用余额', '期末余额', '期末可用余额']]);

  const db = new AppDatabase(dbPath);
  db.init();

  const headers = extractHeaders(templatePath);
  assert.deepStrictEqual(headers, ['原字段A', '原字段B', '原字段C', '原字段D', '原字段E', '原字段F']);

  const template = db.upsertTemplate({
    name: 'template',
    sourceFileName: 'template.xlsx',
    headers
  });

  db.setBackgroundConfig({
    colorHex: '#123456',
    filePath: '',
    sourceFileName: ''
  });
  assert.strictEqual(db.getBackgroundConfig().colorHex, '#123456');

  db.saveMappings(template.id, [
    { templateField: '原字段A', mappedField: 'Credit Amount' },
    { templateField: '原字段B', mappedField: 'Debit Amount' },
    { templateField: '原字段C', mappedField: 'BillDate' },
    { templateField: '原字段D', mappedField: 'ValueDate' },
    { templateField: '原字段E', mappedField: 'MerchantId' },
    { templateField: '原字段F', mappedField: 'Channel' }
  ]);
  db.saveAccountMappings([
    {
      bankAccountId: 'NET_001',
      clearingAccountId: 'CLEAR_9001'
    }
  ]);

  assert(fs.existsSync(bundledEnumPath));
  const enumValues = loadEnumValues(bundledEnumPath);
  assert.strictEqual(enumValues[0], 'BillDate');
  assert(enumValues.includes('Credit Amount'));
  assert(enumValues.includes('MerchantId'));
  assert.strictEqual(enumValues.includes('COMMON字段'), false);

  const detailRows = buildMappedRows({
    inputFilePath: dataPath,
    mappingByField: {
      Balance: '原字段G',
      'Credit Amount': '原字段A',
      'Debit Amount': '原字段B',
      BillDate: '原字段C',
      ValueDate: '原字段D',
      MerchantId: '原字段E',
      Channel: '原字段F'
    },
    orderedTargetFields: ['Balance', 'Credit Amount', 'Debit Amount', 'BillDate', 'ValueDate', 'MerchantId', 'Channel'],
    accountMappingByBankId: {
      NET_001: 'CLEAR_9001'
    }
  });
  writeWorkbookRows({
    rows: detailRows,
    outputFilePath: detailOutputPath
  });

  writeBalanceWorkbook({
    templateFilePath: balanceTemplatePath,
    records: [['CHB', 'HK', 'USD', 'CLEAR_9001', '2026-03-09', '', '', 456.78, '']],
    outputFilePath: balanceOutputPath
  });

  assert(fs.existsSync(detailOutputPath));
  const workbook = XLSX.readFile(detailOutputPath, {
    cellNF: true,
    cellStyles: true,
    raw: true
  });
  const worksheet = workbook.Sheets.COMMON;
  const rows = XLSX.utils.sheet_to_json(
    worksheet,
    { header: 1, defval: '' }
  );
  assert.deepStrictEqual(rows[0], ['Balance', 'Credit Amount', 'Debit Amount', 'BillDate', 'ValueDate', 'MerchantId', 'Channel']);
  assert.strictEqual(rows[1][0], 456.78);
  assert.strictEqual(rows[1][5], 'CLEAR_9001');
  assert.strictEqual(rows[1][6], '88');
  assert.strictEqual(worksheet.A2.t, 'n');
  assert.strictEqual(worksheet.A2.z, '0.00');
  assert.strictEqual(worksheet.B2.t, 'n');
  assert.strictEqual(worksheet.B2.z, '0.00');
  assert.strictEqual(worksheet.C2.t, 'n');
  assert.strictEqual(worksheet.C2.z, '0.00');
  assert.strictEqual(worksheet.D2.t, 'n');
  assert.strictEqual(worksheet.D2.z, 'yyyy-mm-dd');
  assert.strictEqual(worksheet.E2.t, 'n');
  assert.strictEqual(worksheet.E2.z, 'yyyy-mm-dd');
  assert.strictEqual(worksheet.F2.t, 's');
  assert.strictEqual(worksheet.F2.z, '@');
  assert.strictEqual(worksheet.G2.t, 's');
  assert.strictEqual(worksheet.G2.z, '@');

  assert(fs.existsSync(balanceOutputPath));
  const balanceWorkbook = XLSX.readFile(balanceOutputPath, { raw: true });
  const balanceRows = XLSX.utils.sheet_to_json(balanceWorkbook.Sheets[balanceWorkbook.SheetNames[0]], {
    header: 1,
    defval: ''
  });
  assert.strictEqual(balanceRows[1][0], 'CHB');
  assert.strictEqual(balanceRows[1][7], 456.78);

  console.log('smoke test passed');
}

run();
