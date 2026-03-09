const path = require('node:path');
const XLSX = require('xlsx');

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.aoa_to_sheet([
  ['COMMON字段'],
  ['交易日期'],
  ['交易时间'],
  ['摘要'],
  ['对方户名'],
  ['对方账号'],
  ['我方账号'],
  ['交易金额'],
  ['账户余额'],
  ['币种'],
  ['借贷标识'],
  ['用途'],
  ['备注']
]);

XLSX.utils.book_append_sheet(workbook, worksheet, '枚举');
XLSX.writeFile(workbook, path.join(process.cwd(), 'COMMON枚举.xlsx'));
