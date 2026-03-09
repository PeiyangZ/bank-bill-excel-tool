const fs = require('node:fs');
const path = require('node:path');

function pad(value) {
  return String(value).padStart(2, '0');
}

function formatLocalDate(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function formatLocalTimestamp(date) {
  return `${formatLocalDate(date)} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function appendLog(logRoot, error) {
  const now = new Date();
  const date = formatLocalDate(now);
  const time = formatLocalTimestamp(now);
  const targetDir = path.join(logRoot, 'logs');
  const targetFile = path.join(targetDir, `${date}.log`);

  fs.mkdirSync(targetDir, { recursive: true });
  fs.appendFileSync(
    targetFile,
    `[${time}] ${error.stack || error.message || String(error)}\n`,
    'utf8'
  );

  return targetFile;
}

module.exports = {
  appendLog
};
