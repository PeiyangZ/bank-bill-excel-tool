const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { Data } = require('resedit');

const projectRoot = path.resolve(__dirname, '..');
const assetsDir = path.join(projectRoot, 'assets');
const buildDir = path.join(projectRoot, 'build');
const bundledSourcePath = path.join(assetsDir, 'app-icon-source.png');
const runtimePngPath = path.join(assetsDir, 'app-icon.png');
const runtimeIcoPath = path.join(assetsDir, 'app-icon.ico');
const buildPngPath = path.join(buildDir, 'icon.png');
const buildIcoPath = path.join(buildDir, 'icon.ico');
const iconSizes = [16, 24, 32, 48, 64, 128, 256];

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function runSips(args) {
  execFileSync('/usr/bin/sips', args, { stdio: 'pipe' });
}

function readDimensions(filePath) {
  const output = execFileSync('/usr/bin/sips', ['-g', 'pixelWidth', '-g', 'pixelHeight', filePath], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe']
  });
  const widthMatch = output.match(/pixelWidth:\s+(\d+)/);
  const heightMatch = output.match(/pixelHeight:\s+(\d+)/);
  if (!widthMatch || !heightMatch) {
    throw new Error(`无法读取图片尺寸: ${filePath}`);
  }
  return {
    width: Number(widthMatch[1]),
    height: Number(heightMatch[1])
  };
}

function normalizeSourceImage(sourcePath, outputPath) {
  const { width, height } = readDimensions(sourcePath);
  const side = Math.min(width, height);
  runSips(['--cropToHeightWidth', String(side), String(side), sourcePath, '--out', outputPath]);
}

function ensureDirectory(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function buildIcoFromPng(sourcePath, outputPath) {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bank-bill-icon-'));
  try {
    const iconFile = new Data.IconFile();
    iconFile.icons = iconSizes.map((size) => {
      const sizePath = path.join(tempDir, `${size}.png`);
      runSips(['-z', String(size), String(size), sourcePath, '--out', sizePath]);
      const buffer = fs.readFileSync(sizePath);
      return {
        width: size,
        height: size,
        bitCount: 32,
        data: Data.RawIconItem.from(toArrayBuffer(buffer), size, size, 32)
      };
    });
    fs.writeFileSync(outputPath, Buffer.from(iconFile.generate()));
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

function main() {
  const inputPath = process.argv[2] ? path.resolve(process.argv[2]) : bundledSourcePath;

  if (!fs.existsSync(inputPath)) {
    throw new Error(`图标源文件不存在: ${inputPath}`);
  }

  ensureDirectory(assetsDir);
  ensureDirectory(buildDir);

  normalizeSourceImage(inputPath, bundledSourcePath);
  fs.copyFileSync(bundledSourcePath, runtimePngPath);
  fs.copyFileSync(bundledSourcePath, buildPngPath);
  buildIcoFromPng(bundledSourcePath, runtimeIcoPath);
  fs.copyFileSync(runtimeIcoPath, buildIcoPath);

  console.log(`已同步应用图标: ${bundledSourcePath}`);
}

main();
