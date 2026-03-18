const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

function parseArgs(argv) {
  const options = {
    appDir: process.cwd(),
    runs: 5
  };

  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index];

    if (value === '--app-dir' && argv[index + 1]) {
      options.appDir = path.resolve(argv[index + 1]);
      index += 1;
      continue;
    }

    if (value === '--runs' && argv[index + 1]) {
      const nextValue = Number(argv[index + 1]);
      if (Number.isFinite(nextValue) && nextValue > 0) {
        options.runs = Math.floor(nextValue);
      }
      index += 1;
      continue;
    }
  }

  return options;
}

function createTempDir(prefix) {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

function removeDir(dirPath) {
  fs.rmSync(dirPath, { recursive: true, force: true });
}

function average(values) {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function median(values) {
  const sorted = values.slice().sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);

  if (sorted.length % 2 === 0) {
    return (sorted[middle - 1] + sorted[middle]) / 2;
  }

  return sorted[middle];
}

function round(value) {
  return Number(value.toFixed(3));
}

function resolveElectronBinary(rootDir) {
  const candidates = process.platform === 'win32'
    ? [
        path.join(rootDir, 'node_modules', '.bin', 'electron.cmd'),
        path.join(rootDir, 'node_modules', 'electron', 'dist', 'electron.exe')
      ]
    : [
        path.join(rootDir, 'node_modules', '.bin', 'electron'),
        path.join(rootDir, 'node_modules', 'electron', 'dist', 'Electron.app', 'Contents', 'MacOS', 'Electron'),
        path.join(rootDir, 'node_modules', 'electron', 'dist', 'electron')
      ];

  const matched = candidates.find((candidate) => fs.existsSync(candidate));

  if (!matched) {
    throw new Error('未找到可用的 Electron 可执行文件，请先执行 npm ci');
  }

  return matched;
}

function measureOnce({ rootDir, appDir, electronBinary }) {
  const tempRoot = createTempDir('bank-bill-startup-');
  const capturePath = path.join(tempRoot, 'capture.png');
  const metricsPath = path.join(tempRoot, 'startup-metrics.json');
  const userDataDir = path.join(tempRoot, 'userdata');
  const documentsDir = path.join(tempRoot, 'documents');

  try {
    const startedAt = process.hrtime.bigint();
    const child = spawnSync(electronBinary, [appDir], {
      cwd: rootDir,
      env: {
        ...process.env,
        APP_CAPTURE_PATH: capturePath,
        APP_CAPTURE_DELAY_MS: '0',
        APP_STARTUP_METRICS_PATH: metricsPath,
        APP_USER_DATA_DIR: userDataDir,
        APP_DOCUMENTS_DIR: documentsDir
      },
      encoding: 'utf8'
    });
    const elapsedMs = Number(process.hrtime.bigint() - startedAt) / 1_000_000;

    if (child.error) {
      throw child.error;
    }

    if (child.status !== 0) {
      throw new Error(`Electron 退出码为 ${child.status}\n${child.stderr || child.stdout}`);
    }

    if (!fs.existsSync(metricsPath)) {
      throw new Error('未生成启动 metrics 文件，无法完成启动测量');
    }

    const metrics = JSON.parse(fs.readFileSync(metricsPath, 'utf8'));

    return {
      elapsedMs: round(elapsedMs),
      captureExists: fs.existsSync(capturePath),
      metrics
    };
  } finally {
    removeDir(tempRoot);
  }
}

function summarizeRuns(runs) {
  const totalValues = runs.map((run) => run.metrics.durations.totalReadyToShowMs);
  const windowValues = runs.map((run) => run.metrics.durations.createWindowToReadyMs);
  const loadValues = runs.map((run) => run.metrics.durations.loadToReadyMs);
  const processValues = runs.map((run) => run.elapsedMs);
  const rendererTotals = runs
    .map((run) => run.metrics.renderer?.durations?.totalInitMs)
    .filter((value) => typeof value === 'number');
  const rendererGetInfo = runs
    .map((run) => run.metrics.renderer?.durations?.getInfoMs)
    .filter((value) => typeof value === 'number');
  const rendererRefreshTemplates = runs
    .map((run) => run.metrics.renderer?.durations?.refreshTemplatesMs)
    .filter((value) => typeof value === 'number');
  const rendererBindEvents = runs
    .map((run) => run.metrics.renderer?.durations?.bindEventsMs)
    .filter((value) => typeof value === 'number');

  return {
    processElapsedMs: {
      average: round(average(processValues)),
      median: round(median(processValues)),
      min: round(Math.min(...processValues)),
      max: round(Math.max(...processValues))
    },
    readyToShowMs: {
      average: round(average(totalValues)),
      median: round(median(totalValues)),
      min: round(Math.min(...totalValues)),
      max: round(Math.max(...totalValues))
    },
    createWindowToReadyMs: {
      average: round(average(windowValues)),
      median: round(median(windowValues)),
      min: round(Math.min(...windowValues)),
      max: round(Math.max(...windowValues))
    },
    loadToReadyMs: {
      average: round(average(loadValues)),
      median: round(median(loadValues)),
      min: round(Math.min(...loadValues)),
      max: round(Math.max(...loadValues))
    },
    renderer: rendererTotals.length
      ? {
          totalInitMs: {
            average: round(average(rendererTotals)),
            median: round(median(rendererTotals)),
            min: round(Math.min(...rendererTotals)),
            max: round(Math.max(...rendererTotals))
          },
          getInfoMs: {
            average: round(average(rendererGetInfo)),
            median: round(median(rendererGetInfo)),
            min: round(Math.min(...rendererGetInfo)),
            max: round(Math.max(...rendererGetInfo))
          },
          refreshTemplatesMs: {
            average: round(average(rendererRefreshTemplates)),
            median: round(median(rendererRefreshTemplates)),
            min: round(Math.min(...rendererRefreshTemplates)),
            max: round(Math.max(...rendererRefreshTemplates))
          },
          bindEventsMs: {
            average: round(average(rendererBindEvents)),
            median: round(median(rendererBindEvents)),
            min: round(Math.min(...rendererBindEvents)),
            max: round(Math.max(...rendererBindEvents))
          }
        }
      : null
  };
}

function printRunDetails(runs) {
  runs.forEach((run, index) => {
    const durations = run.metrics.durations;
    const rendererDurations = run.metrics.renderer?.durations;
    console.log(
      [
        `Run ${index + 1}`,
        `进程总耗时 ${run.elapsedMs}ms`,
        `ready-to-show ${durations.totalReadyToShowMs}ms`,
        `建窗到可见 ${durations.createWindowToReadyMs}ms`,
        `loadFile 到可见 ${durations.loadToReadyMs}ms`,
        ...(rendererDurations
          ? [
              `渲染初始化 ${rendererDurations.totalInitMs}ms`,
              `getInfo ${rendererDurations.getInfoMs}ms`,
              `模板刷新 ${rendererDurations.refreshTemplatesMs}ms`,
              `事件绑定 ${rendererDurations.bindEventsMs}ms`
            ]
          : [])
      ].join(' | ')
    );
  });
}

function printSummary(summary) {
  console.log('\nSummary');
  console.log(`进程总耗时(平均/中位/最小/最大): ${summary.processElapsedMs.average} / ${summary.processElapsedMs.median} / ${summary.processElapsedMs.min} / ${summary.processElapsedMs.max} ms`);
  console.log(`ready-to-show(平均/中位/最小/最大): ${summary.readyToShowMs.average} / ${summary.readyToShowMs.median} / ${summary.readyToShowMs.min} / ${summary.readyToShowMs.max} ms`);
  console.log(`建窗到可见(平均/中位/最小/最大): ${summary.createWindowToReadyMs.average} / ${summary.createWindowToReadyMs.median} / ${summary.createWindowToReadyMs.min} / ${summary.createWindowToReadyMs.max} ms`);
  console.log(`loadFile 到可见(平均/中位/最小/最大): ${summary.loadToReadyMs.average} / ${summary.loadToReadyMs.median} / ${summary.loadToReadyMs.min} / ${summary.loadToReadyMs.max} ms`);
  if (summary.renderer) {
    console.log(`渲染初始化(平均/中位/最小/最大): ${summary.renderer.totalInitMs.average} / ${summary.renderer.totalInitMs.median} / ${summary.renderer.totalInitMs.min} / ${summary.renderer.totalInitMs.max} ms`);
    console.log(`渲染 getInfo(平均/中位/最小/最大): ${summary.renderer.getInfoMs.average} / ${summary.renderer.getInfoMs.median} / ${summary.renderer.getInfoMs.min} / ${summary.renderer.getInfoMs.max} ms`);
    console.log(`渲染 模板刷新(平均/中位/最小/最大): ${summary.renderer.refreshTemplatesMs.average} / ${summary.renderer.refreshTemplatesMs.median} / ${summary.renderer.refreshTemplatesMs.min} / ${summary.renderer.refreshTemplatesMs.max} ms`);
    console.log(`渲染 事件绑定(平均/中位/最小/最大): ${summary.renderer.bindEventsMs.average} / ${summary.renderer.bindEventsMs.median} / ${summary.renderer.bindEventsMs.min} / ${summary.renderer.bindEventsMs.max} ms`);
  }
}

function main() {
  const options = parseArgs(process.argv.slice(2));
  const rootDir = process.cwd();
  const electronBinary = resolveElectronBinary(rootDir);
  const runs = [];

  for (let index = 0; index < options.runs; index += 1) {
    runs.push(measureOnce({
      rootDir,
      appDir: options.appDir,
      electronBinary
    }));
  }

  const summary = summarizeRuns(runs);
  printRunDetails(runs);
  printSummary(summary);
}

main();
