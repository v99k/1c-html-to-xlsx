const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const { convertFile } = require('./converter');

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, '../preload/preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'));
}

app.whenReady().then(createWindow);
app.on('window-all-closed', () => app.quit());

function getTempDir() {
  const base = app.getPath('temp');
  const sub = 'HtmlToXlsx_' + new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  return path.join(base, sub);
}

ipcMain.handle('choose-html-files', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'HTML', extensions: ['html', 'htm'] }, { name: 'All', extensions: ['*'] },
    ],
  });
  return canceled ? [] : filePaths;
});

ipcMain.handle('choose-save-dir', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
  });
  return canceled ? null : (filePaths && filePaths[0]) || null;
});

ipcMain.handle('convert', async (_, htmlPaths) => {
  const tempDir = getTempDir();
  await fs.mkdir(tempDir, { recursive: true });
  const logPath = path.join(tempDir, 'log.txt');
  const logLines = [];
  const log = (level, msg) => {
    const line = `[${new Date().toISOString()}] [${level}] ${msg}`;
    logLines.push(line);
    mainWindow?.webContents?.send('log', line);
  };
  const results = [];
  for (let i = 0; i < htmlPaths.length; i++) {
    const htmlPath = htmlPaths[i];
    const baseName = path.basename(htmlPath, path.extname(htmlPath));
    let xlsxPath = path.join(tempDir, baseName + '.xlsx');
    let n = 0;
    while (await fs.access(xlsxPath).then(() => true).catch(() => false)) {
      n++;
      xlsxPath = path.join(tempDir, `${baseName} (${n}).xlsx`);
    }
    try {
      log('INFO', `Конвертация: ${htmlPath}`);
      await convertFile(htmlPath, xlsxPath, log);
      log('INFO', `Готово: ${xlsxPath}`);
      results.push({ htmlPath, xlsxPath, ok: true });
    } catch (e) {
      log('ERROR', `${htmlPath}: ${e.message}`);
      results.push({ htmlPath, xlsxPath: null, ok: false, error: e.message });
    }
    mainWindow?.webContents?.send('progress', { done: i + 1, total: htmlPaths.length });
  }
  try {
    await fs.writeFile(logPath, logLines.join('\n'), 'utf8');
  } catch (_) {}
  return { tempDir, logPath, results };
});

ipcMain.handle('save-results', async (_, results, saveDir) => {
  const saved = [];
  for (const r of results) {
    if (!r.ok || !r.xlsxPath) continue;
    const baseName = path.basename(r.htmlPath, path.extname(r.htmlPath));
    let outPath = path.join(saveDir, baseName + '.xlsx');
    let n = 0;
    while (await fs.access(outPath).then(() => true).catch(() => false)) {
      n++;
      outPath = path.join(saveDir, `${baseName} (${n}).xlsx`);
    }
    await fs.copyFile(r.xlsxPath, outPath);
    saved.push(outPath);
  }
  return saved;
});
