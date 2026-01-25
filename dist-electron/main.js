"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const electron_1 = require("electron");
const path_1 = __importDefault(require("path"));
// Handle creating/removing shortcuts on Windows when installing/uninstalling.
if (require('electron-squirrel-startup')) {
    electron_1.app.quit();
}
const createWindow = () => {
    const mainWindow = new electron_1.BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            preload: path_1.default.join(__dirname, 'preload.js'),
            nodeIntegration: true,
            contextIsolation: false,
            webSecurity: false // Optional, but helps avoid some local file loading issues
        },
    });
    if (!electron_1.app.isPackaged) {
        mainWindow.loadURL('http://localhost:5173');
        mainWindow.webContents.openDevTools();
    }
    else {
        mainWindow.loadFile(path_1.default.join(__dirname, '../dist/index.html'));
    }
};
electron_1.app.whenReady().then(() => {
    createWindow();
    electron_1.app.on('activate', () => {
        if (electron_1.BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});
electron_1.app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        electron_1.app.quit();
    }
});
// IPC Handlers
electron_1.ipcMain.handle('select-file', async () => {
    const { dialog } = require('electron');
    const result = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'PowerPoint', extensions: ['pptx'] }]
    });
    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }
    return result.filePaths[0];
});
electron_1.ipcMain.handle('convert-pptx', async (event, filePath) => {
    console.log('Convert request for (raw):', filePath);
    // Resolve to absolute path to handle relative paths (e.g. from dev environment)
    const absolutePath = path_1.default.resolve(filePath);
    console.log('Convert request for (absolute):', absolutePath);
    const fs = require('fs');
    if (!fs.existsSync(absolutePath)) {
        return { success: false, error: `File not found: ${absolutePath}` };
    }
    const os = process.platform;
    const tempDir = electron_1.app.getPath('temp');
    const outputDir = path_1.default.join(tempDir, 'ppt-viewer', path_1.default.basename(absolutePath, path_1.default.extname(absolutePath)));
    console.log('Output Dir:', outputDir);
    try {
        if (os === 'win32') {
            const scriptPath = path_1.default.join(__dirname, '../electron/scripts/convert-win.ps1');
            console.log('Script Path:', scriptPath);
            // Spawn PowerShell
            const { spawn } = require('child_process');
            // We use 'powershell.exe' directly
            const child = spawn('powershell.exe', [
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', scriptPath,
                '-InputPath', absolutePath,
                '-OutputDir', outputDir
            ]);
            return new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (data) => {
                    console.log(`stdout: ${data}`);
                    stdout += data;
                });
                child.stderr.on('data', (data) => {
                    console.error(`stderr: ${data}`);
                    stderr += data;
                });
                child.on('close', async (code) => {
                    if (code === 0) {
                        // Read manifest
                        const manifestPath = path_1.default.join(outputDir, 'manifest.json');
                        try {
                            const fs = require('fs'); // Use sync fs for simplicity inside async callback, or keep promises
                            // Using readFileSync to be safe with blocking logic if preferred, or promises
                            const data = fs.readFileSync(manifestPath, 'utf8').replace(/^\uFEFF/, '');
                            const slides = JSON.parse(data);
                            // Fix image paths to be absolute or protocol based
                            const slidesWithPaths = slides.map((s) => ({
                                ...s,
                                src: `file://${path_1.default.join(outputDir, s.image)}`
                            }));
                            resolve({ success: true, slides: slidesWithPaths });
                        }
                        catch (err) {
                            reject(err);
                        }
                    }
                    else {
                        reject(new Error(`Conversion failed with code ${code}: ${stderr}`));
                    }
                });
            });
        }
        // TODO: Implement other platforms
        return { success: false, error: 'Platform not supported yet' };
    }
    catch (err) {
        console.error('Conversion error:', err);
        return { success: false, error: err.message };
    }
});
electron_1.ipcMain.handle('save-all-notes', async (event, filePath, slides) => {
    console.log('Save All Notes request for:', filePath);
    // Resolve absolute path
    const absolutePath = path_1.default.resolve(filePath);
    const fs = require('fs');
    if (!fs.existsSync(absolutePath)) {
        return { success: false, error: 'File not found' };
    }
    // Create temp JSON file for updates
    // We only need index and notes
    const updates = slides.map((s) => ({
        index: s.index,
        notes: s.notes
    }));
    const tempDir = electron_1.app.getPath('temp');
    const updateJsonPath = path_1.default.join(tempDir, `ppt-update-${Date.now()}.json`);
    try {
        require('fs').writeFileSync(updateJsonPath, JSON.stringify(updates, null, 2), 'utf8');
        const os = process.platform;
        if (os === 'win32') {
            const scriptPath = path_1.default.join(__dirname, '../electron/scripts/save-all-notes-win.ps1');
            const { spawn } = require('child_process');
            const child = spawn('powershell.exe', [
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', scriptPath,
                '-InputPath', absolutePath,
                '-NotesJsonPath', updateJsonPath
            ]);
            return new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (data) => stdout += data);
                child.stderr.on('data', (data) => stderr += data);
                child.on('close', (code) => {
                    // Cleanup temp file
                    try {
                        fs.unlinkSync(updateJsonPath);
                    }
                    catch (e) { }
                    if (code === 0) {
                        resolve({ success: true });
                    }
                    else {
                        reject(new Error(`Save failed: ${stderr} | ${stdout}`));
                    }
                });
            });
        }
        return { success: false, error: 'Save not supported on this platform yet' };
    }
    catch (e) {
        return { success: false, error: e.message };
    }
});
electron_1.ipcMain.handle('generate-video', async (event, { filePath, slidesAudio }) => {
    console.log('MAIN: generate-video handler called');
    console.log('Generate Video request for:', filePath);
    const fs = require('fs');
    const absolutePath = path_1.default.resolve(filePath);
    if (!fs.existsSync(absolutePath)) {
        return { success: false, error: 'PPT File not found' };
    }
    // 1. Show Save Dialog
    const { dialog } = require('electron');
    const { filePath: outputPath } = await dialog.showSaveDialog({
        title: 'Save Video',
        defaultPath: path_1.default.basename(absolutePath, '.pptx') + '.mp4',
        filters: [{ name: 'MPEG-4 Video', extensions: ['mp4'] }]
    });
    if (!outputPath) {
        return { success: false, error: 'User cancelled save' };
    }
    // 2. Save Audio Files to Temp Dir
    const tempDir = fs.mkdtempSync(path_1.default.join(electron_1.app.getPath('temp'), 'ppt-video-gen-'));
    console.log('Using temp dir for audio:', tempDir);
    try {
        for (const item of slidesAudio) {
            // item.audioData is Uint8Array/Buffer
            const audioPath = path_1.default.join(tempDir, `slide_${item.index}.wav`);
            fs.writeFileSync(audioPath, Buffer.from(item.audioData));
        }
        // 3. Call PowerShell Script
        const os = process.platform;
        if (os === 'win32') {
            const scriptPath = path_1.default.join(__dirname, '../electron/scripts/generate-video-win.ps1');
            const { spawn } = require('child_process');
            const child = spawn('powershell.exe', [
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', scriptPath,
                '-InputPath', absolutePath,
                '-AudioDir', tempDir,
                '-OutputPath', outputPath
            ]);
            return new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (data) => { console.log('PS:', data.toString()); stdout += data; });
                child.stderr.on('data', (data) => { console.error('PS Err:', data.toString()); stderr += data; });
                child.on('close', (code) => {
                    // Cleanup temp files
                    try {
                        fs.rmSync(tempDir, { recursive: true, force: true });
                    }
                    catch (e) {
                        console.error("Cleanup error", e);
                    }
                    if (code === 0) {
                        resolve({ success: true, outputPath });
                    }
                    else {
                        reject(new Error(`Generation failed: ${stderr || stdout}`));
                    }
                });
            });
        }
        if (os === 'darwin') {
            const scriptPath = path_1.default.join(__dirname, '../electron/scripts/generate-video-mac.applescript');
            console.log('Script Path (Mac):', scriptPath);
            const { spawn } = require('child_process');
            // Spawn osascript
            const child = spawn('osascript', [
                scriptPath,
                absolutePath,
                tempDir,
                outputPath
            ]);
            return new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (data) => { console.log('OSAScript:', data.toString()); stdout += data; });
                child.stderr.on('data', (data) => { console.error('OSAScript Err:', data.toString()); stderr += data; });
                child.on('close', (code) => {
                    // Cleanup temp files
                    try {
                        fs.rmSync(tempDir, { recursive: true, force: true });
                    }
                    catch (e) {
                        console.error("Cleanup error", e);
                    }
                    if (code === 0) {
                        resolve({ success: true, outputPath });
                    }
                    else {
                        reject(new Error(`Generation failed: ${stderr || stdout}`));
                    }
                });
            });
        }
        return { success: false, error: 'Platform not supported' };
    }
    catch (e) {
        // Cleanup on error
        try {
            fs.rmSync(tempDir, { recursive: true, force: true });
        }
        catch (cleanupErr) { }
        console.error("Generate Video Error:", e);
        return { success: false, error: e.message };
    }
});
