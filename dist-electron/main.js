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
    // Ensure parent dir exists
    if (!fs.existsSync(path_1.default.join(tempDir, 'ppt-viewer'))) {
        fs.mkdirSync(path_1.default.join(tempDir, 'ppt-viewer'));
    }
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
        else if (os === 'darwin') {
            const scriptPath = path_1.default.join(__dirname, '../electron/scripts/convert-mac.applescript');
            console.log('Script Path (Mac):', scriptPath);
            const { spawn } = require('child_process');
            const child = spawn('osascript', [
                scriptPath,
                absolutePath,
                outputDir
            ]);
            return new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (data) => { console.log('OSAScript:', data.toString()); stdout += data; });
                child.stderr.on('data', (data) => { console.error('OSAScript Err:', data.toString()); stderr += data; });
                child.on('close', (code) => {
                    if (code === 0) {
                        try {
                            const fs = require('fs');
                            console.log('--- Output Directory Contents ---');
                            // Simple recursive list for debugging
                            const listDir = (dir) => {
                                const files = fs.readdirSync(dir);
                                files.forEach((file) => {
                                    const p = path_1.default.join(dir, file);
                                    if (fs.statSync(p).isDirectory()) {
                                        console.log(`[DIR] ${p}`);
                                        listDir(p);
                                    }
                                    else {
                                        console.log(`[FILE] ${p}`);
                                    }
                                });
                            };
                            listDir(outputDir);
                            console.log('-------------------------------');
                        }
                        catch (e) {
                            console.error('Error listing dir:', e);
                        }
                        // Read manifest
                        const manifestPath = path_1.default.join(outputDir, 'manifest.json');
                        try {
                            const fs = require('fs');
                            const data = fs.readFileSync(manifestPath, 'utf8').replace(/^\uFEFF/, '');
                            console.log('Manifest Content:', data); // LOGGING ADDED
                            const slides = JSON.parse(data);
                            const slidesWithPaths = slides.map((s) => ({
                                ...s,
                                src: s.image ? `file://${path_1.default.join(outputDir, s.image)}` : null
                            })).filter((s) => s.src !== null); // Filter out bad slides
                            resolve({ success: true, slides: slidesWithPaths });
                        }
                        catch (err) {
                            reject(err);
                        }
                    }
                    else {
                        reject(new Error(`Conversion failed with code ${code}: ${stderr || stdout}`));
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
    try {
        if (process.platform === 'darwin') {
            const app = require('electron').app;
            const homeDir = app.getPath('home');
            const officeContainer = path_1.default.join(homeDir, 'Library/Group Containers/UBF8T346G9.Office');
            // 1. Prepare Data File
            // Format:
            // ###SLIDE_START### 1
            // Notes...
            // ###SLIDE_END###
            let dataContent = "";
            for (const s of slides) {
                if (s.notes) {
                    dataContent += `###SLIDE_START### ${s.index}\n${s.notes}\n###SLIDE_END###\n`;
                }
            }
            const dataPath = path_1.default.join(officeContainer, `notes_data_${Date.now()}.txt`);
            fs.writeFileSync(dataPath, dataContent, 'utf8');
            // 2. Prepare Params File
            const paramsPath = path_1.default.join(officeContainer, 'notes_params.txt');
            // Content: PresentationPath|DataPath
            const paramsContent = `${absolutePath}|${dataPath}`;
            fs.writeFileSync(paramsPath, paramsContent, 'utf8');
            // 3. Trigger Macro
            const scriptPath = path_1.default.join(__dirname, '../electron/scripts/trigger-macro.applescript');
            // Args: macroName, pptPath
            const { spawn } = require('child_process');
            const child = spawn('osascript', [
                scriptPath,
                "UpdateNotes",
                absolutePath
            ]);
            await new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (d) => stdout += d);
                child.stderr.on('data', (d) => stderr += d);
                child.on('close', (code) => {
                    // Cleanup data file
                    try {
                        fs.unlinkSync(dataPath);
                    }
                    catch (e) { }
                    if (code === 0 && !stdout.includes("Error")) {
                        console.log(`UpdateNotes macro triggered.`);
                        resolve();
                    }
                    else {
                        console.error(`Failed to trigger UpdateNotes: ${stderr} ${stdout}`);
                        reject(new Error(stdout || stderr));
                    }
                });
            });
            return { success: true };
        }
        else {
            // Windows implementation...
            return { success: false, error: 'Save not supported on this platform yet' };
        }
    }
    catch (e) {
        return { success: false, error: e.message };
    }
});
electron_1.ipcMain.handle('get-video-save-path', async () => {
    const { dialog } = require('electron');
    const win = require('electron').BrowserWindow.getFocusedWindow();
    const app = require('electron').app;
    const path = require('path');
    const result = await dialog.showSaveDialog(win, {
        title: 'Save Video As',
        defaultPath: path.join(app.getPath('documents'), 'Output.mp4'),
        filters: [{ name: 'MPEG-4 Video', extensions: ['mp4'] }]
    });
    if (result.canceled || !result.filePath) {
        return null;
    }
    return result.filePath;
});
electron_1.ipcMain.handle('generate-video', async (event, { filePath, slidesAudio, videoOutputPath }) => {
    console.log('Generate Video Request (PPAM Flow)');
    const path = require('path');
    const fs = require('fs');
    const { spawn } = require('child_process');
    const os = require('os');
    const app = require('electron').app;
    // videoOutputPath is now passed in
    if (!videoOutputPath) {
        return { success: false, error: "No output path provided." };
    }
    console.log("Target Video Path:", videoOutputPath);
    // Define the Office Group Container path for sandboxed access
    const homeDir = app.getPath('home');
    const officeContainer = path.join(homeDir, 'Library/Group Containers/UBF8T346G9.Office');
    const audioSessionDir = path.join(officeContainer, 'TemporaryAudio', `session-${Date.now()}`);
    // Ensure directory exists
    try {
        fs.mkdirSync(audioSessionDir, { recursive: true });
    }
    catch (e) {
        console.error("Failed to create Office container dir:", e);
        return { success: false, error: "Could not create audio directory in Office container. Check permissions." };
    }
    try {
        if (slidesAudio.length > 0) {
            let batchParams = "";
            // 1. Save all audio files and prepare batch params
            for (const slide of slidesAudio) {
                console.log(`Processing slide ${slide.index}`);
                const buffer = Buffer.from(slide.audioData);
                const audioFileName = `audio_${slide.index}.mp3`;
                const audioFilePath = path.join(audioSessionDir, audioFileName);
                fs.writeFileSync(audioFilePath, buffer);
                console.log(`Saved audio to ${audioFilePath}`);
                // Append to batch params: Index|AudioPath|PresentationPath
                batchParams += `${slide.index}|${audioFilePath}|${filePath}\n`;
            }
            if (process.platform === 'darwin') {
                const scriptPath = path.join(__dirname, '../electron/scripts/trigger-macro.applescript');
                // 2. Write ALL parameters to file ONCE
                // audio_params.txt content: "SlideIndex|AudioPath|PresentationPath" (Multiple lines)
                const paramsPath = path.join(officeContainer, 'audio_params.txt');
                fs.writeFileSync(paramsPath, batchParams, 'utf8');
                // 3. Call the GENERIC macro runner ONCE
                // Args: macroName, pptPath
                console.log("Triggering batch audio insertion macro...");
                const child = spawn('osascript', [
                    scriptPath,
                    "InsertAudio",
                    filePath
                ]);
                await new Promise((resolve, reject) => {
                    let stdout = '';
                    let stderr = '';
                    child.stdout.on('data', (d) => stdout += d);
                    child.stderr.on('data', (d) => stderr += d);
                    child.on('close', (code) => {
                        if (code === 0 && !stdout.includes("Error")) {
                            console.log(`Batch audio macro completed successfully.`);
                            resolve();
                        }
                        else {
                            console.error(`Failed to run batch audio macro: ${stderr} ${stdout}`);
                            reject(new Error(stdout || stderr));
                        }
                    });
                });
            }
            else {
                // Windows fallback (if needed later)
            }
        }
        if (process.platform === 'darwin') {
            const exportScriptPath = path.join(__dirname, '../electron/scripts/export-to-video.applescript');
            // Args: outputPath, presentationPath (filePath)
            const child = spawn('osascript', [
                exportScriptPath,
                videoOutputPath,
                filePath
            ]);
            await new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (d) => stdout += d);
                child.stderr.on('data', (d) => stderr += d);
                child.on('close', (code) => {
                    // Note: 'save as movie' might return 0 but export continues in PPT background.
                    if (code === 0 && !stdout.includes("Error")) {
                        console.log(`Video export initiated: ${videoOutputPath}`);
                        resolve();
                    }
                    else {
                        console.error(`Export failed: ${stderr || stdout}`);
                        reject(new Error(stdout || stderr));
                    }
                });
            });
            return { success: true, outputPath: videoOutputPath };
        }
        else {
            return { success: false, error: "Windows video export not implemented" };
        }
    }
    catch (e) {
        console.error('Generation failed:', e);
        return { success: false, error: e.message };
    }
});
