import { app, BrowserWindow, ipcMain } from 'electron';
import path from 'path';

// Handle creating/removing shortcuts on Windows when installing/uninstalling.
if (require('electron-squirrel-startup')) {
    app.quit();
}

const createWindow = () => {
    const mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: true,
            contextIsolation: false,
            webSecurity: false // Optional, but helps avoid some local file loading issues
        },
    });

    if (!app.isPackaged) {
        mainWindow.loadURL('http://localhost:5173');
        mainWindow.webContents.openDevTools();
    } else {
        mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
    }
};

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// IPC Handlers
ipcMain.handle('select-file', async () => {
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

ipcMain.handle('convert-pptx', async (event, filePath) => {
    console.log('Convert request for (raw):', filePath);

    // Resolve to absolute path to handle relative paths (e.g. from dev environment)
    const absolutePath = path.resolve(filePath);
    console.log('Convert request for (absolute):', absolutePath);

    const fs = require('fs');
    if (!fs.existsSync(absolutePath)) {
        return { success: false, error: `File not found: ${absolutePath}` };
    }

    const os = process.platform;
    const tempDir = app.getPath('temp');
    const outputDir = path.join(tempDir, 'ppt-viewer', path.basename(absolutePath, path.extname(absolutePath)));

    // Ensure parent dir exists
    if (!fs.existsSync(path.join(tempDir, 'ppt-viewer'))) {
        fs.mkdirSync(path.join(tempDir, 'ppt-viewer'));
    }

    console.log('Output Dir:', outputDir);

    try {
        if (os === 'win32') {
            const scriptPath = path.join(__dirname, '../electron/scripts/convert-win.ps1');
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

                child.stdout.on('data', (data: any) => {
                    console.log(`stdout: ${data}`);
                    stdout += data;
                });

                child.stderr.on('data', (data: any) => {
                    console.error(`stderr: ${data}`);
                    stderr += data;
                });

                child.on('close', async (code: number) => {
                    if (code === 0) {
                        // Read manifest
                        const manifestPath = path.join(outputDir, 'manifest.json');
                        try {
                            const fs = require('fs'); // Use sync fs for simplicity inside async callback, or keep promises
                            // Using readFileSync to be safe with blocking logic if preferred, or promises
                            const data = fs.readFileSync(manifestPath, 'utf8').replace(/^\uFEFF/, '');
                            const slides = JSON.parse(data);
                            // Fix image paths to be absolute or protocol based
                            const slidesWithPaths = slides.map((s: any) => ({
                                ...s,
                                src: `file://${path.join(outputDir, s.image)}`
                            }));
                            resolve({ success: true, slides: slidesWithPaths });
                        } catch (err) {
                            reject(err);
                        }
                    } else {
                        reject(new Error(`Conversion failed with code ${code}: ${stderr}`));
                    }
                });
            });
        }

        else if (os === 'darwin') {
            const scriptPath = path.join(__dirname, '../electron/scripts/convert-mac.applescript');
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

                child.stdout.on('data', (data: any) => { console.log('OSAScript:', data.toString()); stdout += data; });
                child.stderr.on('data', (data: any) => { console.error('OSAScript Err:', data.toString()); stderr += data; });

                child.on('close', (code: number) => {
                    if (code === 0) {
                        try {
                            const fs = require('fs');
                            console.log('--- Output Directory Contents ---');
                            // Simple recursive list for debugging
                            const listDir = (dir: string) => {
                                const files = fs.readdirSync(dir);
                                files.forEach((file: string) => {
                                    const p = path.join(dir, file);
                                    if (fs.statSync(p).isDirectory()) {
                                        console.log(`[DIR] ${p}`);
                                        listDir(p);
                                    } else {
                                        console.log(`[FILE] ${p}`);
                                    }
                                });
                            };
                            listDir(outputDir);
                            console.log('-------------------------------');
                        } catch (e) {
                            console.error('Error listing dir:', e);
                        }

                        // Read manifest
                        const manifestPath = path.join(outputDir, 'manifest.json');
                        try {
                            const fs = require('fs');
                            const data = fs.readFileSync(manifestPath, 'utf8').replace(/^\uFEFF/, '');
                            console.log('Manifest Content:', data); // LOGGING ADDED
                            const slides = JSON.parse(data);
                            const slidesWithPaths = slides.map((s: any) => ({
                                ...s,
                                src: s.image ? `file://${path.join(outputDir, s.image)}` : null
                            })).filter((s: any) => s.src !== null); // Filter out bad slides
                            resolve({ success: true, slides: slidesWithPaths });
                        } catch (err) {
                            reject(err);
                        }
                    } else {
                        reject(new Error(`Conversion failed with code ${code}: ${stderr || stdout}`));
                    }
                });
            });
        }

        // TODO: Implement other platforms
        return { success: false, error: 'Platform not supported yet' };


    } catch (err: any) {
        console.error('Conversion error:', err);
        return { success: false, error: err.message };
    }
});

ipcMain.handle('save-all-notes', async (event, filePath, slides) => {
    console.log('Save All Notes request for:', filePath);

    // Resolve absolute path
    const absolutePath = path.resolve(filePath);
    const fs = require('fs');

    if (!fs.existsSync(absolutePath)) {
        return { success: false, error: 'File not found' };
    }

    // Create temp JSON file for updates
    // We only need index and notes
    const updates = slides.map((s: any) => ({
        index: s.index,
        notes: s.notes
    }));

    const tempDir = app.getPath('temp');
    const updateJsonPath = path.join(tempDir, `ppt-update-${Date.now()}.json`);

    try {
        require('fs').writeFileSync(updateJsonPath, JSON.stringify(updates, null, 2), 'utf8');

        const os = process.platform;
        if (os === 'win32') {
            const scriptPath = path.join(__dirname, '../electron/scripts/save-all-notes-win.ps1');

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

                child.stdout.on('data', (data: any) => stdout += data);
                child.stderr.on('data', (data: any) => stderr += data);

                child.on('close', (code: number) => {
                    // Cleanup temp file
                    try { fs.unlinkSync(updateJsonPath); } catch (e) { }

                    if (code === 0) {
                        resolve({ success: true });
                    } else {
                        reject(new Error(`Save failed: ${stderr} | ${stdout}`));
                    }
                });
            });
        } else if (os === 'darwin') {
            const scriptPath = path.join(__dirname, '../electron/scripts/save-all-notes-mac.js');
            // Use 'osascript -l JavaScript' for JXA
            const { spawn } = require('child_process');
            const child = spawn('osascript', [
                '-l', 'JavaScript',
                scriptPath,
                absolutePath,
                updateJsonPath
            ]);

            return new Promise((resolve, reject) => {
                let stdout = '';
                let stderr = '';

                child.stdout.on('data', (data: any) => stdout += data);
                child.stderr.on('data', (data: any) => stderr += data);

                child.on('close', (code: number) => {
                    // Cleanup temp file
                    try { fs.unlinkSync(updateJsonPath); } catch (e) { }

                    if (code === 0) {
                        resolve({ success: true });
                    } else {
                        reject(new Error(`Save failed (Mac): ${stderr} | ${stdout}`));
                    }
                });
            });
        }

        return { success: false, error: 'Save not supported on this platform yet' };


    } catch (e: any) {
        return { success: false, error: e.message };
    }
});

ipcMain.handle('get-video-save-path', async () => {
    const { dialog } = require('electron');
    const win = require('electron').BrowserWindow.getFocusedWindow();
    const app = require('electron').app;
    const path = require('path');

    const result = await dialog.showSaveDialog(win!, {
        title: 'Save Video As',
        defaultPath: path.join(app.getPath('documents'), 'Output.mp4'),
        filters: [{ name: 'MPEG-4 Video', extensions: ['mp4'] }]
    });

    if (result.canceled || !result.filePath) {
        return null;
    }
    return result.filePath;
});

ipcMain.handle('generate-video', async (event, { filePath, slidesAudio, videoOutputPath }) => {
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
    } catch (e) {
        console.error("Failed to create Office container dir:", e);
        return { success: false, error: "Could not create audio directory in Office container. Check permissions." };
    }

    try {
        for (const slide of slidesAudio) {
            console.log(`Processing slide ${slide.index}`);
            const buffer = Buffer.from(slide.audioData);
            const audioFileName = `audio_${slide.index}.mp3`;
            const audioFilePath = path.join(audioSessionDir, audioFileName);

            fs.writeFileSync(audioFilePath, buffer);
            console.log(`Saved audio to ${audioFilePath}`);

            if (process.platform === 'darwin') {
                const scriptPath = path.join(__dirname, '../electron/scripts/trigger-macro.applescript');

                // AppleScript arguments: audioPath, slideIndex, presentationPath
                const child = spawn('osascript', [
                    scriptPath,
                    audioFilePath,
                    slide.index.toString(),
                    filePath
                ]);

                await new Promise<void>((resolve, reject) => {
                    let stdout = '';
                    let stderr = '';
                    child.stdout.on('data', (d: any) => stdout += d);
                    child.stderr.on('data', (d: any) => stderr += d);
                    child.on('close', (code: number) => {
                        if (code === 0 && !stdout.includes("Error")) {
                            console.log(`Macro triggered for slide ${slide.index}`);
                            resolve();
                        } else {
                            console.error(`Failed for slide ${slide.index}: ${stderr} ${stdout}`);
                            reject(new Error(stdout || stderr));
                        }
                    });
                });
            } else {
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

            await new Promise<void>((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (d: any) => stdout += d);
                child.stderr.on('data', (d: any) => stderr += d);
                child.on('close', (code: number) => {
                    // Note: 'save as movie' might return 0 but export continues in PPT background.
                    if (code === 0 && !stdout.includes("Error")) {
                        console.log(`Video export initiated: ${videoOutputPath}`);
                        resolve();
                    } else {
                        console.error(`Export failed: ${stderr || stdout}`);
                        reject(new Error(stdout || stderr));
                    }
                });
            });

            return { success: true, outputPath: videoOutputPath };
        } else {
            return { success: false, error: "Windows video export not implemented" };
        }

    } catch (e: any) {
        console.error('Generation failed:', e);
        return { success: false, error: e.message };
    }
});
