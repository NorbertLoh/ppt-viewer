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
        }

        return { success: false, error: 'Save not supported on this platform yet' };

    } catch (e: any) {
        return { success: false, error: e.message };
    }
});
