import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';
import { TextToSpeechClient } from '@google-cloud/text-to-speech';
import dotenv from 'dotenv';

// Load .env
dotenv.config();

import Store from 'electron-store';
const store = new Store();

// --- Helper to get GCP Key Path ---
function getGcpKeyPath(): string | undefined {
    // Priority: 1. ENV var (dev/runtime override), 2. Stored path
    return process.env.GOOGLE_APPLICATION_CREDENTIALS || store.get('gcpKeyPath') as string;
}

function getTtsProvider(): string {
    // Priority: 1. ENV var, 2. Stored Key -> implies GCP, 3. Default to GCP (so we prompt for key)
    if (process.env.TTS_PROVIDER) return process.env.TTS_PROVIDER;
    if (getGcpKeyPath()) return 'gcp';
    return 'gcp'; // Default to GCP instead of local, so we hit the "missing key" check
}

function resolveScriptPath(scriptName: string): string {
    if (app.isPackaged) {
        // In production, scripts are unpacked to Resources/electron/scripts
        return path.join(process.resourcesPath, 'electron', 'scripts', scriptName);
    } else {
        // In dev, scripts are in electron/scripts relative to main.ts
        return path.join(__dirname, '../electron/scripts', scriptName);
    }
}

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
        // Optional: Open DevTools on specific key combination for debugging production builds
        mainWindow.webContents.on('before-input-event', (event, input) => {
            if (input.control && input.shift && input.key.toLowerCase() === 'i') {
                mainWindow.webContents.toggleDevTools();
                event.preventDefault();
            }
        });
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
            const scriptPath = resolveScriptPath('convert-win.ps1');
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
                                src: `file://${path.join(outputDir, s.image)}`,
                                notes: s.notes ? s.notes.replace(/\\n/g, '\n') : ''
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
            const scriptPath = resolveScriptPath('convert-mac.applescript');
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
                            console.log('Manifest Content Sample (First 100 chars):', data.substring(0, 100));
                            // Debug encoding
                            const rawBuffer = fs.readFileSync(manifestPath);
                            console.log('Manifest Start Hex:', rawBuffer.subarray(0, 20).toString('hex'));
                            const slides = JSON.parse(data);
                            const slidesWithPaths = slides.map((s: any) => ({
                                ...s,
                                src: s.image ? `file://${path.join(outputDir, s.image)}` : null,
                                // Fix escaped newlines from AppleScript/Perl pipeline
                                notes: s.notes ? s.notes.replace(/\\n/g, '\n') : ''
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

// --- Helper for Audio Insertion ---
async function handleAudioInsertion(filePath: string, slidesAudio: any[]) {
    console.log(`handleAudioInsertion for ${slidesAudio.length} slides`);
    const path = require('path');
    const fs = require('fs');
    const { spawn } = require('child_process');
    const app = require('electron').app;

    if (!slidesAudio || slidesAudio.length === 0) return { success: true };

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
        let batchParams = "";

        // 1. Save all audio files and prepare batch params
        for (const slide of slidesAudio) {
            console.log(`Processing slide ${slide.index}`);
            // audioData might be coming as an object from IPC, need to ensure it's a buffer
            const buffer = Buffer.from(slide.audioData);
            const audioFileName = `audio_${slide.index}.mp3`;
            const audioFilePath = path.join(audioSessionDir, audioFileName);

            fs.writeFileSync(audioFilePath, buffer);
            console.log(`Saved audio to ${audioFilePath}`);

            // Append to batch params: Index|AudioPath|PresentationPath
            batchParams += `${slide.index}|${audioFilePath}|${filePath}\n`;
        }

        if (process.platform === 'darwin') {
            const scriptPath = resolveScriptPath('trigger-macro.applescript');

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

            await new Promise<void>((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (d: any) => stdout += d);
                child.stderr.on('data', (d: any) => stderr += d);
                child.on('close', (code: number) => {
                    if (code === 0 && !stdout.includes("Error")) {
                        console.log(`Batch audio macro completed successfully.`);
                        resolve();
                    } else {
                        console.error(`Failed to run batch audio macro: ${stderr} ${stdout}`);
                        reject(new Error(stdout || stderr));
                    }
                });
            });
            return { success: true };
        } else {
            return { success: false, error: "Windows audio insertion not implemented" };
        }
    } catch (e: any) {
        console.error('Audio insertion failed:', e);
        return { success: false, error: e.message };
    }
}

ipcMain.handle('save-all-notes', async (event, filePath, slides, slidesAudio) => {
    console.log('Save All Notes request for:', filePath);

    // Resolve absolute path
    const absolutePath = path.resolve(filePath);
    const fs = require('fs');

    if (!fs.existsSync(absolutePath)) {
        return { success: false, error: 'File not found' };
    }

    try {
        // 1. Insert Audio (if provided)
        if (slidesAudio && slidesAudio.length > 0) {
            console.log('Inserting audio before saving notes...');
            const audioResult = await handleAudioInsertion(absolutePath, slidesAudio);
            if (!audioResult!.success) {
                console.error("Audio insertion failed during save:", audioResult!.error);
                return { success: false, error: "Audio insertion failed: " + audioResult!.error };
            }
        }

        if (process.platform === 'darwin') {
            const app = require('electron').app;
            const homeDir = app.getPath('home');
            const officeContainer = path.join(homeDir, 'Library/Group Containers/UBF8T346G9.Office');

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

            const dataPath = path.join(officeContainer, `notes_data_${Date.now()}.txt`);
            fs.writeFileSync(dataPath, dataContent, 'utf8');

            // 2. Prepare Params File
            const paramsPath = path.join(officeContainer, 'notes_params.txt');
            // Content: PresentationPath|DataPath
            const paramsContent = `${absolutePath}|${dataPath}`;
            fs.writeFileSync(paramsPath, paramsContent, 'utf8');

            // 3. Trigger Macro
            const scriptPath = resolveScriptPath('trigger-macro.applescript');

            // Args: macroName, pptPath
            const { spawn } = require('child_process');
            const child = spawn('osascript', [
                scriptPath,
                "UpdateNotes",
                absolutePath
            ]);

            await new Promise<void>((resolve, reject) => {
                let stdout = '';
                let stderr = '';
                child.stdout.on('data', (d: any) => stdout += d);
                child.stderr.on('data', (d: any) => stderr += d);
                child.on('close', (code: number) => {
                    // Cleanup data file
                    try { fs.unlinkSync(dataPath); } catch (e) { }

                    if (code === 0 && !stdout.includes("Error")) {
                        console.log(`UpdateNotes macro triggered.`);
                        resolve();
                    } else {
                        console.error(`Failed to trigger UpdateNotes: ${stderr} ${stdout}`);
                        reject(new Error(stdout || stderr));
                    }
                });
            });

            return { success: true };

        } else {
            // Windows implementation...
            return { success: false, error: 'Save not supported on this platform yet' };
        }

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
        if (slidesAudio.length > 0) {
            const audioResult = await handleAudioInsertion(filePath, slidesAudio);
            if (!audioResult!.success) {
                return { success: false, error: audioResult!.error };
            }
        }

        if (process.platform === 'darwin') {
            const exportScriptPath = resolveScriptPath('export-to-video.applescript');

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

// --- TTS Handler ---

// --- Settings Handler ---
ipcMain.handle('get-gcp-key-path', async () => {
    return store.get('gcpKeyPath');
});

ipcMain.handle('set-gcp-key', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'JSON', extensions: ['json'] }]
    });

    if (canceled || filePaths.length === 0) {
        return { success: false };
    }

    const keyPath = filePaths[0];

    // Basic validation
    try {
        const fs = require('fs');
        const content = JSON.parse(fs.readFileSync(keyPath, 'utf8'));
        if (!content.type || content.type !== 'service_account') {
            return { success: false, error: 'Invalid Service Account Key JSON' };
        }
    } catch (err) {
        return { success: false, error: 'Invalid JSON file' };
    }

    store.set('gcpKeyPath', keyPath);
    return { success: true, path: keyPath };
});

// --- TTS Handler ---
ipcMain.handle('get-voices', async () => {
    const provider = getTtsProvider();
    console.log(`Get Voices Request. Provider: ${provider}`);
    const keyPath = getGcpKeyPath();

    try {
        if (provider === 'gcp') {
            if (!keyPath) {
                console.warn("TTS_PROVIDER is 'gcp' but GOOGLE_APPLICATION_CREDENTIALS is missing.");
                return [];
            }

            // Explicitly pass credentials if using stored path
            const options: any = {};
            if (keyPath) {
                options.keyFilename = keyPath;
            }

            const client = new TextToSpeechClient(options);
            const [result] = await client.listVoices({ languageCode: 'en-US' });
            // Filter for Chirp 3 HD voices as requested
            const voices = result.voices || [];
            return voices.filter(v => v.name && v.name.includes('Chirp3-HD'));
        } else {
            // Local fallback
            return [
                { name: 'en_UK/apope_low', ssmlGender: 'MALE', languageCodes: ['en-GB'] },
                { name: 'default', ssmlGender: 'NEUTRAL', languageCodes: ['en-US'] }
            ];
        }
    } catch (error) {
        console.error("Failed to list voices:", error);
        return [];
    }
});

ipcMain.handle('generate-speech', async (event, { text, voiceOption }) => {
    // Determine provider: 'gcp' or 'local' (default)
    const provider = getTtsProvider();

    console.log(`TTS Request: "${text.substring(0, 20)}..." using provider: ${provider}, voice: ${voiceOption ? voiceOption.name : 'default'}`);

    try {
        if (provider === 'gcp') {
            // --- Google Cloud TTS ---
            const keyPath = getGcpKeyPath();
            if (!keyPath) {
                throw new Error("TTS_PROVIDER is 'gcp' but GOOGLE_APPLICATION_CREDENTIALS is not set");
            }

            const options: any = {};
            if (keyPath) {
                options.keyFilename = keyPath;
            }

            const client = new TextToSpeechClient(options);

            // Detect SSML (basic check for tags)
            const isSsml = /<[^>]+>/.test(text);

            let input: any;
            if (isSsml) {
                // Ensure it's wrapped in <speak>
                let ssmlText = text;
                if (!ssmlText.trim().startsWith('<speak>')) {
                    ssmlText = `<speak>${ssmlText}</speak>`;
                }
                input = { ssml: ssmlText };
                console.log('Sending SSML request to GCP:', ssmlText);
            } else {
                input = { text: text };
            }

            const request: any = {
                input: input,
                // Use selected voice or default
                voice: voiceOption ? { languageCode: voiceOption.languageCodes[0], name: voiceOption.name } : { languageCode: 'en-US', name: 'en-US-Journey-F' },
                audioConfig: { audioEncoding: 'MP3' },
            };

            const [response] = await client.synthesizeSpeech(request);
            return response.audioContent; // This is Uint8Array or Buffer

        } else {
            // --- Local TTS (Dev) ---
            // Default to local server
            const localUrl = process.env.LOCAL_TTS_URL || 'http://localhost:59125/api/tts';
            // Use selected voice name if available, else env default, else hardcoded default
            const voice = (voiceOption && voiceOption.name) || process.env.LOCAL_TTS_VOICE || 'en_UK/apope_low';

            // Construct URL with params
            const url = new URL(localUrl);
            url.searchParams.append('voice', voice);
            url.searchParams.append('ssml', 'true');

            // Clean text (basic)
            const sanitizedText = text.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');

            const resp = await fetch(url.toString(), {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: sanitizedText
            });

            if (!resp.ok) {
                throw new Error(`Local TTS failed: ${resp.status} ${resp.statusText}`);
            }

            const arrayBuffer = await resp.arrayBuffer();
            return new Uint8Array(arrayBuffer);
        }
    } catch (error: any) {
        console.error("TTS generation failed:", error);
        throw new Error(error.message || "Unknown TTS error");
    }
});
