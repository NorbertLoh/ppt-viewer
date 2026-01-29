"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const electron_1 = require("electron");
// When contextIsolation is false, we can attach directly to window
window.electronAPI = {
    convertPptx: (filePath) => electron_1.ipcRenderer.invoke('convert-pptx', filePath),
    onConversionUpdate: (callback) => electron_1.ipcRenderer.on('conversion-update', callback),
    getPathForFile: (file) => file.path,
    selectFile: () => electron_1.ipcRenderer.invoke('select-file'),
    saveAllNotes: (filePath, slides) => electron_1.ipcRenderer.invoke('save-all-notes', filePath, slides),
};
