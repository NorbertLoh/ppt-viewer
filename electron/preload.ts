import { contextBridge, ipcRenderer, webUtils } from 'electron';

// When contextIsolation is false, we can attach directly to window
(window as any).electronAPI = {
    convertPptx: (filePath: string) => ipcRenderer.invoke('convert-pptx', filePath),
    onConversionUpdate: (callback: (event: any, value: any) => void) => ipcRenderer.on('conversion-update', callback),
    getPathForFile: (file: File) => (file as any).path,
    selectFile: () => ipcRenderer.invoke('select-file'),
    saveAllNotes: (filePath: string, slides: any[]) => ipcRenderer.invoke('save-all-notes', filePath, slides),
};
