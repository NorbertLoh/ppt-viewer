const { ipcRenderer } = (window as any).require('electron');
export interface VoiceOption {
    name: string;
    languageCodes: string[];
    ssmlGender: string;
}

export const generateAudio = async (text: string, voiceOption?: VoiceOption): Promise<string> => {
    const key = text + (voiceOption ? `_${voiceOption.name}` : '_default');

    // Basic caching to avoid regenerating same audio
    if (generateAudio.cache.has(key)) {
        console.log('Using cached audio for:', text.substring(0, 20) + '...', 'Voice:', voiceOption?.name);
        return generateAudio.cache.get(key)!;
    }

    try {
        // 1. Get Audio Buffer from Backend (IPC)
        // The backend handles GCP vs Local logic
        const buffer = await getAudioBuffer(text, voiceOption);

        // 2. Convert to Blob/URL for frontend playback
        const blob = new Blob([buffer as any], { type: 'audio/mp3' });
        const url = URL.createObjectURL(blob);

        generateAudio.cache.set(key, url);
        return url;
    } catch (error) {
        console.error('Error generating audio:', error);
        throw error;
    }
};

export const getAudioBuffer = async (text: string, voiceOption?: VoiceOption): Promise<ArrayBuffer> => {
    const key = text + (voiceOption ? `_${voiceOption.name}` : '_default');

    // Check cache first (if we have URL, fetch blob locally)
    if (generateAudio.cache.has(key)) {
        const url = generateAudio.cache.get(key)!;
        const res = await fetch(url);
        return await res.arrayBuffer();
    }

    // Call Backend
    const audioData: Uint8Array = await ipcRenderer.invoke('generate-speech', { text, voiceOption });

    console.log('Got audio data length:', audioData.byteLength);

    // Cache the result as URL for playback consistency
    // Note: This logic is slightly redundant with generateAudio but useful if getAudioBuffer is called directly
    const blob = new Blob([audioData as any], { type: 'audio/mp3' });
    const url = URL.createObjectURL(blob);
    generateAudio.cache.set(key, url);

    return audioData.buffer as ArrayBuffer;
};

// Add a cache property to the function
generateAudio.cache = new Map<string, string>();
