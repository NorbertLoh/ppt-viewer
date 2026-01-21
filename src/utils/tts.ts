export const generateAudio = async (text: string): Promise<string> => {
    // Basic caching to avoid regenerating same audio
    if (generateAudio.cache.has(text)) {
        console.log('Using cached audio for:', text.substring(0, 20) + '...');
        return generateAudio.cache.get(text)!;
    }

    try {
        const params = new URLSearchParams();
        params.append('voice', 'en_UK/apope_low');
        params.append('text', text);
        params.append('ssml', 'true');

        console.log('Generating audio for:', text.substring(0, 20) + '...');
        const response = await fetch(`http://localhost:59125/api/tts?${params.toString()}`);

        if (!response.ok) {
            throw new Error(`TTS generation failed: ${response.status} ${response.statusText}`);
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);

        generateAudio.cache.set(text, url);
        return url;
    } catch (error) {
        console.error('Error generating audio:', error);
        throw error;
    }
};

// Add a cache property to the function
generateAudio.cache = new Map<string, string>();
