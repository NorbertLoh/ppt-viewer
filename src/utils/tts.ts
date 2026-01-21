const getAudioBlob = async (text: string): Promise<Blob> => {
    // Generate URL params for metadata
    const urlParams = new URLSearchParams();
    urlParams.append('voice', 'en_UK/apope_low');
    urlParams.append('ssml', 'true');

    // Clean text: remove invalid XML control characters (like Vertical Tab \x0b)
    // Keep CR/LF (\n, \r) and Tab (\t)
    const sanitizedText = text.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');

    console.log('Generating audio blob for:', sanitizedText.substring(0, 20) + '...');

    // Server wraps in <speak> automatically, so we just send sanitized text
    const bodyText = sanitizedText;

    // Use POST to avoid URL length limits
    const response = await fetch(`http://localhost:59125/api/tts?${urlParams.toString()}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'text/plain' // or application/ssml+xml
        },
        body: bodyText
    });

    if (!response.ok) {
        throw new Error(`TTS generation failed: ${response.status} ${response.statusText}`);
    }

    return await response.blob();
};

export const generateAudio = async (text: string): Promise<string> => {
    // Basic caching to avoid regenerating same audio
    if (generateAudio.cache.has(text)) {
        console.log('Using cached audio for:', text.substring(0, 20) + '...');
        return generateAudio.cache.get(text)!;
    }

    try {
        const blob = await getAudioBlob(text);
        const url = URL.createObjectURL(blob);

        generateAudio.cache.set(text, url);
        return url;
    } catch (error) {
        console.error('Error generating audio:', error);
        throw error;
    }
};

export const getAudioBuffer = async (text: string): Promise<ArrayBuffer> => {
    // Check if we have a URL cached. If so, we can fetch it back to get blob (fast local fetch)
    // Or just re-fetch from TTS.
    // If we have the Blob in memory... we don't. We only have the URL.
    // So we fetch the URL.
    if (generateAudio.cache.has(text)) {
        const url = generateAudio.cache.get(text)!;
        const res = await fetch(url);
        return await res.arrayBuffer();
    }

    // Not cached, generate new
    const blob = await getAudioBlob(text);
    // Might as well cache it while we're here
    const url = URL.createObjectURL(blob);
    generateAudio.cache.set(text, url);

    return await blob.arrayBuffer();
};

// Add a cache property to the function
generateAudio.cache = new Map<string, string>();
