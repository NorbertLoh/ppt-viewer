import { useState, useEffect } from 'react';
import { Select, Group, Loader, Text } from '@mantine/core';

export interface Voice {
    name: string;
    languageCodes: string[];
    ssmlGender: string;
}

interface VoiceSelectorProps {
    value: Voice | null;
    onChange: (voice: Voice) => void;
}

export function VoiceSelector({ value, onChange }: VoiceSelectorProps) {
    const [voices, setVoices] = useState<Voice[]>([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        const fetchVoices = async () => {
            if (!window.electronAPI) return;
            setLoading(true);
            try {
                const fetchedVoices: Voice[] = await window.electronAPI.getVoices();
                console.log('Fetched voices:', fetchedVoices);
                setVoices(fetchedVoices);

                // Set default if none selected or current selection is invalid
                if (fetchedVoices.length > 0) {
                    // Try to find a reasonable default if not provided
                    if (!value) {
                        // Prefer a Chirp3-HD voice if available (though backend should only return them now)
                        const defaultVoice = fetchedVoices.find(v => v.name.includes('Chirp3-HD')) || fetchedVoices[0];
                        onChange(defaultVoice);
                    }
                }
            } catch (err: any) {
                console.error("Error fetching voices:", err);
                setError("Failed to load voices");
            } finally {
                setLoading(false);
            }
        };

        fetchVoices();
    }, []); // Only fetch once on mount

    const handleChange = (selectedValue: string | null) => {
        const voice = voices.find(v => v.name === selectedValue);
        if (voice) {
            onChange(voice);
        }
    };

    if (loading) return <Loader size="xs" />;

    if (error) return <Text size="xs" c="red">{error}</Text>;

    return (
        <Group>
            <Select
                placeholder="Select Voice"
                data={voices.map(v => ({
                    value: v.name,
                    label: `${v.name.split('/').pop()} (${v.ssmlGender})`
                }))}
                value={value?.name || null}
                onChange={handleChange}
                searchable
                size="xs"
                w={250}
            />
        </Group>
    );
}
