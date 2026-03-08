import { useState, useEffect } from 'react';
import { Select, Group, Loader, Text } from '@mantine/core';

export interface Voice {
    name: string;
    languageCodes: string[];
    ssmlGender: string;
    provider?: string;
}

interface VoiceSelectorProps {
    value: Voice | null;
    onChange: (voice: Voice) => void;
    providerFilter?: 'gcp' | 'local';
}

export function VoiceSelector({ value, onChange, providerFilter }: VoiceSelectorProps) {
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

                // Set default if none selected and NO value was given (only if we want to force a selection)
                // However, SettingsModal has a `value={null}` state for new aliases temporarily. 
                // We should NOT auto-trigger `onChange` because it immediately saves over the mapping array!
                // If it's a completely empty value and it's NOT the settings modal forcing it empty:
                // Actually, it's safer to just let the user pick.
                // But for pure ViewerPage (if any), maybe it needs a default?
                // The issue: when SettingsModal renders an existing alias with a SAVED voice, `value` IS provided. 
                // But for a split second, or if the saved voice doesn't match the filtered `voices` list (different provider!), `value` is kept but we shouldn't overwrite it.
                // Let's remove the auto-onChange entirely. The initial `value` provided by props is fine.
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

    const filteredVoices = providerFilter
        ? voices.filter(v => v.provider === providerFilter)
        : voices;

    return (
        <Group>
            <Select
                placeholder="Select Voice"
                data={filteredVoices.map(v => ({
                    value: v.name,
                    label: `${v.name.split('/').pop()} (${v.provider === 'gcp' ? 'Google' : 'Local'}, ${v.ssmlGender})`
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
