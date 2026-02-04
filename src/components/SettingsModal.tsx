import { Modal, Button, Text, Group, Stack, Code } from '@mantine/core';
import { useState, useEffect } from 'react';

interface SettingsModalProps {
    opened: boolean;
    onClose: () => void;
}

export function SettingsModal({ opened, onClose }: SettingsModalProps) {
    const [keyPath, setKeyPath] = useState<string | null>(null);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        if (opened) {
            checkKey();
        }
    }, [opened]);

    const checkKey = async () => {
        const path = await window.electronAPI.getGcpKeyPath();
        setKeyPath(path || null);
    };

    const handleSetKey = async () => {
        setError(null);
        try {
            const result = await window.electronAPI.setGcpKey();
            if (result.success && result.path) {
                setKeyPath(result.path);
            } else if (result.error) {
                setError(result.error);
            }
        } catch (err) {
            console.error(err);
            setError("Failed to set key");
        }
    };

    return (
        <Modal opened={opened} onClose={onClose} title="Settings" centered>
            <Stack>
                <Text fw={500}>Google Cloud TTS Configuration</Text>
                <Text size="sm" c="dimmed">
                    To use high-quality voices (Chirp 3 HD), you must provide a valid Google Cloud Service Account JSON key.
                </Text>

                <Group justify="space-between" align="center" p="xs" style={{ border: '1px solid var(--mantine-color-gray-8)', borderRadius: 4 }}>
                    <Text size="sm" fw={700}>Current Key:</Text>
                    {keyPath ? (
                        <Code color="green" style={{ maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            {keyPath}
                        </Code>
                    ) : (
                        <Text size="sm" c="red">Not Configured</Text>
                    )}
                </Group>

                {error && <Text c="red" size="sm">{error}</Text>}

                <Group justify="flex-end">
                    <Button onClick={handleSetKey} variant="light">
                        Select Key File...
                    </Button>
                    <Button onClick={onClose}>Close</Button>
                </Group>
            </Stack>
        </Modal>
    );
}
