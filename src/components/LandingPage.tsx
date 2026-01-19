import { Group, Text, Button } from '@mantine/core';
import { Dropzone, type DropzoneProps, IMAGE_MIME_TYPE } from '@mantine/dropzone';

interface LandingPageProps extends Partial<DropzoneProps> {
    onSelectFile?: () => void;
}

export function LandingPage({ onSelectFile, ...props }: LandingPageProps) {
    return (
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
            <Dropzone
                onDrop={(files) => console.log('accepted files', files)}
                onReject={(files) => console.log('rejected files', files)}
                maxSize={50 * 1024 ** 2}
                accept={['application/vnd.openxmlformats-officedocument.presentationml.presentation', ...IMAGE_MIME_TYPE]} // Allowing images for testing too
                {...props}
                style={{
                    flex: 1,
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    background: 'var(--mantine-color-body)',
                    borderRadius: 0,
                    border: 'none',
                }}
            >
                <Group justify="center" gap="xl" style={{ pointerEvents: 'none' }}>
                    <div>
                        <Text size="xl" inline c="dimmed">
                            Drag a PowerPoint file here
                        </Text>
                    </div>
                </Group>
            </Dropzone>
            <Group justify="center" p="md" style={{ background: 'var(--mantine-color-dark-8)' }}>
                <Button onClick={onSelectFile} variant="light">
                    Or Select File Manually
                </Button>
            </Group>
        </div>
    );
}
