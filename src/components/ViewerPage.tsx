import { Box, Group, Button, Image, ScrollArea, Textarea, Title } from '@mantine/core';
import { useState, useEffect } from 'react';
import type { Slide } from '../electron';

interface ViewerPageProps {
    slides: Slide[];
    onSave: (updatedSlides: Slide[]) => void;
    onBack: () => void;
}

export function ViewerPage({ slides: initialSlides, onSave, onBack }: ViewerPageProps) {
    const [slides, setSlides] = useState<Slide[]>(initialSlides);
    const [activeSlideIndex, setActiveSlideIndex] = useState(0);

    const activeSlide = slides[activeSlideIndex] || { src: '', notes: '' };

    // Reset index when initialSlides change (e.g. new file loaded)
    useEffect(() => {
        setSlides(initialSlides);
        setActiveSlideIndex(0);
    }, [initialSlides]);

    const handleNotesChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
        const newText = e.target.value;
        const newSlides = [...slides];
        newSlides[activeSlideIndex] = { ...newSlides[activeSlideIndex], notes: newText };
        setSlides(newSlides);
    };

    return (
        <div style={{ height: '100vh', width: '100vw', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
            {/* Header / Toolbar */}
            <Group justify="space-between" p="xs" style={{ borderBottom: '1px solid var(--mantine-color-dark-4)', background: 'var(--mantine-color-dark-7)' }}>
                <Button variant="subtle" size="xs" onClick={onBack}>&larr; Back</Button>
                <Title order={5}>Viewer</Title>
                <Button variant="filled" color="blue" size="xs" onClick={() => onSave(slides)}>Save All Changes</Button>
            </Group>

            <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
                {/* Left Panel: Thumbnails */}
                <div style={{ width: '250px', height: '100%', borderRight: '1px solid var(--mantine-color-dark-4)', display: 'flex', flexDirection: 'column' }}>
                    <ScrollArea style={{ flex: 1 }} type="auto">
                        <Box p="md">
                            {slides.map((slide, index) => (
                                <Box
                                    key={slide.index}
                                    onClick={() => setActiveSlideIndex(index)}
                                    style={{
                                        marginBottom: '1rem',
                                        cursor: 'pointer',
                                        border: activeSlideIndex === index ? '2px solid var(--mantine-color-blue-6)' : '2px solid transparent',
                                        borderRadius: '4px'
                                    }}
                                >
                                    <Image src={slide.src} radius="sm" />
                                </Box>
                            ))}
                        </Box>
                    </ScrollArea>
                </div>

                {/* Right Panel */}
                <div style={{ flex: 1, height: '100%', display: 'flex', flexDirection: 'column', backgroundColor: 'var(--mantine-color-body)' }}>
                    {/* Top: Slide View */}
                    <Box style={{ flex: 2, position: 'relative', borderBottom: '1px solid var(--mantine-color-dark-4)', padding: '1rem', background: 'var(--mantine-color-dark-8)', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                        <Image
                            src={activeSlide.src}
                            fit="contain"
                            style={{ maxHeight: '100%', maxWidth: '100%' }}
                        />
                    </Box>

                    {/* Bottom: Notes */}
                    <Box style={{ flex: 1, padding: '1rem', display: 'flex', flexDirection: 'column' }}>
                        <Textarea
                            label="Presenter Notes"
                            value={activeSlide.notes}
                            onChange={handleNotesChange}
                            minRows={4}
                            maxRows={10}
                            style={{ flex: 1, display: 'flex', flexDirection: 'column' }}
                            styles={{
                                wrapper: { flex: 1, display: 'flex', flexDirection: 'column' },
                                input: { flex: 1 }
                            }}
                        />
                    </Box>
                </div>
            </div>
        </div>
    );
}
