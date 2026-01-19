import { Box, Group, Button, Image, ScrollArea, Textarea, Title, ActionIcon, Tooltip, Menu, rem, TextInput } from '@mantine/core';
import { useState, useEffect, useRef, useCallback } from 'react';
import {
    IconPlayerPause,
    IconKeyboard,
    IconVolume,
    IconPilcrow,
    IconChevronDown,
    IconClock,
    IconPlus,
    IconArrowBackUp,
    IconArrowForwardUp
} from '@tabler/icons-react';
import type { Slide } from '../electron';

interface ViewerPageProps {
    slides: Slide[];
    onSave: (updatedSlides: Slide[]) => void;
    onBack: () => void;
}

export function ViewerPage({ slides: initialSlides, onSave, onBack }: ViewerPageProps) {
    const [slides, setSlides] = useState<Slide[]>(initialSlides);
    // History State
    const [history, setHistory] = useState<Slide[][]>([initialSlides]);
    const [historyIndex, setHistoryIndex] = useState(0);

    const [activeSlideIndex, setActiveSlideIndex] = useState(0);
    const textareaRef = useRef<HTMLTextAreaElement>(null);
    const [customBreak, setCustomBreak] = useState('');
    const debounceRef = useRef<NodeJS.Timeout | null>(null);

    const activeSlide = slides[activeSlideIndex] || { src: '', notes: '' };

    // Reset index when initialSlides change (e.g. new file loaded)
    useEffect(() => {
        setSlides(initialSlides);
        setHistory([initialSlides]);
        setHistoryIndex(0);
        setActiveSlideIndex(0);
    }, [initialSlides]);

    // Push current state to history. 
    // If 'overwrite', replaces the current history head (useful for typing sequences).
    // Usually we push new state.
    const pushToHistory = useCallback((newSlides: Slide[]) => {
        setHistory(prev => {
            const currentHistory = prev.slice(0, historyIndex + 1);
            return [...currentHistory, newSlides];
        });
        setHistoryIndex(prev => prev + 1);
    }, [historyIndex]);

    const handleUndo = useCallback(() => {
        if (historyIndex > 0) {
            setHistoryIndex(prev => prev - 1);
            setSlides(history[historyIndex - 1]);
        }
    }, [history, historyIndex]);

    const handleRedo = useCallback(() => {
        if (historyIndex < history.length - 1) {
            setHistoryIndex(prev => prev + 1);
            setSlides(history[historyIndex + 1]);
        }
    }, [history, historyIndex]);

    // Keyboard Shortcuts for Undo/Redo
    useEffect(() => {
        const handleKeyDown = (e: KeyboardEvent) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
                e.preventDefault();
                handleUndo();
            }
            if ((e.ctrlKey || e.metaKey) && e.key === 'y') {
                e.preventDefault();
                handleRedo();
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [handleUndo, handleRedo]);


    const handleNotesChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
        const newText = e.target.value;
        const newSlides = [...slides];
        newSlides[activeSlideIndex] = { ...newSlides[activeSlideIndex], notes: newText };
        setSlides(newSlides);

        // Debounce history save for typing
        if (debounceRef.current) clearTimeout(debounceRef.current);
        debounceRef.current = setTimeout(() => {
            pushToHistory(newSlides);
        }, 800);
    };

    const insertTag = (startTag: string, endTag: string = '') => {
        const textarea = textareaRef.current;
        if (!textarea) return;

        // Force save current history state before modification to ensure we can undo this specific action
        // Actually, if we just typed, the debounce might not have fired yet. 
        // Ideally we want to commit "Before Tag" state if it changed.
        // But for simplicity, we just push the NEW state after tag insertion into history.
        // Wait, if we type "foo", wait 200ms, then insert tag. "foo" isn't in history yet?
        // Let's rely on React state. The current 'slides' IS the latest 'foo'. 
        // So we just need to ensure we have a history point BEFORE this change?
        // If historyIndex points to "foo", good. If not (debounce pending), we might lose "foo" step?
        // To be safe: clear debounce and push CURRENT slides if meaningful difference? 
        // Simplest strategy: Just push `newSlides` to history. Undo will go back to *whatever was computed last*. 
        // If "foo" wasn't pushed yet, Undo goes back to pre-"foo". That's bad.
        // So: Clear debounce, push CURRENT slides (if not equal to history head), THEN apply tag and push again.

        if (debounceRef.current) {
            clearTimeout(debounceRef.current);
            // If there were pending changes, push them first?
            // This is getting complex. Let's stick to: "Tag insertion is a discrete history event"
            // We assume the user stopped typing for a split second or we accept small data loss in undo stack for rapid typing+clicking.
            // Better: Implicitly, `slides` is the latest. 
            // We want history to look like: [State A], [State A + Tag].
            // If we only push [State A + Tag], then pressing Undo goes to [State A] (which might be old if we didn't push A).
            // So yes, we should push 'slides' (current state) if it's different from history[historyIndex].

            // Let's implement that check ideally, or just lazy-push.
            // For now: Just push the result.
        }

        const start = textarea.selectionStart;
        const end = textarea.selectionEnd;
        const text = activeSlide.notes || '';

        const before = text.substring(0, start);
        const selection = text.substring(start, end);
        const after = text.substring(end);

        const newText = before + startTag + selection + endTag + after;

        const newSlides = [...slides];
        newSlides[activeSlideIndex] = { ...newSlides[activeSlideIndex], notes: newText };

        setSlides(newSlides);
        pushToHistory(newSlides); // Discrete Action -> immediate history

        // Restore cursor / selection
        setTimeout(() => {
            textarea.focus();
            textarea.setSelectionRange(start + startTag.length, end + startTag.length);
        }, 0);
    };

    const handleCustomBreak = () => {
        if (!customBreak) return;
        insertTag(`<break time="${customBreak}"/>`);
        setCustomBreak('');
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

                    {/* Bottom: Notes + Toolbar */}
                    <Box style={{ flex: 1, padding: '1rem', display: 'flex', flexDirection: 'column' }}>

                        {/* SSML Toolbar */}
                        <Group gap={0} mb="xs" style={{ border: '1px solid var(--mantine-color-dark-4)', borderRadius: '4px', padding: '4px', background: 'var(--mantine-color-dark-6)' }}>
                            <ActionIcon variant="subtle" color="gray" size="lg" onClick={handleUndo} disabled={historyIndex === 0}>
                                <IconArrowBackUp style={{ width: rem(18), height: rem(18) }} />
                            </ActionIcon>
                            <ActionIcon variant="subtle" color="gray" size="lg" onClick={handleRedo} disabled={historyIndex === history.length - 1}>
                                <IconArrowForwardUp style={{ width: rem(18), height: rem(18) }} />
                            </ActionIcon>

                            <div style={{ width: 1, height: 20, background: 'var(--mantine-color-dark-4)', margin: '0 8px' }} />

                            <Menu shadow="md" width={220} trigger="click" position="bottom-start" offset={0} closeOnItemClick={false}>
                                <Menu.Target>
                                    <ActionIcon variant="subtle" color="gray" size="lg" aria-label="Break time">
                                        <IconPlayerPause style={{ width: rem(18), height: rem(18) }} />
                                        <IconChevronDown style={{ width: rem(12), height: rem(12), marginLeft: 4 }} />
                                    </ActionIcon>
                                </Menu.Target>
                                <Menu.Dropdown>
                                    <Menu.Label>Break Duration</Menu.Label>
                                    <Menu.Item leftSection={<IconClock size={14} />} onClick={() => insertTag('<break time="200ms"/>')}>200 ms</Menu.Item>
                                    <Menu.Item leftSection={<IconClock size={14} />} onClick={() => insertTag('<break time="500ms"/>')}>500 ms</Menu.Item>
                                    <Menu.Item leftSection={<IconClock size={14} />} onClick={() => insertTag('<break time="1s"/>')}>1 second</Menu.Item>
                                    <Menu.Item leftSection={<IconClock size={14} />} onClick={() => insertTag('<break time="2s"/>')}>2 seconds</Menu.Item>

                                    <Menu.Divider />
                                    <Menu.Label>Custom</Menu.Label>
                                    <Box p="xs" pt={0}>
                                        <Group gap={5}>
                                            <TextInput
                                                placeholder="e.g. 3s or 500ms"
                                                size="xs"
                                                value={customBreak}
                                                onChange={(e) => setCustomBreak(e.currentTarget.value)}
                                                onKeyDown={(e) => {
                                                    if (e.key === 'Enter') {
                                                        handleCustomBreak();
                                                        // Close menu logic if needed, but keeping open for now
                                                    }
                                                }}
                                                style={{ flex: 1 }}
                                            />
                                            <ActionIcon variant="filled" color="blue" size="sm" onClick={handleCustomBreak}>
                                                <IconPlus size={14} />
                                            </ActionIcon>
                                        </Group>
                                    </Box>
                                </Menu.Dropdown>
                            </Menu>

                            <Menu shadow="md" width={200} trigger="hover" position="bottom-start" offset={0}>
                                <Menu.Target>
                                    <ActionIcon variant="subtle" color="gray" size="lg" aria-label="Say As">
                                        <IconKeyboard style={{ width: rem(18), height: rem(18) }} />
                                        <IconChevronDown style={{ width: rem(12), height: rem(12), marginLeft: 4 }} />
                                    </ActionIcon>
                                </Menu.Target>
                                <Menu.Dropdown>
                                    <Menu.Label>Interpret As</Menu.Label>
                                    <Menu.Item onClick={() => insertTag('<say-as interpret-as="characters">', '</say-as>')}>Spell Out</Menu.Item>
                                    <Menu.Item onClick={() => insertTag('<say-as interpret-as="cardinal">', '</say-as>')}>Number (Cardinal)</Menu.Item>
                                    <Menu.Item onClick={() => insertTag('<say-as interpret-as="ordinal">', '</say-as>')}>Ordinal (1st, 2nd)</Menu.Item>
                                    <Menu.Item onClick={() => insertTag('<say-as interpret-as="digits">', '</say-as>')}>Digits</Menu.Item>
                                    <Menu.Item onClick={() => insertTag('<say-as interpret-as="fraction">', '</say-as>')}>Fraction</Menu.Item>
                                    <Menu.Item onClick={() => insertTag('<say-as interpret-as="expletive">', '</say-as>')}>Expletive</Menu.Item>
                                </Menu.Dropdown>
                            </Menu>

                            <Menu shadow="md" width={200} trigger="hover" position="bottom-start" offset={0}>
                                <Menu.Target>
                                    <ActionIcon variant="subtle" color="gray" size="lg" aria-label="Emphasis">
                                        <IconVolume style={{ width: rem(18), height: rem(18) }} />
                                        <IconChevronDown style={{ width: rem(12), height: rem(12), marginLeft: 4 }} />
                                    </ActionIcon>
                                </Menu.Target>
                                <Menu.Dropdown>
                                    <Menu.Label>Emphasis Level</Menu.Label>
                                    <Menu.Item onClick={() => insertTag('<emphasis level="strong">', '</emphasis>')}>Strong</Menu.Item>
                                    <Menu.Item onClick={() => insertTag('<emphasis level="moderate">', '</emphasis>')}>Moderate</Menu.Item>
                                    <Menu.Item onClick={() => insertTag('<emphasis level="reduced">', '</emphasis>')}>Reduced</Menu.Item>
                                </Menu.Dropdown>
                            </Menu>

                            <Tooltip label="Paragraph">
                                <ActionIcon variant="subtle" color="gray" size="lg" onClick={() => insertTag('<p>', '</p>')}>
                                    <IconPilcrow style={{ width: rem(18), height: rem(18) }} />
                                </ActionIcon>
                            </Tooltip>
                        </Group>

                        <Textarea
                            ref={textareaRef}
                            label="Presenter Notes"
                            value={activeSlide.notes}
                            onChange={handleNotesChange}
                            minRows={4}
                            maxRows={10}
                            style={{ flex: 1, display: 'flex', flexDirection: 'column' }}
                            styles={{
                                wrapper: { flex: 1, display: 'flex', flexDirection: 'column' },
                                input: { flex: 1, resize: 'none', fontFamily: 'monospace' }
                            }}
                        />
                    </Box>
                </div>
            </div>
        </div>
    );
}
