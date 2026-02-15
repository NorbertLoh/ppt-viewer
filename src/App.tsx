import { useState } from 'react';
import { Center, Text, ActionIcon } from '@mantine/core';
import { IconSettings } from '@tabler/icons-react';
import { LandingPage } from './components/LandingPage';
import { ViewerPage } from './components/ViewerPage';
import { SettingsModal } from './components/SettingsModal';
import type { Slide } from './electron'; // Import type

function App() {
  const [loading, setLoading] = useState(false);
  const [slides, setSlides] = useState<Slide[] | null>(null);
  const [currentFilePath, setCurrentFilePath] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const handleManualSelect = async () => {
    if (window.electronAPI) {
      try {
        const path = await window.electronAPI.selectFile();
        if (path) {
          processFile(path);
        }
      } catch (e) {
        console.error(e);
      }
    }
  };

  const processFile = async (filePath: string) => {
    setLoading(true);
    setError(null);
    setCurrentFilePath(filePath);
    try {
      if (window.electronAPI) {
        const response = await window.electronAPI.convertPptx(filePath);
        if (response.success) {
          setSlides(response.slides);
        } else {
          setError(response.error || 'Conversion failed');
        }
      } else {
        console.warn('Electron API not found');
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleSaveAll = async (updatedSlides: Slide[]) => {
    // ViewerPage handles the actual saving to file (including audio) via IPC.
    // We just need to update our local state here.
    setSlides(updatedSlides);
    // Optional: Show a subtle success message if needed, but ViewerPage already alerted.
  };

  const handleFileDrop = async (files: File[]) => {
    if (files.length > 0) {
      const file = files[0];
      // With contextIsolation: false, path should be available directly on the file object
      // or via our simplified preload
      let filePath = (file as any).path;

      console.log('File Object:', file);
      console.log('Detected Path:', filePath);

      if (!filePath) {
        setError('Could not detect file path. Please try using the "Select File" button below.');
        return;
      }

      processFile(filePath);
    }
  };

  if (loading) {
    return (
      <Center h="100vh">
        <Text ml="md">Processing...</Text>
      </Center>
    );
  }

  if (error) {
    return (
      <Center h="100vh" style={{ flexDirection: 'column' }}>
        <Text c="red" size="xl">Error: {error}</Text>
        <Text
          c="blue"
          style={{ cursor: 'pointer', marginTop: '1rem' }}
          onClick={() => { setError(null); setSlides(null); }}
        >
          Try Again
        </Text>
      </Center>
    );
  }

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      {!slides ? (
        <>
          <ActionIcon
            variant="subtle"
            size="lg"
            pos="absolute"
            top={10}
            left={10}
            onClick={() => setSettingsOpen(true)}
          >
            <IconSettings size={24} />
          </ActionIcon>
          <LandingPage onDrop={handleFileDrop} onSelectFile={handleManualSelect} />
          {loading && <Text ta="center" mt="sm">Analysing file...</Text>}
        </>
      ) : (
        <ViewerPage slides={slides} onSave={handleSaveAll} onBack={() => setSlides(null)} filePath={currentFilePath || ''} />
      )}
      <SettingsModal opened={settingsOpen} onClose={() => setSettingsOpen(false)} />
    </div>
  );
}

export default App;
