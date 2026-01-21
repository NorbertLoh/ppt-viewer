import { useState } from 'react';
import { Center, Text } from '@mantine/core';
import { LandingPage } from './components/LandingPage';
import { ViewerPage } from './components/ViewerPage';
import type { Slide } from './electron'; // Import type

function App() {
  const [loading, setLoading] = useState(false);
  const [slides, setSlides] = useState<Slide[] | null>(null);
  const [currentFilePath, setCurrentFilePath] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

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
    if (!currentFilePath || !window.electronAPI) return;

    setLoading(true);
    try {
      const result = await window.electronAPI.saveAllNotes(currentFilePath, updatedSlides);
      if (!result.success) {
        alert('Failed to save: ' + result.error);
      } else {
        alert('Notes saved successfully!');
        setSlides(updatedSlides); // Update generic state
      }
    } catch (e: any) {
      alert('Error saving notes: ' + e.message);
    } finally {
      setLoading(false);
    }
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
          <LandingPage onDrop={handleFileDrop} onSelectFile={handleManualSelect} />
          {loading && <Text ta="center" mt="sm">Analysing file...</Text>}
        </>
      ) : (
        <ViewerPage slides={slides} onSave={handleSaveAll} onBack={() => setSlides(null)} filePath={currentFilePath || ''} />
      )}
    </div>
  );
}

export default App;
