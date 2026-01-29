# PPT Viewer & Video Generator

An Electron-based application that views PowerPoint presentations, extracts speaker notes, generates TTS audio using a local server, and automates PowerPoint to create MP4 videos with synchronized audio.

## Prerequisites

-   **OS**: Windows 10/11 (Required for PowerPoint automation)
-   **Software**:
    -   [Node.js](https://nodejs.org/) (v16+)
    -   Microsoft PowerPoint (Desktop version installed)
    -   [Docker Desktop](https://www.docker.com/products/docker-desktop/) (For local TTS)

## Setup

1.  Clone the repository:
    ```bash
    git clone https://github.com/NorbertLoh/ppt-viewer.git
    cd ppt-viewer
    ```

2.  Install dependencies:
    ```bash
    npm install
    ```

## Running the TTS Server

This application requires a local Mycroft Mimic 3 TTS server running in a Docker container.

Run the following command in a terminal to start the TTS server:

```bash
docker run -it --user root -p 59125:59125 mycroftai/mimic3
```

> **Note**: This maps port `59125` locally, which the application relies on to generate audio.

## Running the Application

Once the TTS server is running, start the application in development mode:

```bash
npm run dev
```

1.  Select a PowerPoint file (`.pptx`).
2.  View slides and notes.
3.  Click **"Generate Video"** to create an MP4 with AI narration.
