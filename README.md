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

## Configuration

### 1. Google Cloud Text-to-Speech
To use the high-quality Chirp 3 HD voices, you need a Google Cloud Service Account key.

1.  Place your Service Account JSON key file in the root directory of the project.
2.  Rename the file to `gcp-key.json`.
3.  Ensure your `.env` file matches the following:
    ```env
    TTS_PROVIDER=gcp
    GOOGLE_APPLICATION_CREDENTIALS=./gcp-key.json
    ```

### 2. PowerPoint Add-in (PPAM)
For better stability with video export and audio insertion, install the helper add-in.

1.  Open PowerPoint.
2.  Go to **Tools** > **PowerPoint Add-ins...**
3.  Click the **+** button.
4.  Navigate to `electron/scripts/helper.ppam` in this project and select it.
5.  Check any security prompts to allow the macros to run.

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
