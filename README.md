# PPT Viewer & Video Generator

An Electron-based application that views PowerPoint presentations, extracts speaker notes, generates TTS audio using a local server, and automates PowerPoint to create MP4 videos with synchronized audio.

## Prerequisites

-   **OS**: Windows 10/11 (Required for PowerPoint automation)
-   **Software**:
    -   [Node.js](https://nodejs.org/) (v16+)
    -   Microsoft PowerPoint (Desktop version installed)
    -   [Docker Desktop](https://www.docker.com/products/docker-desktop/) (Required only for **Local/Dev** mode)

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

### 1. Production Mode (Google Cloud TTS) - Recommended
This mode uses Google's high-quality "Chirp 3 HD" voices.

1.  **Obtain Key**: Create a Service Account in Google Cloud and download the JSON key file.
2.  **Launch App**: Open the PPT Viewer application.
3.  **Configure**: Click the **Settings (Gear) icon** in the top-left corner.
4.  **Upload**: Click **"Select JSON Key"** and choose your downloaded `.json` file.
    *   The app will securely store the path and automatically switch to Google Cloud TTS.

### 2. Development Mode (Local TTS)
Useful for offline development or testing without incurring costs. Uses Mycroft Mimic 3.

**Start the Local TTS Server:**
Run the following command in a terminal:

```bash
docker run -it --user root -p 59125:59125 mycroftai/mimic3
```

> **Note**: This maps port `59125` locally. If you have not uploaded a GCP key, the app will attempt to fallback to this local server.

### 3. PowerPoint Add-in (PPAM)
For better stability with video export and audio insertion, install the helper add-in.

1.  Open PowerPoint.
2.  Go to **Tools** > **PowerPoint Add-ins...**
3.  Click the **+** button.
4.  Navigate to `electron/scripts/ppt-tools.ppam` in this project (or the release folder) and select it.
5.  Check any security prompts to allow the macros to run.

## Running the Application

Once the TTS server is running, start the application in development mode:

```bash
npm run dev
```

1.  Select a PowerPoint file (`.pptx`).
2.  View slides and notes.
3.  Click **"Generate Video"** to create an MP4 with AI narration.

## Troubleshooting

### "App cannot be opened because Apple cannot check it for malicious software"
Since this application is not signed with an Apple Developer ID, you may see this warning on macOS.

**To fix this:**
1.  **Right-click** (or Control-click) the `PPT Viewer.app`.
2.  Select **Open** from the context menu.
3.  Click **Open** in the dialog box that appears.

Alternatively, you can remove the quarantine attribute via terminal:
```bash
xattr -cr "/path/to/PPT Viewer.app"
```
