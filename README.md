# WhisperTools Module Documentation

## Overview
The `WhisperTools` PowerShell module provides tools for:
- Converting video files to audio files.
- Transcribing audio files to text using OpenAI Whisper.
- Combining video-to-audio conversion and transcription in a single step.

## Installation

### Prerequisites
1. Ensure `winget` (Windows Package Manager) is installed on your system. If not, install it from the Microsoft Store.
2. Install the `WhisperTools` module by placing the `WhisperTools.psm1` file in a directory under your PowerShell `Modules` folder (e.g., `Documents\WindowsPowerShell\Modules\WhisperTools`).

### Installing Dependencies
Run the following commands to install the required dependencies:

1. Install FFmpeg:
   ```powershell
   Install-VideoToAudio
   ```

2. Install Python, Torch, and Whisper:
   ```powershell
   Install-AudioToTxt
   ```

## Usage

### Cmdlets

#### 1. `Convert-VideoToAudio`
Converts a video file to an audio file.

**Parameters:**
- `-InputFile` (Mandatory): Path to the input video file.
- `-Output` (Optional): Path to the output audio file. Defaults to the same directory as the input file with a `.wav` extension.

**Example:**
```powershell
Convert-VideoToAudio -InputFile 'C:\path\to\video.mp4'
```

#### 2. `Convert-AudioToTxt`
Transcribes an audio file to text.

**Parameters:**
- `-InputFile` (Mandatory): Path to the input audio file.
- `-Output` (Optional): Path to the output text file. Defaults to the same directory as the input file with a `.txt` extension.
- `-Model` (Optional): Whisper model to use. Options: `tiny`, `base`, `small`, `medium`, `large-v3`. Default: `medium`.
- `-Language` (Optional): Language code for transcription. Default: `ru`.
- `-Threads` (Optional): Number of threads to use.
- `-Format` (Optional): Output format. Options: `txt`, `srt`, `vtt`. Default: `txt`.

**Example:**
```powershell
Convert-AudioToTxt -InputFile 'C:\path\to\audio.wav' -Format 'srt'
```

#### 3. `Convert-VideoToTxt`
Combines video-to-audio conversion and transcription in a single step.

**Parameters:**
- `-InputFile` (Mandatory): Path to the input video file.
- `-Output` (Optional): Path to the output text file. Defaults to the same directory as the input file with a `.txt` extension.
- `-Model` (Optional): Whisper model to use. Options: `tiny`, `base`, `small`, `medium`, `large-v3`. Default: `medium`.
- `-Language` (Optional): Language code for transcription. Default: `ru`.
- `-Threads` (Optional): Number of threads to use.
- `-Format` (Optional): Output format. Options: `txt`, `srt`, `vtt`. Default: `txt`.

**Example:**
```powershell
Convert-VideoToTxt -InputFile 'C:\path\to\video.mp4' -Format 'srt'
```

## Notes
- Ensure all dependencies are installed before running the cmdlets.
- The module uses CPU-only processing for transcription.
- Output files are saved in the same directory as the input file by default.

## Support
For issues or feature requests, please contact the module maintainer.
