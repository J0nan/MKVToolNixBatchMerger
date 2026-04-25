# MKVToolNix Batch Merger

A Python-based GUI application for batch merging MKV files using MKVToolNix. This tool allows you to merge video, audio, and subtitle tracks from corresponding MKV files in two different folders into a single output folder.

## Features

- **Batch Processing**: Automatically finds matching filenames in two input folders.
- **Multithreaded Muxing**: Process multiple files concurrently to save time, with a customizable thread count.
- **Smart Duration Check**: Compares lengths of paired files before multiplexing. Presents a detailed prompt of mismatched items (`mm:ss.ms` precision), allowing you to proceed with all, seamlessly exclude just the problematic files, or abort entirely.
- **Visual Track Selection**: Interactively select which audio, video, or subtitle tracks to keep from each file.
- **Global Properties**: Easily modify MKV titles and control whether chapters, global tags, or attachments are copied.
- **Auto-Detect Path**: Automatically sets up Windows MKVToolNix paths to save manual browsing.
- **Graceful Cancellation**: Instantly cancel any ongoing batch merging with a button click.
- **Error Logging**: Generates standard log files (`mkv_merger_error_log.txt`) automatically tracking terminal errors during processing.
- **Batch Export**: Save configurations via `.bat` or `.sh` script generation, empowering headless terminal running outside the app!
- **Preset System**: Save and load track and global property configurations as JSON presets for frequent workflows.

## Prerequisites

- **Python 3.7+**
- **MKVToolNix** (specifically the `mkvmerge` command-line tool). You can download it from [MKVToolNix Download Page](https://mkvtoolnix.download/).

## Installation

1. Install the required python packages using pip:
   ```bash
   pip install -r requirements.txt
   ```
2. Ensure you know the installation path of your MKVToolNix (e.g., `C:\Program Files\MKVToolNix` on Windows).

## Usage

1. **Launch the App**:
   Run the following command in your terminal/command prompt:
   ```bash
   python pymkv_merger_app.py
   ```

2. **Select Directories**:
   - `MKVToolNix Path`: Browse and select the folder where MKVToolNix is installed.
   - `Input Folder 1`: Select the first folder containing MKV files.
   - `Input Folder 2`: Select the second folder containing matching MKV files.
   - `Output Folder`: Select the directory where the merged files will be saved.

3. **Configure Batch Settings**:
   - Change the `Max Threads` value to increase or decrease concurrent multiplexing. A sensible default based on your CPU is set.
   - Click **Analyze Files & Select Tracks**. The application will analyze the first matching file it finds to show you all available tracks.

4. **Select Tracks & Properties**:
   - A new window lets you review tracks from `File 1` and `File 2`. Uncheck `Include` for any tracks you don't want in the final output.
   - You can edit Language tags, Track Names, and set Default/Forced flags.
   - Control global tags, chapters, and attachments via the "Global Properties" box.

5. **Start Merging**:
   - Click **Start Merging**. A progress bar will show the live merging status.

## Note
This application utilizes `pymkv` as a wrapper for `mkvmerge`. If you experience JSON parse errors, ensure the paths to `mkvmerge.exe` and `mkvextract.exe` are correct and accessible.