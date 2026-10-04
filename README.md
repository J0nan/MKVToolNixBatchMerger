# MKVToolNix Batch Merger

A Python/Tkinter GUI that batch-multiplexes matching MKV files from two folders (for example, one folder with the base video and one with dubbed audio or extra subtitles) into a single output folder, powered by MKVToolNix's `mkvmerge`.

## Features

### Processing
- **Batch Matching**: Automatically finds files with identical names across the two input folders and processes the whole batch at once.
- **Multithreaded Muxing**: Merge several files concurrently with a configurable thread count (1–16, defaults to `min(4, CPU cores)`).
- **Smart Duration Check**: Compares pair lengths before multiplexing. Pairs that differ by more than 100 ms are listed with `hh:mm:ss.mmm` precision, and you choose to merge everything, exclude only the mismatched files, or cancel.
- **Graceful Cancellation**: Cancel an ongoing batch with one click; already-finished files are kept.
- **Error Logging**: Failures per file are appended to `mkv_merger_error_log.txt` in the output folder.

### Track Selection (scales to many tracks)
- Tabbed window with **File 1**, **File 2**, and **Output Options** tabs, so the action buttons stay visible no matter how many tracks a file has.
- **Scrollable track table** per file (mouse wheel + scrollbars) with a sticky column header — comfortable even with dozens of audio/video/subtitle tracks.
- **Bulk controls**: *Include all* / *Exclude all* per file, a **type filter** (All / Video / Audio / Subtitles / Other), and a live *"N of M selected"* counter.
- Per-track editing of **Language** tag, **Track Name**, and **Default** / **Forced** flags, plus an *Include* checkbox per row.

### Output Options & Inspection
- **Metadata Title** for the muxed files.
- **Chapters / Global Tags / Attachments**: include-exclusions per source file. Options are auto-disabled with a note when the sample files don't contain them.
- **Built-in viewers**: inspect each file's global tags (raw JSON) and attachments list.
- **Attachment extraction**: double-click an attachment in the viewer to save it via `mkvextract`.

### Workflow
- **Preset System**: Save and reload the full track selection + output options as a JSON preset.
- **Batch Script Export**: Generate a standalone `.bat` / `.sh` script with the same configuration, for headless runs outside the app.
- **Settings Persistence**: Paths and thread count are saved automatically (`mkv_merger_settings.json`).
- **Auto-Detect Path**: On Windows, the default MKVToolNix install folder (`C:\Program Files\MKVToolNix`) is detected automatically.

## Prerequisites

- **Python 3.9+** (`pymkv2` ≥ 2.2 requires Python 3.10+)
- **MKVToolNix** — the `mkvmerge` command-line tool is required; `mkvextract` is optional (used only for attachment extraction). Download from the [MKVToolNix download page](https://mkvtoolnix.download/).
- tkinter (ships with standard Python installs)

## Installation

```bash
pip install -r requirements.txt
```

`requirements.txt` installs **pymkv2** (imported as `pymkv`), a wrapper around the MKVToolNix tools.

## Usage

1. **Launch the app**:
   ```bash
   python pymkv_merger_app.py
   ```

2. **Fill in the main window**:
   | Field | Meaning |
   |-------|---------|
   | MKVToolNix Path | Folder where MKVToolNix is installed (auto-detected on Windows) |
   | Input Folder 1 | First folder of MKV files (base for the merge) |
   | Input Folder 2 | Second folder; only files whose names match Folder 1 are processed |
   | Output Folder | Where merged files are written (same file name as the source pair) |
   | Max Threads | Concurrent muxing jobs (1–16) |

3. **Click "Analyze Files & Select Tracks"**: the app matches filenames, then checks the duration of every pair. If mismatches (> 100 ms) exist, a dialog lists them and offers **Continue merging ALL**, **EXCLUDE mismatched files**, or **Cancel**.

4. **Pick tracks in the tabbed window** (the first matched pair is analyzed as the sample; selections are applied to the batch by track ID — tracks missing from a pair are skipped):
   - *File 1* / *File 2* tabs: scrollable tables; tick *Include* per track or use *Include all* / *Exclude all*; narrow the view with the type *Filter*; the *N of M selected* counter tracks the current choice. Edit *Language* and *Name*, toggle *Default* / *Forced* per track.
   - *Output Options* tab: metadata title and per-file include switches for chapters, global tags, and attachments, with **View tags** / **View attachments...** inspectors.
   - Bottom bar (always visible): **Save Preset**, **Load Preset**, **Export Batch Script**, **Start Merging**.

5. **Start merging**: a progress window shows live status per file and a cancel button. On success the merged `.mkv` files appear in the output folder; on partial failure details are written to `mkv_merger_error_log.txt` there.

## Files this app creates

| File | Location | Purpose |
|------|----------|---------|
| `mkv_merger_settings.json` | App folder | Persisted paths and max threads |
| Preset `*.json` | Chosen by you | Saved track selection + output options |
| `mkv_merger_error_log.txt` | Output folder | Per-file merge errors |
| Exported `.bat` / `.sh` | Chosen by you | Headless merge commands with the current configuration |

## Note

This application uses **pymkv2** (imported as `pymkv`) as a wrapper for `mkvmerge`. If you get JSON parse errors, make sure the MKVToolNix folder path is correct and contains both `mkvmerge` and (for attachment extraction) `mkvextract`.
