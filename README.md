# OneNote Exporter

A Windows desktop application for exporting Microsoft OneNote notebooks to portable file formats. Built with C# and WPF, it communicates with OneNote Desktop through the COM API — no third-party cloud services, no subscriptions, no internet connection required.

![Main window showing the notebook list and export options](screenshots/main-window.png)

---

## Requirements

- Windows 10 or Windows 11
- OneNote Desktop 2016 or later (including OneNote as part of Microsoft 365)

> The UWP/Store version of OneNote (installed from the Microsoft Store without a Microsoft 365 subscription) is not supported.

---

## How It Works

On startup, the application connects to OneNote through the Windows COM API. It reads the list of all notebooks registered in OneNote and displays them in the main window. You select which notebooks to export, choose a destination folder and output format, then click Export.

![Notebook selection with export format dropdown](screenshots/notebook-selection.png)

The export process works as follows:

1. The application tells OneNote to open the selected notebook (required for cloud-hosted notebooks stored on SharePoint or OneDrive).
2. OneNote writes the file to disk asynchronously in the background.
3. The application monitors the output file, polling its size every two seconds until the file size has been stable for at least ten seconds, indicating the write is complete.
4. Once all selected notebooks are exported, the destination folder opens automatically in Windows Explorer.

Because OneNote handles the actual writing, export time depends entirely on notebook size and whether the notebook is stored locally or in the cloud. Large cloud notebooks can take several minutes.

---

## Class Notebooks and myLMS

OneNote Exporter fully supports **class notebooks** (Klassennotizbücher), including those provided through school or university platforms such as myLMS. These notebooks appear in the list like any other notebook. Select them and export as usual — the application handles the cloud sync and download automatically before writing the export file.

---

## Export Formats

| Format | Extension | Notes |
|---|---|---|
| OneNote Package | `.onepkg` | Recommended. Self-contained archive, re-importable into OneNote. |
| PDF | `.pdf` | Good for archiving and sharing. May time out on very large notebooks. |
| XPS | `.xps` | Microsoft's alternative to PDF. Less commonly used. |
| Local Backup Copy | — | Copies the raw OneNote backup folder directly (fastest option). |

### Local Backup Copy

When this option is selected, the application copies the contents of OneNote's automatic backup folder (`%LOCALAPPDATA%\Microsoft\OneNote\16.0\Sicherung`) directly to the destination you specify. This is not a published export — it is a file system copy of the raw backup files OneNote maintains automatically.

All notebooks are included when using this mode. Individual selection is not available.

![Progress display during an active export](screenshots/export-progress.png)

---

## Cancellation

An export can be cancelled at any time using the Cancel button, which appears during an active export. Cancellation is cooperative: the application stops waiting for the current file, but OneNote may continue writing in the background. If the cancel does not respond in time, the application will force-terminate the OneNote process.

---

## Timeouts

| Notebook type | Timeout |
|---|---|
| Local notebooks | 20 minutes |
| Cloud notebooks (SharePoint / OneDrive) | 30 minutes |

If a timeout occurs, OneNote may still be writing the file. Check the destination folder a few minutes after the timeout message appears — the file may complete on its own.

---

## Known Limitations

- **Password-protected sections** cannot be exported. OneNote returns error `0x8004201A` for notebooks containing locked sections. Unlock all sections before exporting.
- **PDF and XPS formats** can cause RPC timeouts (`0x800706BA`) on very large notebooks. If this happens, use the OneNote Package format instead.
- **Cloud notebooks must be synced** before export. If a notebook is listed as offline in OneNote, the export will fail.

---

## Building from Source

```powershell
# Build
dotnet build

# Run in debug mode
dotnet run

# Publish as self-contained executable
dotnet publish -c Release -r win-x86 --self-contained true
```

---

## Architecture Overview

The application is a single WPF project with no subprocesses. All communication with OneNote happens in-process through the COM API.

```
MainWindow.xaml.cs
    |
    +-- Services/OneNoteService.cs    (COM API wrapper)
    |       GetNotebooks()
    |       ExportNotebook()
    |       ExportAllNotebooks()
    |
    +-- Helpers/FileHelper.cs         (Local backup, folder operations)
    |
    +-- Models/Models.cs              (NotebookInfo, ExportResult, ViewModels)
```

COM calls are blocking and run on background threads via `Task.Run`. Progress reporting uses `IProgress<string>`, which marshals updates back to the UI thread automatically.

---

## License

[Your license here]
