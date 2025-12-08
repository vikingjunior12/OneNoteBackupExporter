# Release Notes - Version 0.8.0

**Release Date:** January 6, 2025

## 🎉 What's New

### Multi-Format Export Support

Version 0.8.0 introduces **multi-format export** capabilities, giving you more flexibility when backing up your OneNote notebooks!

#### New Export Formats

In addition to the existing OneNote Package format, you can now export your notebooks as:

- **📄 XPS Document (.xps)** - Windows-native fixed-layout format with excellent formatting preservation. Opens natively in Windows 10/11 with the built-in XPS Viewer. Ideal for archival purposes when you want to maintain the exact layout and formatting.

- **📄 PDF Document (.pdf)** - Universal document format that works on any device and operating system. Perfect for sharing your notes with others or viewing on non-Windows devices. Note: Some complex OneNote structures may experience minor layout changes during conversion.

- **📦 OneNote Package (.onepkg)** - The original native backup format (recommended for best quality and re-importing into OneNote).

### How to Use

1. Open OneNote Backup Exporter
2. Select your notebooks
3. Choose your preferred export format from the dropdown menu
4. Click "Export Selected" or "Export All"

The export format is preserved for all notebooks in a batch export, ensuring consistency across your backup.

## 🔧 Technical Details

- All three formats use OneNote's official COM API for maximum compatibility
- XPS export provides better formatting preservation compared to older MHTML format
- PDF export offers universal accessibility across all platforms
- Export format selection flows seamlessly through the entire application stack (Frontend → Go Backend → C# Helper → OneNote COM API)

## 📋 System Requirements

- Windows 10 or Windows 11
- OneNote Desktop 2016 ("OneNote für Windows")
- .NET 8.0 Runtime

## 🐛 Known Limitations

- Password-protected sections may fail during export
- Very large notebooks (>500MB) may experience timeouts
- PDF export may not preserve all OneNote formatting perfectly for complex page layouts
- Only compatible with OneNote Desktop version (not OneNote UWP/Windows 10 app)

## 💡 Tips

- **For archival/backup:** Use OneNote Package (.onepkg) - highest quality, can be re-imported
- **For Windows sharing:** Use XPS (.xps) - excellent formatting preservation, Windows-native
- **For universal access:** Use PDF (.pdf) - works everywhere, slight formatting compromises

---

**Full Changelog:** https://github.com/yourusername/OneNoteBackupExporter/compare/v0.7.0...v0.8.0

**Download:** [Release v0.8.0](https://github.com/yourusername/OneNoteBackupExporter/releases/tag/v0.8.0)
