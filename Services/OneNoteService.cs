using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteExporter.Models;

namespace OneNoteExporter.Services;

/// <summary>
/// Wraps the OneNote COM API. Must be constructed and called on a background thread
/// (via Task.Run) because COM calls and Thread.Sleep block.
/// Dispose() must be called when the application closes.
/// </summary>
public class OneNoteService : IDisposable
{
    private Application? _oneNote;
    private bool _disposed = false;
    private bool _oneNoteWasRunning = false;
    private bool _oneNoteClosedAttempted = false;

    public OneNoteService()
    {
        var oneNoteProcesses = Process.GetProcessesByName("ONENOTE");
        _oneNoteWasRunning = oneNoteProcesses.Length > 0;

        try
        {
            _oneNote = new Application();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Failed to initialize OneNote Desktop. " +
                "Please ensure that OneNote 2016 is installed.", ex);
        }
    }

    // ── Public API ───────────────────────────────────────────────────────────

    public VersionInfo GetVersionInfo()
    {
        var info = new VersionInfo { OneNoteInstalled = _oneNote != null };

        if (_oneNote != null)
        {
            try
            {
                _oneNote.GetHierarchy("", HierarchyScope.hsNotebooks, out _);
                info.OneNoteVersion = "OneNote Desktop (COM API available)";
            }
            catch
            {
                info.OneNoteVersion = "Unknown";
            }
        }

        return info;
    }

    public List<NotebookInfo> GetNotebooks()
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized.");

        var notebooks = new List<NotebookInfo>();

        _oneNote.GetHierarchy("", HierarchyScope.hsNotebooks, out string xml);

        var xdoc = XDocument.Parse(xml);
        var ns   = xdoc.Root?.Name.Namespace;
        if (ns == null) return notebooks;

        foreach (var nb in xdoc.Descendants(ns + "Notebook"))
        {
            notebooks.Add(new NotebookInfo
            {
                Id               = nb.Attribute("ID")?.Value              ?? "",
                Name             = nb.Attribute("name")?.Value            ?? "Unnamed",
                Path             = nb.Attribute("path")?.Value            ?? "",
                LastModified     = nb.Attribute("lastModifiedTime")?.Value ?? "",
                IsCurrentlyViewed = nb.Attribute("isCurrentlyViewed")?.Value == "true"
            });
        }

        return notebooks;
    }

    /// <summary>
    /// Exports a single notebook. Blocks the calling thread until finished.
    /// Call via Task.Run from the UI layer.
    /// </summary>
    public ExportResult ExportNotebook(
        string notebookId,
        string destinationPath,
        string exportFormat = "onepkg",
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized.");

        TryCloseOneNoteGracefully(progress);

        var result = new ExportResult();

        try
        {
            // Get notebook name and path from hierarchy
            _oneNote.GetHierarchy(notebookId, HierarchyScope.hsSelf, out string xml);
            var xdoc         = XDocument.Parse(xml);
            var notebookName = xdoc.Root?.Attribute("name")?.Value ?? "Notebook";
            var notebookPath = xdoc.Root?.Attribute("path")?.Value ?? "";

            bool isCloud = notebookPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                           notebookPath.StartsWith("http://",  StringComparison.OrdinalIgnoreCase);

            progress?.Report($"Exporting: {notebookName} (format: {exportFormat})");

            // Determine file extension and PublishFormat
            string fileExtension = exportFormat.ToLowerInvariant() switch
            {
                "xps" => ".xps",
                "pdf" => ".pdf",
                _     => ".onepkg"
            };

            PublishFormat publishFormat = exportFormat.ToLowerInvariant() switch
            {
                "xps" => PublishFormat.pfXPS,
                "pdf" => PublishFormat.pfPDF,
                _     => PublishFormat.pfOneNotePackage
            };

            // Sanitize filename
            var sanitizedName = string.Join("_", notebookName.Split(Path.GetInvalidFileNameChars()));
            var fullPath      = Path.Combine(destinationPath, sanitizedName + fileExtension);

            Directory.CreateDirectory(destinationPath);

            // Open notebook (required for cloud notebooks)
            _oneNote.OpenHierarchy(notebookPath, "", out string openedId, CreateFileType.cftNone);

            // Remove existing file so WaitForFile doesn't detect the old file as done
            if (File.Exists(fullPath))
            {
                progress?.Report("Removing previous export file...");
                File.Delete(fullPath);
            }

            // Trigger the export (Publish returns immediately; OneNote writes async)
            progress?.Report("OneNote is writing in the background...");
            _oneNote.Publish(openedId, fullPath, publishFormat, "");

            // Wait for the file to appear and stabilise
            result = WaitForFile(fullPath, isCloud, progress, ct);

            if (result.Success)
                result.ExportedPath = fullPath;
        }
        catch (OperationCanceledException)
        {
            result.Success = false;
            result.Message = "Export cancelled.";
        }
        catch (UnauthorizedAccessException ex)
        {
            result.Success = false;
            result.Message = $"Access denied: {ex.Message}";
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x8004201A))
        {
            result.Success = false;
            result.Message = "OneNote Error 0x8004201A: Cannot export. " +
                             "The notebook may have password-protected sections or be offline.";
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800706BA))
        {
            result.Success = false;
            result.Message = "OneNote RPC timeout (0x800706BA). " +
                             "Large notebooks may fail with PDF/XPS. Try .onepkg format instead.";
        }
        catch (COMException ex)
        {
            result.Success = false;
            result.Message = $"OneNote COM error 0x{ex.HResult:X}: {ex.Message}";
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = $"Error during export: {ex.Message}";
        }

        return result;
    }

    /// <summary>
    /// Exports all notebooks sequentially. Calls ExportNotebook for each.
    /// </summary>
    public ExportResult ExportAllNotebooks(
        string destinationPath,
        string exportFormat = "onepkg",
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        var result       = new ExportResult { Success = true };
        int exported     = 0, failed = 0;
        var messages     = new List<string>();

        try
        {
            var notebooks = GetNotebooks();
            progress?.Report($"Starting export of {notebooks.Count} notebook(s)...");

            foreach (var nb in notebooks)
            {
                ct.ThrowIfCancellationRequested();

                progress?.Report($"Exporting {exported + failed + 1}/{notebooks.Count}: {nb.Name}");

                var r = ExportNotebook(nb.Id, destinationPath, exportFormat, progress, ct);

                if (r.Success)
                {
                    exported++;
                    messages.Add($"✓ {nb.Name}");
                }
                else
                {
                    failed++;
                    messages.Add($"✗ {nb.Name}: {r.Message}");
                }
            }

            result.Success      = failed == 0;
            result.Message      = $"Export completed: {exported} successful, {failed} failed\n\n" +
                                   string.Join("\n", messages);
            result.ExportedPath = destinationPath;
        }
        catch (OperationCanceledException)
        {
            result.Success = false;
            result.Message = "Export cancelled.";
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = $"Fatal export error: {ex.Message}";
        }

        return result;
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    /// <summary>
    /// Polls until the exported file exists and its size has been stable for 10 seconds.
    /// Cloud notebooks get a 30-minute timeout; local notebooks get 20 minutes.
    /// </summary>
    private static ExportResult WaitForFile(
        string fullPath,
        bool isCloud,
        IProgress<string>? progress,
        CancellationToken ct)
    {
        int  maxAttempts     = isCloud ? 900 : 600; // 30 min : 20 min
        int  checkIntervalMs = 2000;
        long previousSize    = -1;
        int  stableCount     = 0;
        bool fileCreated     = false;
        long fileSize        = 0;

        // Initial wait – give OneNote a moment to start writing
        if (ct.WaitHandle.WaitOne(checkIntervalMs)) ct.ThrowIfCancellationRequested();

        for (int i = 0; i < maxAttempts; i++)
        {
            ct.ThrowIfCancellationRequested();

            if (File.Exists(fullPath))
            {
                fileCreated = true;
                fileSize    = new FileInfo(fullPath).Length;

                if (fileSize > 0 && fileSize == previousSize)
                {
                    stableCount++;
                    if (stableCount >= 5)
                    {
                        progress?.Report($"✓ File written! Final size: {FormatBytes(fileSize)}");
                        return new ExportResult
                        {
                            Success  = true,
                            Message  = $"Exported successfully ({FormatBytes(fileSize)})",
                            ExportedPath = fullPath
                        };
                    }
                    progress?.Report($"Size stable at {FormatBytes(fileSize)} ({stableCount * 2}s)...");
                }
                else if (fileSize > 0)
                {
                    double mbps = previousSize > 0
                        ? (fileSize - previousSize) / 1024.0 / 1024.0 / (checkIntervalMs / 1000.0)
                        : 0;
                    progress?.Report(mbps > 0
                        ? $"Writing: {FormatBytes(fileSize)} (~{mbps:0.0} MB/s)"
                        : $"Writing: {FormatBytes(fileSize)}");
                    stableCount = 0;
                }
                else
                {
                    progress?.Report("File created, waiting for data...");
                    stableCount = 0;
                }

                previousSize = fileSize;
            }
            else
            {
                if (i % 10 == 0)
                    progress?.Report($"Waiting for file creation... ({i * checkIntervalMs / 1000}s)");
            }

            if (ct.WaitHandle.WaitOne(checkIntervalMs)) ct.ThrowIfCancellationRequested();
        }

        // Timed out
        int timeoutMin = maxAttempts * checkIntervalMs / 60_000;
        if (fileCreated && fileSize == 0)
            return new ExportResult { Success = false, Message = "File was created but is empty (0 bytes)." };

        return new ExportResult
        {
            Success = false,
            Message = $"Timeout after {timeoutMin} min. OneNote may still be writing in the background. " +
                      $"Check {Path.GetDirectoryName(fullPath)} again in a few minutes."
        };
    }

    private void TryCloseOneNoteGracefully(IProgress<string>? progress)
    {
        if (_oneNoteClosedAttempted) return;
        _oneNoteClosedAttempted = true;

        var procs = Process.GetProcessesByName("ONENOTE");
        if (procs.Length == 0) return;

        progress?.Report($"Closing {procs.Length} OneNote process(es) before export...");

        foreach (var p in procs)
        {
            try
            {
                if (p.CloseMainWindow())
                    p.WaitForExit(5000);
            }
            catch { /* ignore */ }
            finally { p.Dispose(); }
        }

        Thread.Sleep(1000);
    }

    private static string FormatBytes(long bytes)
    {
        string[] units = { "Bytes", "KB", "MB", "GB" };
        double   value = bytes;
        int      order = 0;
        while (value >= 1024 && order < units.Length - 1) { order++; value /= 1024; }
        return $"{value:0.##} {units[order]}";
    }

    // ── IDisposable ──────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;

        if (_oneNote != null)
        {
            try
            {
                Marshal.ReleaseComObject(_oneNote);
                _oneNote = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { /* ignore cleanup errors */ }
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
