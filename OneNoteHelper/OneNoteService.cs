using System.Xml.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteHelper;

public class OneNoteService : IDisposable
{
    private Application? _oneNote;
    private bool _disposed = false;
    private bool _oneNoteWasRunning = false;
    private bool _oneNoteClosedAttempted = false;

    public OneNoteService()
    {
        try
        {
            // Check if OneNote is already running
            var oneNoteProcesses = Process.GetProcessesByName("ONENOTE");
            _oneNoteWasRunning = oneNoteProcesses.Length > 0;

            if (_oneNoteWasRunning)
            {
                Console.Error.WriteLine("WARNING: OneNote is already running. This may cause problems.");
                Console.Error.WriteLine("Recommendation: Close OneNote and restart the export.");

                // Try to connect anyway, but with awareness that it might hang
                Console.Error.WriteLine("Attempting to connect anyway...");
            }

            _oneNote = new Application();
            Console.Error.WriteLine("OneNote COM connection established successfully.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Failed to initialize OneNote Desktop. " +
                "Please ensure that OneNote 2016 is installed.", ex);
        }
    }

    public VersionInfo GetVersionInfo()
    {
        var info = new VersionInfo
        {
            OneNoteInstalled = _oneNote != null
        };

        if (_oneNote != null)
        {
            try
            {
                // Try to get OneNote version info
                _oneNote.GetHierarchy("", HierarchyScope.hsNotebooks, out string xml);
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
            throw new InvalidOperationException("OneNote is not initialized");

        var notebooks = new List<NotebookInfo>();

        try
        {
            // Get the notebook hierarchy XML
            _oneNote.GetHierarchy("", HierarchyScope.hsNotebooks, out string xml);

            // Parse the XML
            var xdoc = XDocument.Parse(xml);
            var ns = xdoc.Root?.Name.Namespace;

            if (ns == null) return notebooks;

            var notebookElements = xdoc.Descendants(ns + "Notebook");

            foreach (var nb in notebookElements)
            {
                var notebook = new NotebookInfo
                {
                    Id = nb.Attribute("ID")?.Value ?? "",
                    Name = nb.Attribute("name")?.Value ?? "Unnamed",
                    Path = nb.Attribute("path")?.Value ?? "",
                    LastModified = nb.Attribute("lastModifiedTime")?.Value ?? "",
                    IsCurrentlyViewed = nb.Attribute("isCurrentlyViewed")?.Value == "true"
                };

                notebooks.Add(notebook);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Error retrieving notebooks: " + ex.Message, ex);
        }

        return notebooks;
    }

    public ExportResult ExportNotebook(string notebookId, string destinationPath, string exportFormat = "onepkg")
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized");

        // Try to close OneNote gracefully before export (only happens once)
        TryCloseOneNoteGracefully();

        var result = new ExportResult();

        try
        {
            // Get notebook info for naming
            _oneNote.GetHierarchy(notebookId, HierarchyScope.hsSelf, out string xml);
            var xdoc = XDocument.Parse(xml);
            var ns = xdoc.Root?.Name.Namespace;
            var notebookName = xdoc.Root?.Attribute("name")?.Value ?? "Notebook";
            var notebookPath = xdoc.Root?.Attribute("path")?.Value ?? "";

            Console.Error.WriteLine($"=== Exporting Notebook ===");
            Console.Error.WriteLine($"Name: {notebookName}");
            Console.Error.WriteLine($"ID: {notebookId}");
            Console.Error.WriteLine($"Path: {notebookPath}");
            Console.Error.WriteLine($"Format: {exportFormat}");

            // Check if this is a SharePoint/OneDrive notebook (starts with https://)
            bool isCloudNotebook = notebookPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                                   notebookPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase);

            if (isCloudNotebook)
            {
                Console.Error.WriteLine($"ℹ INFO: This is a cloud notebook (SharePoint/OneDrive)");
            }

            // Determine file extension based on export format
            string fileExtension = exportFormat.ToLowerInvariant() switch
            {
                "xps" => ".xps",
                "pdf" => ".pdf",
                _ => ".onepkg"
            };

            // Sanitize filename
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitizedName = string.Join("_", notebookName.Split(invalidChars));
            var fileName = $"{sanitizedName}{fileExtension}";
            var fullPath = Path.Combine(destinationPath, fileName);

            // Ensure destination directory exists
            Directory.CreateDirectory(destinationPath);
            Console.Error.WriteLine($"Destination directory: {destinationPath}");
            Console.Error.WriteLine($"Destination file: {fullPath}");

            // Step 1: Ensure notebook is open (critical for SharePoint/OneDrive notebooks)
            Console.Error.WriteLine($"Opening notebook hierarchy...");
            _oneNote.OpenHierarchy(
                notebookPath,
                "",
                out string openedNotebookId,
                CreateFileType.cftNone
            );
            Console.Error.WriteLine($"Opened notebook ID: {openedNotebookId}");

            // Step 2: Publish to the specified format using the opened notebook ID
            // Determine PublishFormat enum value based on export format
            PublishFormat format = exportFormat.ToLowerInvariant() switch
            {
                "xps" => PublishFormat.pfXPS,
                "pdf" => PublishFormat.pfPDF,
                _ => PublishFormat.pfOneNotePackage
            };

            Console.Error.WriteLine($"Starting export operation...");
            Console.Error.WriteLine($"Note: OneNote writes the file asynchronously in the background");

            try
            {
                _oneNote.Publish(
                    openedNotebookId,  // Use the opened ID, not the original ID
                    fullPath,
                    format,
                    ""
                );

                Console.Error.WriteLine("✓ Publish() call successful (OneNote is now writing in the background)");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.Error.WriteLine($"✗ Publish() failed: 0x{ex.HResult:X}");
                throw; // Rethrow exception
            }

            // Wait for the publish operation to complete
            // IMPORTANT: Publish() returns immediately but OneNote writes the file asynchronously!
            Console.Error.WriteLine("Waiting for export operation to complete (OneNote writes asynchronously)...");
            System.Threading.Thread.Sleep(2000);

            // Wait for file creation AND for stable file size (OneNote writes in background)
            // Large notebooks (100+ MB) can take several minutes to write!
            bool fileCreated = false;
            long fileSize = 0;
            long previousSize = -1;
            int stableCount = 0;
            // Increased timeout: 20-30 minutes for large notebooks
            int maxAttempts = isCloudNotebook ? 900 : 600; // Cloud: up to 30 minutes, Local: up to 20 minutes
            int checkIntervalMs = 2000;

            Console.Error.WriteLine($"Monitoring file creation and write progress: {fullPath}");
            Console.Error.WriteLine($"Note: Large notebooks (100+ MB) may take several minutes...");

            for (int i = 0; i < maxAttempts; i++)
            {
                if (File.Exists(fullPath))
                {
                    fileCreated = true;
                    fileSize = new FileInfo(fullPath).Length;

                    // Check if file size is stable (not changing = OneNote finished writing)
                    // Wait 10 seconds to ensure OneNote is truly done writing
                    if (fileSize > 0 && fileSize == previousSize)
                    {
                        stableCount++;
                        // Wait for 5 checks (10 seconds) of stable size
                        int requiredStableChecks = 5;

                        if (stableCount >= requiredStableChecks)
                        {
                            // Size hasn't changed for 10 seconds = done!
                            Console.Error.WriteLine($"✓ File fully written! Final size: {FormatBytes(fileSize)} (stable for {stableCount * 2} seconds)");
                            break;
                        }

                        if (stableCount == 1 || stableCount == 3)
                        {
                            Console.Error.WriteLine($"File size stable at {FormatBytes(fileSize)} for {stableCount * 2} seconds (waiting {requiredStableChecks * 2} seconds total)");
                        }
                    }
                    else if (fileSize > 0)
                    {
                        // Show progress
                        if (previousSize > 0)
                        {
                            double mbPerSec = (fileSize - previousSize) / 1024.0 / 1024.0 / (checkIntervalMs / 1000.0);
                            Console.Error.WriteLine($"File growing: {FormatBytes(fileSize)} (+{FormatBytes(fileSize - previousSize)}, ~{mbPerSec:0.0} MB/s)");
                        }
                        else
                        {
                            Console.Error.WriteLine($"File growing: {FormatBytes(fileSize)}");
                        }
                        stableCount = 0; // Reset counter, file is still growing
                    }
                    else
                    {
                        Console.Error.WriteLine($"File exists but is still empty (0 bytes) - OneNote is still writing...");
                        stableCount = 0;
                    }

                    previousSize = fileSize;
                }
                else
                {
                    // File doesn't exist yet
                    if (i == 0)
                    {
                        Console.Error.WriteLine($"File not yet created, waiting...");
                        Console.Error.WriteLine($"Note: OneNote is working in the background, this may take up to 1 minute...");
                    }
                    else if (i % 10 == 0)
                    {
                        int secondsElapsed = (i * checkIntervalMs) / 1000;
                        Console.Error.WriteLine($"Still waiting for file creation... ({secondsElapsed}s elapsed)");
                    }
                }

                System.Threading.Thread.Sleep(checkIntervalMs);
            }

            if (fileCreated && fileSize > 0)
            {
                result.Success = true;
                result.Message = $"Notebook '{notebookName}' exported successfully ({FormatBytes(fileSize)})";
                result.ExportedPath = fullPath;
                Console.Error.WriteLine($"✓ Export successful: {fullPath} ({FormatBytes(fileSize)})");
            }
            else if (fileCreated && fileSize == 0)
            {
                result.Success = false;
                result.Message = $"Export failed: File was created but is empty (0 bytes). " +
                               $"This indicates a problem during the Publish() operation.";
                Console.Error.WriteLine($"✗ File exists but is empty!");
                Console.Error.WriteLine($"✗ OneNote may have aborted the export");
            }
            else
            {
                int timeoutMinutes = (maxAttempts * checkIntervalMs) / 60000;
                result.Success = false;
                result.Message = $"Timeout: File was not created after {timeoutMinutes} minutes. " +
                               $"IMPORTANT: OneNote may still be working in the background! " +
                               $"Check {destinationPath} again in a few minutes.";

                Console.Error.WriteLine($"✗ Timeout reached after {timeoutMinutes} minutes");
                Console.Error.WriteLine($"✗ File does not exist: {fullPath}");
                Console.Error.WriteLine($"");
                Console.Error.WriteLine($"⚠ IMPORTANT:");
                Console.Error.WriteLine($"  OneNote may still be writing the file in the background!");
                Console.Error.WriteLine($"  The export command was successfully sent to OneNote.");
                Console.Error.WriteLine($"  For very large notebooks (>500 MB), the write operation");
                Console.Error.WriteLine($"  may take longer than {timeoutMinutes} minutes.");
                Console.Error.WriteLine($"");
                Console.Error.WriteLine($"Recommendation:");
                Console.Error.WriteLine($"  1. Do NOT restart the program (would start a new export)");
                Console.Error.WriteLine($"  2. Check the directory {destinationPath}");
                Console.Error.WriteLine($"     again in 5-10 minutes");
                Console.Error.WriteLine($"  3. Monitor OneNote.exe in Task Manager (high CPU/Disk = still working)");
                Console.Error.WriteLine($"");
                Console.Error.WriteLine($"If no file exists after 15 minutes, possible causes:");
                Console.Error.WriteLine($"  - Password-protected sections in the notebook");
                Console.Error.WriteLine($"  - Notebook is offline or not accessible");
                Console.Error.WriteLine($"  - OneNote process has no write permissions to destination directory");
                Console.Error.WriteLine($"  - For cloud notebooks: Network issues or offline");
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            result.Success = false;
            result.Message = "Access denied: " + ex.Message;
            Console.Error.WriteLine($"✗ UnauthorizedAccessException: {ex.Message}");
        }
        catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x8004201A))
        {
            result.Success = false;
            result.Message = "OneNote Error 0x8004201A: This notebook cannot be exported. " +
                           "Possible causes: Password-protected sections or the notebook is offline.";
            Console.Error.WriteLine($"✗ COM Exception 0x8004201A");
            Console.Error.WriteLine($"✗ Common causes:");
            Console.Error.WriteLine($"  - Notebook contains password-protected sections");
            Console.Error.WriteLine($"  - OneNote has no access to SharePoint/OneDrive");
            Console.Error.WriteLine($"  - The notebook is offline or not synced");
            Console.Error.WriteLine($"");
            Console.Error.WriteLine($"Recommended solutions:");
            Console.Error.WriteLine($"  1. Open the notebook in OneNote Desktop and ensure it's synced");
            Console.Error.WriteLine($"  2. Unlock all password-protected sections before export");
            Console.Error.WriteLine($"  3. Check your network connection for cloud notebooks");
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            result.Success = false;
            result.Message = $"OneNote COM error: 0x{ex.HResult:X} - {ex.Message}";
            Console.Error.WriteLine($"✗ COM Exception: {ex.Message}");
            Console.Error.WriteLine($"✗ HRESULT: 0x{ex.HResult:X}");
            Console.Error.WriteLine($"");

            // Provide specific guidance based on error code
            switch (ex.HResult)
            {
                case unchecked((int)0x80042010):
                    Console.Error.WriteLine($"Note: This error often indicates the notebook is not accessible.");
                    Console.Error.WriteLine($"Open the notebook in OneNote Desktop and ensure it loads correctly.");
                    break;
                case unchecked((int)0x80070005):
                    Console.Error.WriteLine($"Note: Access denied. Check your permissions for this notebook.");
                    break;
                case unchecked((int)0x800706BA):
                    Console.Error.WriteLine($"Note: RPC Server Unavailable (0x800706BA)");
                    Console.Error.WriteLine($"");
                    Console.Error.WriteLine($"This error typically occurs with:");
                    Console.Error.WriteLine($"  1. LARGE notebooks (>100 MB) - OneNote may timeout during PDF/XPS conversion");
                    Console.Error.WriteLine($"  2. Cloud notebooks (SharePoint/OneDrive) - Network connectivity issues");
                    Console.Error.WriteLine($"  3. Notebooks with many embedded objects (images, files, ink)");
                    Console.Error.WriteLine($"");
                    Console.Error.WriteLine($"Solutions:");
                    Console.Error.WriteLine($"  ✓ Use .onepkg format instead (works reliably for large notebooks)");
                    Console.Error.WriteLine($"  ✓ Open the notebook in OneNote Desktop first");
                    Console.Error.WriteLine($"  ✓ Close OneNote completely before export");
                    Console.Error.WriteLine($"  ✓ For cloud notebooks: Ensure stable internet connection");
                    Console.Error.WriteLine($"  ✓ Try exporting sections individually instead of entire notebook");
                    Console.Error.WriteLine($"");
                    Console.Error.WriteLine($"Technical reason: PDF/XPS export requires OneNote to render every page,");
                    Console.Error.WriteLine($"which can take very long for large notebooks and may cause RPC timeout.");
                    break;
                default:
                    Console.Error.WriteLine($"Note: Unknown error code. Try:");
                    Console.Error.WriteLine($"  - Open notebook in OneNote Desktop and ensure it loads");
                    Console.Error.WriteLine($"  - Restart OneNote");
                    Console.Error.WriteLine($"  - For cloud notebooks: Check your network connection");
                    break;
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = "Error during export: " + ex.Message;
            Console.Error.WriteLine($"✗ Exception: {ex.GetType().Name}");
            Console.Error.WriteLine($"✗ Message: {ex.Message}");
            Console.Error.WriteLine($"✗ StackTrace: {ex.StackTrace}");
        }

        return result;
    }

    // Export all notebooks sequentially with live progress updates to stderr
    public ExportResult ExportAllNotebooks(string destinationPath, string exportFormat = "onepkg")
    {
        // Try to close OneNote gracefully before export (only happens once)
        TryCloseOneNoteGracefully();

        var result = new ExportResult { Success = true };
        var exportedCount = 0;
        var failedCount = 0;
        var messages = new List<string>();

        try
        {
            var notebooks = GetNotebooks();
            Console.Error.WriteLine($"=== Starting export of {notebooks.Count} notebook(s) ===");
            Console.Error.WriteLine($"=== Export Format: {exportFormat} ===");

            foreach (var notebook in notebooks)
            {
                Console.Error.WriteLine($"");
                Console.Error.WriteLine($"📓 Exporting notebook {exportedCount + failedCount + 1}/{notebooks.Count}: {notebook.Name}");

                var exportResult = ExportNotebook(notebook.Id, destinationPath, exportFormat);

                if (exportResult.Success)
                {
                    exportedCount++;
                    messages.Add($"✓ {notebook.Name}");
                    Console.Error.WriteLine($"✓ Export successful: {notebook.Name}");
                }
                else
                {
                    failedCount++;
                    messages.Add($"✗ {notebook.Name}: {exportResult.Message}");
                    Console.Error.WriteLine($"✗ Export failed: {notebook.Name}");
                }
            }

            Console.Error.WriteLine($"");
            Console.Error.WriteLine($"=== Export completed: {exportedCount} successful, {failedCount} failed ===");

            result.Success = failedCount == 0;
            result.Message = $"Export abgeschlossen: {exportedCount} erfolgreich, {failedCount} fehlgeschlagen\n" +
                           string.Join("\n", messages);
            result.ExportedPath = destinationPath;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = "Fehler beim Exportieren mehrerer Notizbücher: " + ex.Message;
            Console.Error.WriteLine($"✗ FATAL ERROR: {ex.Message}");
        }

        return result;
    }

    private static string FormatBytes(long bytes)
    {
        string[] sizes = { "Bytes", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }

    private void TryCloseOneNoteGracefully()
    {
        // Only attempt once per session
        if (_oneNoteClosedAttempted)
        {
            return;
        }
        _oneNoteClosedAttempted = true;

        try
        {
            Process[] oneNoteProcesses = Process.GetProcessesByName("ONENOTE");

            if (oneNoteProcesses.Length == 0)
            {
                Console.Error.WriteLine("No OneNote processes found to close.");
                return;
            }

            Console.Error.WriteLine($"INFO: Found {oneNoteProcesses.Length} OneNote process(es). Attempting graceful shutdown before export...");

            foreach (Process process in oneNoteProcesses)
            {
                try
                {
                    // Try to close main window (like clicking X)
                    bool closed = process.CloseMainWindow();

                    if (closed)
                    {
                        Console.Error.WriteLine($"Sent close signal to OneNote (PID: {process.Id})");

                        // Wait up to 5 seconds for graceful shutdown
                        bool exited = process.WaitForExit(5000);

                        if (exited)
                        {
                            Console.Error.WriteLine($"✓ OneNote closed gracefully (PID: {process.Id})");
                        }
                        else
                        {
                            Console.Error.WriteLine($"ℹ OneNote did not close within 5 seconds (PID: {process.Id}), continuing anyway...");
                        }
                    }
                    else
                    {
                        Console.Error.WriteLine($"ℹ Could not close OneNote main window (PID: {process.Id}), continuing anyway...");
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"ℹ Error closing OneNote process (PID: {process.Id}): {ex.Message}");
                }
                finally
                {
                    process.Dispose();
                }
            }

            // Give OneNote a moment to fully shut down
            System.Threading.Thread.Sleep(1000);
            Console.Error.WriteLine("Proceeding with export...");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ℹ Error during OneNote close attempt: {ex.Message}");
            Console.Error.WriteLine("Continuing with export anyway...");
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            if (_oneNote != null)
            {
                try
                {
                    // IMPORTANT: We NEVER close the OneNote process programmatically!
                    // Reason: Even graceful closing can cause OneNote to show the error dialog
                    // "OneNote could not be started on the last attempt" on next start.
                    //
                    // Instead, we just release the COM reference and let OneNote continue running.
                    // The user can manually close OneNote if desired.

                    Console.Error.WriteLine("Releasing COM reference (OneNote continues running in background)");

                    if (!_oneNoteWasRunning)
                    {
                        Console.Error.WriteLine("Note: OneNote was started automatically and is running in the background.");
                        Console.Error.WriteLine("         You can manually close OneNote if desired.");
                    }

                    // Release the COM object properly to avoid keeping OneNote in memory
                    Marshal.ReleaseComObject(_oneNote);
                    _oneNote = null;

                    // Force garbage collection to release COM references
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error cleaning up OneNote connection: {ex.Message}");
                }
            }

            _disposed = true;
        }

        GC.SuppressFinalize(this);
    }
}
