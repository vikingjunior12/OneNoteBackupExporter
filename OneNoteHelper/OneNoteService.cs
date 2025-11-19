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

    public OneNoteService()
    {
        try
        {
            // Check if OneNote is already running
            var oneNoteProcesses = Process.GetProcessesByName("ONENOTE");
            _oneNoteWasRunning = oneNoteProcesses.Length > 0;

            if (_oneNoteWasRunning)
            {
                Console.Error.WriteLine("WARNUNG: OneNote läuft bereits. Dies kann zu Problemen führen.");
                Console.Error.WriteLine("Empfehlung: Schließen Sie OneNote und starten Sie den Export erneut.");

                // Try to connect anyway, but with awareness that it might hang
                Console.Error.WriteLine("Versuche trotzdem, Verbindung herzustellen...");
            }

            _oneNote = new Application();
            Console.Error.WriteLine("OneNote COM-Verbindung erfolgreich hergestellt.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "OneNote Desktop konnte nicht initialisiert werden. " +
                "Bitte stellen Sie sicher, dass OneNote 2016 installiert ist.", ex);
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
                info.OneNoteVersion = "OneNote Desktop (COM API verfügbar)";
            }
            catch
            {
                info.OneNoteVersion = "Unbekannt";
            }
        }

        return info;
    }

    public List<NotebookInfo> GetNotebooks()
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote ist nicht initialisiert");

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
                    Name = nb.Attribute("name")?.Value ?? "Unbenannt",
                    Path = nb.Attribute("path")?.Value ?? "",
                    LastModified = nb.Attribute("lastModifiedTime")?.Value ?? "",
                    IsCurrentlyViewed = nb.Attribute("isCurrentlyViewed")?.Value == "true"
                };

                notebooks.Add(notebook);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Fehler beim Abrufen der Notizbücher: " + ex.Message, ex);
        }

        return notebooks;
    }

    public ExportResult ExportNotebook(string notebookId, string destinationPath)
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote ist nicht initialisiert");

        var result = new ExportResult();

        try
        {
            // Get notebook info for naming
            _oneNote.GetHierarchy(notebookId, HierarchyScope.hsSelf, out string xml);
            var xdoc = XDocument.Parse(xml);
            var ns = xdoc.Root?.Name.Namespace;
            var notebookName = xdoc.Root?.Attribute("name")?.Value ?? "Notizbuch";
            var notebookPath = xdoc.Root?.Attribute("path")?.Value ?? "";

            Console.Error.WriteLine($"=== Exportiere Notizbuch ===");
            Console.Error.WriteLine($"Name: {notebookName}");
            Console.Error.WriteLine($"ID: {notebookId}");
            Console.Error.WriteLine($"Pfad: {notebookPath}");

            // Check if this is a SharePoint/OneDrive notebook (starts with https://)
            bool isCloudNotebook = notebookPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                                   notebookPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase);

            if (isCloudNotebook)
            {
                Console.Error.WriteLine($"ℹ INFO: Dies ist ein Cloud-Notizbuch (SharePoint/OneDrive)");
                Console.Error.WriteLine($"ℹ Versuche .onepkg-Export mit vollständiger Synchronisierung...");
            }

            // Sanitize filename
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitizedName = string.Join("_", notebookName.Split(invalidChars));
            var fileName = $"{sanitizedName}.onepkg";
            var fullPath = Path.Combine(destinationPath, fileName);

            // Ensure destination directory exists
            Directory.CreateDirectory(destinationPath);
            Console.Error.WriteLine($"Zielverzeichnis: {destinationPath}");
            Console.Error.WriteLine($"Zieldatei: {fullPath}");

            // Step 1: Ensure notebook is open (critical for SharePoint/OneDrive notebooks)
            Console.Error.WriteLine($"Öffne Notizbuch-Hierarchie...");
            _oneNote.OpenHierarchy(
                notebookPath,
                "",
                out string openedNotebookId,
                CreateFileType.cftNone
            );
            Console.Error.WriteLine($"Geöffnete Notebook-ID: {openedNotebookId}");

            // Step 2: Force full synchronization - critical for cloud notebooks!
            // Note: One sync is usually enough if the notebook was already opened in OneNote Desktop
            Console.Error.WriteLine($"Synchronisiere Notizbuch{(isCloudNotebook ? " (Cloud-Notizbuch)" : "")}...");

            _oneNote.SyncHierarchy(openedNotebookId);

            // Wait for sync to complete (longer for cloud notebooks)
            int syncWaitMs = isCloudNotebook ? 5000 : 2000;
            Console.Error.WriteLine($"Warte {syncWaitMs / 1000} Sekunden auf Synchronisierung...");
            System.Threading.Thread.Sleep(syncWaitMs);

            Console.Error.WriteLine("✓ Synchronisierung abgeschlossen");

            // Step 3: Publish to .onepkg format using the opened notebook ID
            Console.Error.WriteLine($"Starte Export-Vorgang...");
            Console.Error.WriteLine($"Hinweis: OneNote schreibt die Datei asynchron im Hintergrund");

            try
            {
                _oneNote.Publish(
                    openedNotebookId,  // Use the opened ID, not the original ID
                    fullPath,
                    PublishFormat.pfOneNotePackage,
                    ""
                );

                Console.Error.WriteLine("✓ Publish()-Aufruf erfolgreich (OneNote schreibt jetzt im Hintergrund)");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.Error.WriteLine($"✗ Publish() fehlgeschlagen: 0x{ex.HResult:X}");

                // Retry once with additional sync for cloud notebooks
                if (isCloudNotebook)
                {
                    Console.Error.WriteLine($"Versuche erneut nach zusätzlicher Synchronisierung...");
                    _oneNote.SyncHierarchy(openedNotebookId);
                    System.Threading.Thread.Sleep(5000);

                    _oneNote.Publish(
                        openedNotebookId,
                        fullPath,
                        PublishFormat.pfOneNotePackage,
                        ""
                    );

                    Console.Error.WriteLine("✓ Publish()-Aufruf erfolgreich (2. Versuch)");
                }
                else
                {
                    throw; // Rethrow for local notebooks
                }
            }

            // Wait for the publish operation to complete
            // IMPORTANT: Publish() returns immediately but OneNote writes the file asynchronously!
            Console.Error.WriteLine("Warte auf Abschluss des Export-Vorgangs (OneNote schreibt asynchron)...");
            System.Threading.Thread.Sleep(2000);

            // Wait for file creation AND for stable file size (OneNote writes in background)
            // Large notebooks (100+ MB) can take several minutes to write!
            bool fileCreated = false;
            long fileSize = 0;
            long previousSize = -1;
            int stableCount = 0;
            int maxAttempts = isCloudNotebook ? 150 : 90; // Cloud: up to 5 minutes, Local: up to 3 minutes
            int checkIntervalMs = 2000;

            Console.Error.WriteLine($"Überwache Datei-Erstellung und Schreibfortschritt: {fullPath}");
            Console.Error.WriteLine($"Hinweis: Große Notizbücher (100+ MB) können mehrere Minuten benötigen...");

            for (int i = 0; i < maxAttempts; i++)
            {
                if (File.Exists(fullPath))
                {
                    fileCreated = true;
                    fileSize = new FileInfo(fullPath).Length;

                    // Check if file size is stable (not changing = OneNote finished writing)
                    if (fileSize > 0 && fileSize == previousSize)
                    {
                        stableCount++;
                        if (stableCount >= 2)
                        {
                            // Size hasn't changed for 2 checks (4 seconds) and is > 0 = done!
                            Console.Error.WriteLine($"✓ Datei vollständig geschrieben! Finale Größe: {FormatBytes(fileSize)}");
                            break;
                        }
                        Console.Error.WriteLine($"Dateigröße stabil bei {FormatBytes(fileSize)} (Bestätigung {stableCount}/2)");
                    }
                    else if (fileSize > 0)
                    {
                        // Show progress
                        if (previousSize > 0)
                        {
                            double mbPerSec = (fileSize - previousSize) / 1024.0 / 1024.0 / (checkIntervalMs / 1000.0);
                            Console.Error.WriteLine($"Datei wächst: {FormatBytes(fileSize)} (+{FormatBytes(fileSize - previousSize)}, ~{mbPerSec:0.0} MB/s)");
                        }
                        else
                        {
                            Console.Error.WriteLine($"Datei wächst: {FormatBytes(fileSize)}");
                        }
                        stableCount = 0; // Reset counter, file is still growing
                    }
                    else
                    {
                        Console.Error.WriteLine($"Datei existiert, ist aber noch leer (0 Bytes) - OneNote schreibt noch...");
                        stableCount = 0;
                    }

                    previousSize = fileSize;
                }
                else
                {
                    // File doesn't exist yet
                    if (i == 0)
                    {
                        Console.Error.WriteLine($"Datei noch nicht erstellt, warte...");
                        Console.Error.WriteLine($"Hinweis: OneNote arbeitet im Hintergrund, dies kann bis zu 1 Minute dauern...");
                    }
                    else if (i % 10 == 0)
                    {
                        int secondsElapsed = (i * checkIntervalMs) / 1000;
                        Console.Error.WriteLine($"Warte weiterhin auf Datei-Erstellung... ({secondsElapsed}s vergangen)");
                    }
                }

                System.Threading.Thread.Sleep(checkIntervalMs);
            }

            if (fileCreated && fileSize > 0)
            {
                result.Success = true;
                result.Message = $"Notizbuch '{notebookName}' erfolgreich exportiert ({FormatBytes(fileSize)})";
                result.ExportedPath = fullPath;
                Console.Error.WriteLine($"✓ Export erfolgreich: {fullPath} ({FormatBytes(fileSize)})");
            }
            else if (fileCreated && fileSize == 0)
            {
                result.Success = false;
                result.Message = $"Export fehlgeschlagen: Datei wurde erstellt, ist aber leer (0 Bytes). " +
                               $"Dies deutet auf ein Problem während des Publish()-Vorgangs hin.";
                Console.Error.WriteLine($"✗ Datei existiert, ist aber leer!");
                Console.Error.WriteLine($"✗ OneNote hat möglicherweise den Export abgebrochen");
            }
            else
            {
                int timeoutMinutes = (maxAttempts * checkIntervalMs) / 60000;
                result.Success = false;
                result.Message = $"Timeout: Datei wurde nach {timeoutMinutes} Minuten nicht erstellt. " +
                               $"WICHTIG: OneNote arbeitet möglicherweise noch im Hintergrund! " +
                               $"Prüfen Sie {destinationPath} in einigen Minuten erneut.";

                Console.Error.WriteLine($"✗ Timeout nach {timeoutMinutes} Minuten erreicht");
                Console.Error.WriteLine($"✗ Datei existiert nicht: {fullPath}");
                Console.Error.WriteLine($"");
                Console.Error.WriteLine($"⚠ WICHTIG:");
                Console.Error.WriteLine($"  OneNote schreibt die Datei möglicherweise noch im Hintergrund!");
                Console.Error.WriteLine($"  Der Export-Befehl wurde erfolgreich an OneNote gesendet.");
                Console.Error.WriteLine($"  Bei sehr großen Notizbüchern (>500 MB) kann der Schreibvorgang");
                Console.Error.WriteLine($"  länger als {timeoutMinutes} Minuten dauern.");
                Console.Error.WriteLine($"");
                Console.Error.WriteLine($"Empfehlung:");
                Console.Error.WriteLine($"  1. Lassen Sie das Programm NICHT neu starten (würde neuen Export starten)");
                Console.Error.WriteLine($"  2. Prüfen Sie das Verzeichnis {destinationPath}");
                Console.Error.WriteLine($"     in 5-10 Minuten erneut");
                Console.Error.WriteLine($"  3. Überwachen Sie die OneNote.exe im Task-Manager (hohe CPU/Disk = arbeitet noch)");
                Console.Error.WriteLine($"");
                Console.Error.WriteLine($"Falls nach 15 Minuten keine Datei vorhanden ist, mögliche Ursachen:");
                Console.Error.WriteLine($"  - Passwortgeschützte Bereiche im Notizbuch");
                Console.Error.WriteLine($"  - Notizbuch nicht vollständig synchronisiert");
                Console.Error.WriteLine($"  - OneNote-Prozess hat keine Schreibrechte auf Zielverzeichnis");
                Console.Error.WriteLine($"  - Bei SharePoint: Notizbuch ist offline oder nicht verfügbar");
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            result.Success = false;
            result.Message = "Zugriff verweigert: " + ex.Message;
            Console.Error.WriteLine($"✗ UnauthorizedAccessException: {ex.Message}");
        }
        catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x8004201A))
        {
            result.Success = false;
            result.Message = "OneNote Fehler 0x8004201A: Dieses Notizbuch kann nicht exportiert werden. " +
                           "Mögliche Ursachen: Passwortgeschützte Bereiche, fehlende Synchronisierung, oder das Notizbuch ist offline.";
            Console.Error.WriteLine($"✗ COM Exception 0x8004201A");
            Console.Error.WriteLine($"✗ Häufige Ursachen bei SharePoint-Notizbüchern:");
            Console.Error.WriteLine($"  - Notizbuch enthält passwortgeschützte Bereiche");
            Console.Error.WriteLine($"  - Notizbuch ist nicht vollständig synchronisiert");
            Console.Error.WriteLine($"  - OneNote hat keinen Zugriff auf SharePoint");
            Console.Error.WriteLine($"  - Das Notizbuch ist offline");
            Console.Error.WriteLine($"");
            Console.Error.WriteLine($"Empfohlene Lösungen:");
            Console.Error.WriteLine($"  1. Öffnen Sie das Notizbuch in OneNote Desktop und warten Sie, bis 'Alle Änderungen synchronisiert' erscheint");
            Console.Error.WriteLine($"  2. Entsperren Sie alle passwortgeschützten Abschnitte vor dem Export");
            Console.Error.WriteLine($"  3. Lösen Sie alle Synchronisierungskonflikte im Notizbuch");
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            result.Success = false;
            result.Message = $"OneNote COM-Fehler: 0x{ex.HResult:X} - {ex.Message}";
            Console.Error.WriteLine($"✗ COM Exception: {ex.Message}");
            Console.Error.WriteLine($"✗ HRESULT: 0x{ex.HResult:X}");
            Console.Error.WriteLine($"");

            // Provide specific guidance based on error code
            switch (ex.HResult)
            {
                case unchecked((int)0x80042010):
                    Console.Error.WriteLine($"Hinweis: Dieser Fehler deutet oft auf Synchronisierungsprobleme hin.");
                    Console.Error.WriteLine($"Öffnen Sie das Notizbuch in OneNote Desktop und warten Sie auf vollständige Synchronisierung.");
                    break;
                case unchecked((int)0x80070005):
                    Console.Error.WriteLine($"Hinweis: Zugriff verweigert. Überprüfen Sie Ihre Berechtigungen für dieses Notizbuch.");
                    break;
                default:
                    Console.Error.WriteLine($"Hinweis: Unbekannter Fehlercode. Versuchen Sie:");
                    Console.Error.WriteLine($"  - Notizbuch in OneNote Desktop öffnen und manuell synchronisieren");
                    Console.Error.WriteLine($"  - OneNote neu starten");
                    Console.Error.WriteLine($"  - Bei SharePoint-Notizbüchern: Prüfen Sie Ihre Netzwerkverbindung");
                    break;
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = "Fehler beim Exportieren: " + ex.Message;
            Console.Error.WriteLine($"✗ Exception: {ex.GetType().Name}");
            Console.Error.WriteLine($"✗ Message: {ex.Message}");
            Console.Error.WriteLine($"✗ StackTrace: {ex.StackTrace}");
        }

        return result;
    }

    public ExportResult ExportAllNotebooks(string destinationPath)
    {
        var result = new ExportResult { Success = true };
        var exportedCount = 0;
        var failedCount = 0;
        var messages = new List<string>();

        try
        {
            var notebooks = GetNotebooks();

            foreach (var notebook in notebooks)
            {
                Console.Error.WriteLine($"\nExportiere Notizbuch {exportedCount + 1}/{notebooks.Count}: {notebook.Name}");

                var exportResult = ExportNotebook(notebook.Id, destinationPath);

                if (exportResult.Success)
                {
                    exportedCount++;
                    messages.Add($"✓ {notebook.Name}");
                }
                else
                {
                    failedCount++;
                    messages.Add($"✗ {notebook.Name}: {exportResult.Message}");
                }
            }

            result.Success = failedCount == 0;
            result.Message = $"Export abgeschlossen: {exportedCount} erfolgreich, {failedCount} fehlgeschlagen\n" +
                           string.Join("\n", messages);
            result.ExportedPath = destinationPath;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = "Fehler beim Exportieren mehrerer Notizbücher: " + ex.Message;
        }

        return result;
    }

    private ExportResult ExportCloudNotebookAlternative(string notebookId, string notebookName, string destinationPath)
    {
        var result = new ExportResult();

        try
        {
            Console.Error.WriteLine($"=== Alternative Export-Methode für Cloud-Notizbuch ===");

            // Sanitize filename
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitizedName = string.Join("_", notebookName.Split(invalidChars));

            // Create directory for this notebook
            var notebookDir = Path.Combine(destinationPath, sanitizedName);
            Directory.CreateDirectory(notebookDir);

            Console.Error.WriteLine($"Exportiere als PDF (Cloud-Notebooks können nicht als .onepkg exportiert werden)");

            // Get full hierarchy with all pages
            _oneNote.GetHierarchy(notebookId, HierarchyScope.hsPages, out string xml);
            var xdoc = XDocument.Parse(xml);
            var ns = xdoc.Root?.Name.Namespace;

            if (ns == null)
            {
                result.Success = false;
                result.Message = "Fehler beim Parsen der Notizbuch-Hierarchie";
                return result;
            }

            int pageCount = 0;
            int successCount = 0;
            int failCount = 0;

            // Export each section
            var sections = xdoc.Descendants(ns + "Section");
            foreach (var section in sections)
            {
                var sectionName = section.Attribute("name")?.Value ?? "Unbenannt";
                var sanitizedSectionName = string.Join("_", sectionName.Split(invalidChars));
                var sectionDir = Path.Combine(notebookDir, sanitizedSectionName);
                Directory.CreateDirectory(sectionDir);

                Console.Error.WriteLine($"  Exportiere Abschnitt: {sectionName}");

                // Export each page in this section
                var pages = section.Descendants(ns + "Page");
                foreach (var page in pages)
                {
                    pageCount++;
                    var pageId = page.Attribute("ID")?.Value;
                    var pageName = page.Attribute("name")?.Value ?? $"Seite{pageCount}";
                    var sanitizedPageName = string.Join("_", pageName.Split(invalidChars));

                    if (string.IsNullOrEmpty(pageId))
                        continue;

                    try
                    {
                        var pdfPath = Path.Combine(sectionDir, $"{sanitizedPageName}.pdf");

                        // Try to export as PDF
                        _oneNote.Publish(pageId, pdfPath, PublishFormat.pfPDF, "");

                        if (File.Exists(pdfPath))
                        {
                            successCount++;
                            Console.Error.WriteLine($"    ✓ {pageName}");
                        }
                        else
                        {
                            failCount++;
                            Console.Error.WriteLine($"    ✗ {pageName} (Datei nicht erstellt)");
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        Console.Error.WriteLine($"    ✗ {pageName}: {ex.Message}");
                    }
                }
            }

            result.Success = failCount == 0;
            result.Message = $"Cloud-Notizbuch als PDF exportiert: {successCount} Seiten erfolgreich, {failCount} fehlgeschlagen\n" +
                           $"⚠ Hinweis: Cloud-Notizbücher (SharePoint/OneDrive) können nicht als .onepkg exportiert werden.\n" +
                           $"Alternative: Speichern Sie das Notizbuch in OneNote Desktop als lokales Notizbuch.";
            result.ExportedPath = notebookDir;

            Console.Error.WriteLine($"Export abgeschlossen: {successCount}/{pageCount} Seiten");
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = $"Fehler beim alternativen Export: {ex.Message}";
            Console.Error.WriteLine($"✗ Fehler: {ex.Message}");
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
                    // "OneNote konnte beim letzten Versuch nicht gestartet werden" on next start.
                    //
                    // Instead, we just release the COM reference and let OneNote continue running.
                    // The user can manually close OneNote if desired.

                    Console.Error.WriteLine("Gebe COM-Referenz frei (OneNote läuft weiter im Hintergrund)");

                    if (!_oneNoteWasRunning)
                    {
                        Console.Error.WriteLine("Hinweis: OneNote wurde automatisch gestartet und läuft im Hintergrund.");
                        Console.Error.WriteLine("         Sie können OneNote manuell schließen wenn gewünscht.");
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
                    Console.Error.WriteLine($"Fehler beim Aufräumen der OneNote-Verbindung: {ex.Message}");
                }
            }

            _disposed = true;
        }

        GC.SuppressFinalize(this);
    }
}
