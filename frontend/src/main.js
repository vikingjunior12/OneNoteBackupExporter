import './style.css';
import './app.css';

import { GetNotebooks, ExportNotebook, ExportNotebookNoExplorer, ExportAllNotebooks, GetOneNoteVersion, BrowseFolder, GetDefaultDownloadsPath, CancelExport, ExportBackup, CheckLocalBackupAvailable, OpenFolder } from '../wailsjs/go/main/App';
import { EventsOn, EventsOff } from '../wailsjs/runtime/runtime';

// Initialize the app with a nice UI
document.querySelector('#app').innerHTML = `
    <div class="container">
        <div class="compact-header">
            <h1>OneNote Backup Exporter</h1>
            <div class="version-info" id="version-info">Checking OneNote installation...</div>
        </div>

        <div class="main-content">
            <div class="notebooks-section">
                <div class="section-header">
                    <h2>Available Notebooks:</h2>
                    <button class="btn btn-small" id="refresh-btn">Refresh</button>
                </div>
                <div class="notebook-list" id="notebook-list">Loading notebooks...</div>
            </div>

            <div class="export-section">
                <div class="input-box">
                    <input class="input" id="dest-path" type="text" placeholder="Destination folder (e.g. C:\\Backup)" autocomplete="off" />
                    <button class="btn browse-btn" id="browse-btn">Browse...</button>
                </div>
                <div class="format-selector">
                    <label for="format-select">Export Format:</label>
                    <select class="format-dropdown" id="format-select">
                        <option value="onepkg">OneNote Package (.onepkg) - Recommended, best quality</option>
                        <option value="xps">XPS Document (.xps) - Good layout preservation</option>
                        <option value="pdf">PDF Document (.pdf) - Universal, lower quality</option>
                        <option value="localbackup" id="localbackup-option">Local Backup Copy - Fast, direct file copy</option>
                    </select>
                </div>
                <div class="backup-warning" id="backup-warning" style="display: none;">
                    ⚠️ Local backup copies ALL notebooks (selection will be ignored)
                </div>
                <div class="export-buttons">
                    <button class="btn btn-primary" id="export-selected-btn" disabled>Export Selected</button>
                    <button class="btn btn-danger" id="cancel-btn" style="display: none;">❌ Cancel</button>
                </div>
                <div class="export-status" id="export-status"></div>
                <div class="status" id="status"></div>
            </div>
        </div>
    </div>
`;

// DOM elements
const versionInfoElement = document.getElementById("version-info");
const notebookListElement = document.getElementById("notebook-list");
const destPathElement = document.getElementById("dest-path");
const formatSelect = document.getElementById("format-select");
const backupWarning = document.getElementById("backup-warning");
const localbackupOption = document.getElementById("localbackup-option");
const browseButton = document.getElementById("browse-btn");
const refreshButton = document.getElementById("refresh-btn");
const exportSelectedButton = document.getElementById("export-selected-btn");
const cancelButton = document.getElementById("cancel-btn");
const statusElement = document.getElementById("status");
const exportStatus = document.getElementById("export-status");

let notebooks = [];
let selectedNotebooks = new Set();
let backupAvailability = null;

// Track if listeners are already set up (prevents duplicates on HMR reload)
let listenersInitialized = false;

// CRITICAL: Global lock to prevent multiple simultaneous exports
let exportInProgress = false;

// Initialize the app
document.addEventListener('DOMContentLoaded', async () => {
    await checkOneNoteVersion();
    await checkLocalBackupAvailability();
    await loadNotebooks();

    // Only set up listeners once (prevents memory leaks on HMR reload)
    if (!listenersInitialized) {
        setupNotebookListeners();
        setupBrowseButton();
        setupRefreshButton();
        setupExportButtons();
        setupCancelButton();
        setupFormatChangeListener();
        listenersInitialized = true;
    }

    // Set default destination path
    try {
        const defaultPath = await GetDefaultDownloadsPath();
        destPathElement.value = defaultPath + "\\OneNote-Export";
    } catch (err) {
        console.error("Error getting default path:", err);
    }
});

// Check OneNote version and availability
async function checkOneNoteVersion() {
    try {
        const versionInfo = await GetOneNoteVersion();

        if (versionInfo.oneNoteInstalled) {
            versionInfoElement.textContent = `✓ ${versionInfo.oneNoteVersion}`;
            versionInfoElement.className = "version-info success";
        } else {
            versionInfoElement.textContent = "⚠ OneNote Desktop not found";
            versionInfoElement.className = "version-info warning";
        }
    } catch (err) {
        console.error("Error checking OneNote version:", err);
        versionInfoElement.textContent = "⚠ " + err.message;
        versionInfoElement.className = "version-info error";
    }
}

// Check if local OneNote backups are available
async function checkLocalBackupAvailability() {
    try {
        backupAvailability = await CheckLocalBackupAvailable();

        if (backupAvailability.available) {
            // Update the option text to show info
            localbackupOption.textContent = `Local Backup Copy (${backupAvailability.notebookCount} notebooks)`;
            localbackupOption.disabled = false;
            console.log("Local backup available:", backupAvailability.message);
        } else {
            // Disable the option and show why it's not available
            localbackupOption.textContent = `Local Backup Copy - Not available (${backupAvailability.reason})`;
            localbackupOption.disabled = true;
            console.warn("Local backup not available:", backupAvailability.message);
        }
    } catch (err) {
        console.error("Error checking local backup availability:", err);
        localbackupOption.textContent = "Local Backup Copy - Error checking availability";
        localbackupOption.disabled = true;
    }
}

// Load notebooks from OneNote
async function loadNotebooks() {
    try {
        notebookListElement.innerHTML = '<div class="loading">Loading notebooks...</div>';

        notebooks = await GetNotebooks();

        if (notebooks && notebooks.length > 0) {
            renderNotebooks(notebooks);
        } else {
            notebookListElement.innerHTML = "<div class='no-notebooks'>No notebooks found</div>";
        }
    } catch (err) {
        console.error(err);
        notebookListElement.innerHTML = `<div class='error'>Error loading notebooks: ${err.message}</div>`;
    }
}

// Render the notebooks list
function renderNotebooks(notebooks) {
    let html = '<div class="notebook-items">';

    notebooks.forEach(notebook => {
        const isViewed = notebook.isCurrentlyViewed ? ' (currently open)' : '';
        const lastModified = notebook.lastModified ? formatDate(notebook.lastModified) : 'Unknown';

        html += `
            <div class="notebook-item" data-id="${notebook.id}">
                <label class="notebook-checkbox">
                    <input type="checkbox" value="${notebook.id}" class="notebook-check">
                    <div class="notebook-info">
                        <div class="notebook-name">📓 ${notebook.name}${isViewed}</div>
                        <div class="notebook-meta">Last modified: ${lastModified}</div>
                    </div>
                </label>
            </div>
        `;
    });

    html += '</div>';
    notebookListElement.innerHTML = html;

    // NOTE: Event listeners are now set up once via event delegation (see setupNotebookListeners)
    // This prevents memory leaks from stacking listeners on every render
}

// Update export button state based on selection
function updateExportButtonState() {
    exportSelectedButton.disabled = selectedNotebooks.size === 0;
    exportSelectedButton.textContent = selectedNotebooks.size > 0
        ? `Export ${selectedNotebooks.size} notebook(s)`
        : 'Export Selected';
}

// Format date string
function formatDate(dateString) {
    try {
        const date = new Date(dateString);
        return date.toLocaleDateString('de-DE', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return dateString;
    }
}

// Setup notebook listeners using event delegation (prevents memory leaks)
function setupNotebookListeners() {
    // Use event delegation: listen on parent element instead of individual checkboxes
    // This way we only have ONE listener regardless of how many notebooks exist
    notebookListElement.addEventListener('change', (e) => {
        if (e.target.classList.contains('notebook-check')) {
            const notebookId = e.target.value;
            if (e.target.checked) {
                selectedNotebooks.add(notebookId);
            } else {
                selectedNotebooks.delete(notebookId);
            }
            updateExportButtonState();
        }
    });
}

// Setup the browse button
function setupBrowseButton() {
    browseButton.addEventListener('click', async () => {
        try {
            const selectedDir = await BrowseFolder();

            if (selectedDir) {
                destPathElement.value = selectedDir;
                statusElement.textContent = "Path selected: " + selectedDir;
                statusElement.className = "status info";
            }
        } catch (err) {
            console.error("Error opening folder selection dialog:", err);
            statusElement.textContent = "Error opening folder selection dialog";
            statusElement.className = "status error";
        }
    });
}

// Setup the refresh button
function setupRefreshButton() {
    refreshButton.addEventListener('click', async () => {
        selectedNotebooks.clear();
        await loadNotebooks();
        updateExportButtonState();
    });
}

// Setup export buttons
function setupExportButtons() {
    // Export selected notebooks
    exportSelectedButton.addEventListener('click', async () => {
        // CRITICAL: Prevent multiple simultaneous exports
        if (exportInProgress) {
            statusElement.textContent = "⚠ An export is already running! Please wait until it is completed.";
            statusElement.className = "status warning";
            return;
        }

        const destPath = destPathElement.value.trim();

        if (!destPath) {
            statusElement.textContent = "Please specify a destination folder";
            statusElement.className = "status error";
            return;
        }

        if (selectedNotebooks.size === 0) {
            statusElement.textContent = "Please select at least one notebook";
            statusElement.className = "status error";
            return;
        }

        // Set the global lock
        exportInProgress = true;
        cancelButton.style.display = "inline-block"; // Show cancel button

        // Clear old status messages
        statusElement.textContent = "";
        statusElement.className = "status";

        await exportNotebooks(Array.from(selectedNotebooks), destPath);
    });
}

// Setup cancel button
function setupCancelButton() {
    cancelButton.addEventListener('click', async () => {
        if (!exportInProgress) {
            return; // Should not happen, button should be hidden
        }

        // Ask for confirmation
        if (!confirm("Do you really want to cancel the export?\n\nWarning: This will terminate OneNote and OneNoteHelper!")) {
            return;
        }

        console.log("[Cancel] User requested cancellation");
        statusElement.textContent = "Cancelling export...";
        statusElement.className = "status warning";

        try {
            const result = await CancelExport();
            console.log("[Cancel] ", result.message);
        } catch (err) {
            console.error("[Cancel] Error: ", err);
            statusElement.textContent = "Error during cancellation: " + err.message;
            statusElement.className = "status error";
        }
    });
}

// Setup format change listener for smart UI behavior
function setupFormatChangeListener() {
    formatSelect.addEventListener('change', () => {
        const isLocalBackup = formatSelect.value === 'localbackup';

        if (isLocalBackup) {
            // Show warning
            backupWarning.style.display = 'block';

            // Auto-check all notebooks and disable checkboxes
            const checkboxes = document.querySelectorAll('.notebook-check');
            checkboxes.forEach(checkbox => {
                checkbox.checked = true;
                checkbox.disabled = true;
                selectedNotebooks.add(checkbox.value);
            });

            // Gray out notebook list to indicate it's disabled
            notebookListElement.style.opacity = '0.6';
            notebookListElement.style.pointerEvents = 'none';

            // Update button text
            if (backupAvailability && backupAvailability.available) {
                exportSelectedButton.textContent = `Export ALL Notebooks (${backupAvailability.notebookCount} total)`;
            } else {
                exportSelectedButton.textContent = 'Export ALL Notebooks';
            }
            exportSelectedButton.disabled = false;
        } else {
            // Hide warning
            backupWarning.style.display = 'none';

            // Re-enable checkboxes and clear selection
            const checkboxes = document.querySelectorAll('.notebook-check');
            checkboxes.forEach(checkbox => {
                checkbox.checked = false;
                checkbox.disabled = false;
            });
            selectedNotebooks.clear();

            // Restore notebook list appearance
            notebookListElement.style.opacity = '1';
            notebookListElement.style.pointerEvents = 'auto';

            // Reset button
            updateExportButtonState();
        }
    });
}

// Export selected notebooks one by one
async function exportNotebooks(notebookIds, destPath) {
    try {
        disableButtons(true);

        // Clear old status and reset progress display
        statusElement.textContent = "";
        statusElement.className = "status";
        exportStatus.innerHTML = '<span class="spinner"></span>Starting export...';

        // Get selected format
        const format = formatSelect.value;

        // Special handling for localbackup format
        if (format === 'localbackup') {
            // Double-check availability before export
            if (!backupAvailability || !backupAvailability.available) {
                statusElement.textContent = "❌ Local backup not available: " + (backupAvailability ? backupAvailability.message : "Unknown error");
                statusElement.className = "status error";
                exportInProgress = false;
                cancelButton.style.display = "none";
                disableButtons(false);
                hideProgress();
                return;
            }

            exportStatus.innerHTML = '<span class="spinner"></span>Copying local backup files...';

            try {
                const result = await ExportBackup(destPath);

                hideProgress();

                if (result.success) {
                    statusElement.textContent = `✓ Local backup successfully copied!\n\nExported to: ${destPath}`;
                    statusElement.className = "status success";
                } else {
                    statusElement.textContent = `❌ Export failed: ${result.message}`;
                    statusElement.className = "status error";
                }
            } catch (err) {
                hideProgress();
                statusElement.textContent = `❌ Error during export: ${err.message}`;
                statusElement.className = "status error";
            }

            exportInProgress = false;
            cancelButton.style.display = "none";
            disableButtons(false);
            return;
        }

        // Regular COM-based export (onepkg, xps, pdf)
        let successCount = 0;
        let failCount = 0;
        const messages = [];

        for (let i = 0; i < notebookIds.length; i++) {
            const notebookId = notebookIds[i];
            const notebook = notebooks.find(nb => nb.id === notebookId);
            const notebookName = notebook ? notebook.name : 'Unknown';

            // Show current status
            exportStatus.innerHTML = `<span class="spinner"></span>Exporting notebook ${i + 1}/${notebookIds.length}: ${notebookName}\nOneNote is writing in the background, this may take several minutes...`;

            console.log(`[DEBUG] Starting export of notebook ${i + 1}/${notebookIds.length}`);

            try {
                // Use ExportNotebookNoExplorer to prevent opening Explorer for each notebook
                const result = await ExportNotebookNoExplorer(notebookId, destPath, format);

                if (result.success) {
                    successCount++;
                    messages.push(`✓ ${notebookName}`);
                } else {
                    failCount++;
                    messages.push(`✗ ${notebookName}: ${result.message}`);
                }
            } catch (err) {
                failCount++;
                messages.push(`✗ ${notebookName}: ${err.message}`);
            }

            // Show completed status for this notebook
            exportStatus.innerHTML = `✓ Completed: ${i + 1}/${notebookIds.length} notebooks`;

            console.log(`[DEBUG] Completed notebook ${i + 1}/${notebookIds.length}`);

            // Brief pause to show the status update
            await new Promise(resolve => setTimeout(resolve, 500));
        }

        hideProgress();

        // Show final result
        const finalMessage = `Export completed: ${successCount} successful, ${failCount} failed\n\n${messages.join('\n')}`;
        statusElement.textContent = finalMessage;
        statusElement.className = failCount === 0 ? "status success" : "status warning";

        // Open Explorer once after all exports are complete (only if at least one was successful)
        if (successCount > 0) {
            try {
                await OpenFolder(destPath);
            } catch (err) {
                console.error('Failed to open folder:', err);
            }
        }

        // Release the global lock and hide cancel button
        exportInProgress = false;
        cancelButton.style.display = "none";

        disableButtons(false);

    } catch (err) {
        hideProgress();
        console.error(err);
        statusElement.textContent = "Error during export: " + err.message;
        statusElement.className = "status error";

        // Release the global lock and hide cancel button
        exportInProgress = false;
        cancelButton.style.display = "none";

        disableButtons(false);
    }
}

// Show status message
function showProgress(message) {
    exportStatus.innerHTML = `<span class="spinner"></span>${message}`;
    statusElement.textContent = "";
}

// Show completion message
function showCompletion(message) {
    exportStatus.innerHTML = `✓ ${message}`;
    statusElement.textContent = "";
}

// Hide status message
function hideProgress() {
    exportStatus.textContent = "";
}

// Disable/enable buttons during export
function disableButtons(disabled) {
    exportSelectedButton.disabled = disabled;
    refreshButton.disabled = disabled;
    browseButton.disabled = disabled;
}

// Vite HMR cleanup (prevents memory leaks in dev mode)
if (import.meta.hot) {
    import.meta.hot.dispose(() => {
        // Clean up before HMR reload
        console.log('[HMR] Cleaning up before reload...');
        listenersInitialized = false;
        selectedNotebooks.clear();
    });
}
