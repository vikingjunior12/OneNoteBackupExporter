import './style.css';
import './app.css';

import { GetNotebooks, ExportNotebook, ExportAllNotebooks, GetOneNoteVersion, BrowseFolder, GetDefaultDownloadsPath, CancelExport } from '../wailsjs/go/main/App';
import { EventsOn, EventsOff } from '../wailsjs/runtime/runtime';

// Initialize the app with a nice UI
document.querySelector('#app').innerHTML = `
    <div class="container">
        <div class="compact-header">
            <h1>OneNote Backup Exporter</h1>
            <div class="version-info" id="version-info">Prüfe OneNote Installation...</div>
        </div>

        <div class="main-content">
            <div class="notebooks-section">
                <div class="section-header">
                    <h2>Verfügbare Notizbücher:</h2>
                    <button class="btn btn-small" id="refresh-btn">Aktualisieren</button>
                </div>
                <div class="notebook-list" id="notebook-list">Lade Notizbücher...</div>
            </div>

            <div class="export-section">
                <div class="input-box">
                    <input class="input" id="dest-path" type="text" placeholder="Zielordner (z.B. C:\\Backup)" autocomplete="off" />
                    <button class="btn browse-btn" id="browse-btn">Durchsuchen...</button>
                </div>
                <div class="export-buttons">
                    <button class="btn btn-primary" id="export-selected-btn" disabled>Ausgewählte exportieren</button>
                    <button class="btn btn-secondary" id="export-all-btn">Alle exportieren</button>
                    <button class="btn btn-danger" id="cancel-btn" style="display: none;">❌ Abbrechen</button>
                </div>
                <div class="progress-container" id="progress-container" style="display: none;">
                    <div class="progress-bar-container">
                        <div class="progress-bar" id="progress-bar"></div>
                        <div class="progress-percent" id="progress-percent">0%</div>
                    </div>
                    <div class="progress-text" id="progress-text"></div>
                </div>
                <div class="status" id="status"></div>
            </div>
        </div>
    </div>
`;

// DOM elements
const versionInfoElement = document.getElementById("version-info");
const notebookListElement = document.getElementById("notebook-list");
const destPathElement = document.getElementById("dest-path");
const browseButton = document.getElementById("browse-btn");
const refreshButton = document.getElementById("refresh-btn");
const exportSelectedButton = document.getElementById("export-selected-btn");
const exportAllButton = document.getElementById("export-all-btn");
const cancelButton = document.getElementById("cancel-btn");
const statusElement = document.getElementById("status");
const progressContainer = document.getElementById("progress-container");
const progressBar = document.getElementById("progress-bar");
const progressPercent = document.getElementById("progress-percent");
const progressText = document.getElementById("progress-text");

let notebooks = [];
let selectedNotebooks = new Set();

// Track if listeners are already set up (prevents duplicates on HMR reload)
let listenersInitialized = false;

// CRITICAL: Global lock to prevent multiple simultaneous exports
let exportInProgress = false;

// Initialize the app
document.addEventListener('DOMContentLoaded', async () => {
    await checkOneNoteVersion();
    await loadNotebooks();

    // Only set up listeners once (prevents memory leaks on HMR reload)
    if (!listenersInitialized) {
        setupNotebookListeners();
        setupBrowseButton();
        setupRefreshButton();
        setupExportButtons();
        setupCancelButton();
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
            versionInfoElement.textContent = "⚠ OneNote Desktop nicht gefunden";
            versionInfoElement.className = "version-info warning";
        }
    } catch (err) {
        console.error("Error checking OneNote version:", err);
        versionInfoElement.textContent = "⚠ " + err.message;
        versionInfoElement.className = "version-info error";
    }
}

// Load notebooks from OneNote
async function loadNotebooks() {
    try {
        notebookListElement.innerHTML = '<div class="loading">Lade Notizbücher...</div>';

        notebooks = await GetNotebooks();

        if (notebooks && notebooks.length > 0) {
            renderNotebooks(notebooks);
        } else {
            notebookListElement.innerHTML = "<div class='no-notebooks'>Keine Notizbücher gefunden</div>";
        }
    } catch (err) {
        console.error(err);
        notebookListElement.innerHTML = `<div class='error'>Fehler beim Laden der Notizbücher: ${err.message}</div>`;
    }
}

// Render the notebooks list
function renderNotebooks(notebooks) {
    let html = '<div class="notebook-items">';

    notebooks.forEach(notebook => {
        const isViewed = notebook.isCurrentlyViewed ? ' (aktuell geöffnet)' : '';
        const lastModified = notebook.lastModified ? formatDate(notebook.lastModified) : 'Unbekannt';

        html += `
            <div class="notebook-item" data-id="${notebook.id}">
                <label class="notebook-checkbox">
                    <input type="checkbox" value="${notebook.id}" class="notebook-check">
                    <div class="notebook-info">
                        <div class="notebook-name">📓 ${notebook.name}${isViewed}</div>
                        <div class="notebook-meta">Zuletzt geändert: ${lastModified}</div>
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
        ? `${selectedNotebooks.size} Notizbuch(er) exportieren`
        : 'Ausgewählte exportieren';
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
                statusElement.textContent = "Pfad ausgewählt: " + selectedDir;
                statusElement.className = "status info";
            }
        } catch (err) {
            console.error("Fehler beim Öffnen des Ordnerauswahldialogs:", err);
            statusElement.textContent = "Fehler beim Öffnen des Dateiauswahldialogs";
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
            statusElement.textContent = "⚠ Ein Export läuft bereits! Bitte warten Sie bis dieser abgeschlossen ist.";
            statusElement.className = "status warning";
            return;
        }

        const destPath = destPathElement.value.trim();

        if (!destPath) {
            statusElement.textContent = "Bitte geben Sie einen Zielordner an";
            statusElement.className = "status error";
            return;
        }

        if (selectedNotebooks.size === 0) {
            statusElement.textContent = "Bitte wählen Sie mindestens ein Notizbuch aus";
            statusElement.className = "status error";
            return;
        }

        // Set the global lock
        exportInProgress = true;
        cancelButton.style.display = "inline-block"; // Show cancel button

        await exportNotebooks(Array.from(selectedNotebooks), destPath);
    });

    // Export all notebooks
    exportAllButton.addEventListener('click', async () => {
        // CRITICAL: Prevent multiple simultaneous exports
        if (exportInProgress) {
            statusElement.textContent = "⚠ Ein Export läuft bereits! Bitte warten Sie bis dieser abgeschlossen ist.";
            statusElement.className = "status warning";
            return;
        }

        const destPath = destPathElement.value.trim();

        if (!destPath) {
            statusElement.textContent = "Bitte geben Sie einen Zielordner an";
            statusElement.className = "status error";
            return;
        }

        try {
            // Set the global lock IMMEDIATELY
            exportInProgress = true;
            cancelButton.style.display = "inline-block"; // Show cancel button

            // Show initial progress bar at 1% - START VISIBLE
            progressContainer.style.display = "block";
            progressBar.style.background = "linear-gradient(90deg, #ffc107, #ff9800)";
            progressBar.style.transition = "width 0.5s ease";
            progressBar.style.width = "1%";
            progressPercent.textContent = "0%";
            progressText.innerHTML = '<span class="spinner"></span>Export wird gestartet...';

            disableButtons(true);

            let lastMessageTime = Date.now();
            let simulatedProgress = 1; // Start at 1% so bar is immediately visible
            let progressInterval = null;

            console.log("[DEBUG] Starting progress simulation at 1%...");

            // Function to update progress bar
            const updateProgressBar = () => {
                const elapsed = Math.floor((Date.now() - lastMessageTime) / 1000);

                // Increase progress slowly until 80%
                if (simulatedProgress < 80) {
                    // Speed: reach 80% in about 10 minutes (600 seconds)
                    // Increase by ~0.13% per second (80% / 600s)
                    simulatedProgress += 0.27; // Every 2 seconds = 0.27% * 30 = ~8% per minute
                    if (simulatedProgress > 80) simulatedProgress = 80;

                    // Show progress with last message or heartbeat
                    let message = "Export läuft... OneNote arbeitet im Hintergrund.";
                    if (elapsed > 5) {
                        message = `Export läuft... (${elapsed}s seit letztem Update)\nOneNote arbeitet im Hintergrund, bitte warten Sie.`;
                    }

                    // DIRECT DOM manipulation - force immediate update
                    // Use Math.max(1, ...) to ensure bar is always at least 1% wide (visible)
                    const displayProgress = Math.max(1, Math.round(simulatedProgress));
                    progressBar.style.width = displayProgress + "%";
                    progressPercent.textContent = Math.round(simulatedProgress) + "%";
                    progressBar.style.background = "linear-gradient(90deg, #ffc107, #ff9800)";
                    progressText.innerHTML = `<span class="spinner"></span>${message}`;

                    console.log(`[DEBUG] Progress updated: ${displayProgress}% (actual: ${simulatedProgress.toFixed(2)}%)`);
                }
            };

            // Call immediately once to show initial state
            updateProgressBar();

            // Then update every 2 seconds
            progressInterval = setInterval(updateProgressBar, 2000);

            // Set up event listener for LIVE progress updates from C# helper (during export)
            EventsOn('export-progress', (data) => {
                if (data && data.message) {
                    lastMessageTime = Date.now();
                    console.log("[Progress] " + data.message);
                    // Update only the message text, keep the simulated progress bar
                    progressText.innerHTML = `<span class="spinner"></span>${data.message}`;
                }
            });

            // Set up event listener for completion
            EventsOn('export-complete', (data) => {
                console.log("[Complete] ", data);

                // Clean up progress simulation
                if (progressInterval) {
                    clearInterval(progressInterval);
                }

                // Clean up event listeners
                EventsOff('export-progress');
                EventsOff('export-complete');
                EventsOff('export-cancelled');

                // Release the global lock and hide cancel button
                exportInProgress = false;
                cancelButton.style.display = "none";

                // Show GREEN completion bar at 100%
                if (data.success) {
                    showCompletion("Export erfolgreich abgeschlossen!");
                    // Hide progress after 2 seconds and show final message
                    setTimeout(() => {
                        hideProgress();
                        statusElement.textContent = data.message;
                        statusElement.className = "status success";
                    }, 2000);
                } else {
                    hideProgress();
                    statusElement.textContent = data.message;
                    statusElement.className = "status error";
                }

                disableButtons(false);
            });

            // Set up event listener for cancellation
            EventsOn('export-cancelled', (data) => {
                console.log("[Cancelled] ", data);

                // Clean up progress simulation
                if (progressInterval) {
                    clearInterval(progressInterval);
                }

                // Clean up event listeners
                EventsOff('export-progress');
                EventsOff('export-complete');
                EventsOff('export-cancelled');

                // Release the global lock and hide cancel button
                exportInProgress = false;
                cancelButton.style.display = "none";

                hideProgress();
                statusElement.textContent = data.message || "Export wurde abgebrochen";
                statusElement.className = "status warning";
                disableButtons(false);
            });

            // Start the export (returns immediately, runs in background)
            const startResult = await ExportAllNotebooks(destPath);
            console.log("[Start] Export started: ", startResult.message);
            lastMessageTime = Date.now();

            // Export is now running in background, events will arrive in real-time

        } catch (err) {
            // Clean up event listeners on error
            EventsOff('export-progress');
            EventsOff('export-complete');
            EventsOff('export-cancelled');

            // Release the global lock and hide cancel button
            exportInProgress = false;
            cancelButton.style.display = "none";

            hideProgress();
            console.error(err);
            statusElement.textContent = "Fehler beim Exportieren: " + err.message;
            statusElement.className = "status error";
            disableButtons(false);
        }
    });
}

// Setup cancel button
function setupCancelButton() {
    cancelButton.addEventListener('click', async () => {
        if (!exportInProgress) {
            return; // Should not happen, button should be hidden
        }

        // Ask for confirmation
        if (!confirm("Möchten Sie den Export wirklich abbrechen?\n\nWarnung: Dies beendet OneNote und OneNoteHelper!")) {
            return;
        }

        console.log("[Cancel] User requested cancellation");
        statusElement.textContent = "Export wird abgebrochen...";
        statusElement.className = "status warning";

        try {
            const result = await CancelExport();
            console.log("[Cancel] ", result.message);
        } catch (err) {
            console.error("[Cancel] Error: ", err);
            statusElement.textContent = "Fehler beim Abbrechen: " + err.message;
            statusElement.className = "status error";
        }
    });
}

// Export selected notebooks one by one
async function exportNotebooks(notebookIds, destPath) {
    try {
        disableButtons(true);

        let successCount = 0;
        let failCount = 0;
        const messages = [];
        let dummyProgressInterval = null;

        for (let i = 0; i < notebookIds.length; i++) {
            const notebookId = notebookIds[i];
            const notebook = notebooks.find(nb => nb.id === notebookId);
            const notebookName = notebook ? notebook.name : 'Unbekannt';

            // Calculate base progress (where we should be between notebooks)
            const baseProgress = (i / notebookIds.length) * 100;
            const nextProgress = ((i + 1) / notebookIds.length) * 100;
            const progressRange = nextProgress - baseProgress;

            // Start with base progress for this notebook
            let currentDummyProgress = baseProgress;

            // Show initial progress
            progressContainer.style.display = "block";
            progressBar.style.background = "linear-gradient(90deg, #ffc107, #ff9800)";
            progressBar.style.transition = "width 0.5s ease";
            progressBar.style.width = Math.max(1, Math.round(baseProgress)) + "%";
            progressPercent.textContent = Math.round(baseProgress) + "%";
            progressText.innerHTML = `<span class="spinner"></span>Exportiere Notizbuch ${i + 1}/${notebookIds.length}: ${notebookName}\nOneNote schreibt im Hintergrund, dies kann mehrere Minuten dauern...`;

            console.log(`[DEBUG] Starting export of notebook ${i + 1}/${notebookIds.length} at ${Math.round(baseProgress)}%`);

            // Start dummy progress animation within this notebook's range
            dummyProgressInterval = setInterval(() => {
                // Slowly increase within the range allocated to this notebook (but max 90% of range)
                if (currentDummyProgress < baseProgress + (progressRange * 0.9)) {
                    currentDummyProgress += progressRange * 0.02; // Increase by 2% of range every interval

                    const displayProgress = Math.max(1, Math.round(currentDummyProgress));
                    progressBar.style.width = displayProgress + "%";
                    progressPercent.textContent = Math.round(currentDummyProgress) + "%";
                    progressBar.style.background = "linear-gradient(90deg, #ffc107, #ff9800)";

                    console.log(`[DEBUG] Dummy progress: ${displayProgress}% (actual: ${currentDummyProgress.toFixed(2)}%)`);
                }
            }, 1000); // Update every second

            try {
                const result = await ExportNotebook(notebookId, destPath);

                // Stop dummy progress
                if (dummyProgressInterval) {
                    clearInterval(dummyProgressInterval);
                    dummyProgressInterval = null;
                }

                if (result.success) {
                    successCount++;
                    messages.push(`✓ ${notebookName}`);
                } else {
                    failCount++;
                    messages.push(`✗ ${notebookName}: ${result.message}`);
                }
            } catch (err) {
                // Stop dummy progress on error
                if (dummyProgressInterval) {
                    clearInterval(dummyProgressInterval);
                    dummyProgressInterval = null;
                }

                failCount++;
                messages.push(`✗ ${notebookName}: ${err.message}`);
            }

            // Show completed progress for this notebook (jump to next milestone)
            const completedProgress = ((i + 1) / notebookIds.length) * 100;
            progressBar.style.width = Math.round(completedProgress) + "%";
            progressPercent.textContent = Math.round(completedProgress) + "%";
            progressBar.style.background = "linear-gradient(90deg, #007bff, #0056b3)"; // Blue for completed segment
            progressText.innerHTML = `Abgeschlossen: ${i + 1}/${notebookIds.length} Notizbücher`;

            console.log(`[DEBUG] Completed notebook ${i + 1}/${notebookIds.length} at ${Math.round(completedProgress)}%`);

            // Brief pause to show the progress update
            await new Promise(resolve => setTimeout(resolve, 500));
        }

        // Clean up interval if still running
        if (dummyProgressInterval) {
            clearInterval(dummyProgressInterval);
        }

        hideProgress();

        // Show final result
        const finalMessage = `Export abgeschlossen: ${successCount} erfolgreich, ${failCount} fehlgeschlagen\n\n${messages.join('\n')}`;
        statusElement.textContent = finalMessage;
        statusElement.className = failCount === 0 ? "status success" : "status warning";

        // Release the global lock and hide cancel button
        exportInProgress = false;
        cancelButton.style.display = "none";

        disableButtons(false);

    } catch (err) {
        hideProgress();
        console.error(err);
        statusElement.textContent = "Fehler beim Exportieren: " + err.message;
        statusElement.className = "status error";

        // Release the global lock and hide cancel button
        exportInProgress = false;
        cancelButton.style.display = "none";

        disableButtons(false);
    }
}

// Show progress indicator with percentage
// Color based on percentage: orange (0-79%), blue (80-99%), green (100%)
function showProgress(message, percent) {
    progressContainer.style.display = "block";

    // CRITICAL: Force re-render by removing and re-adding classes
    progressBar.className = 'progress-bar';

    const roundedPercent = Math.round(percent);

    // Color coding based on progress
    if (percent >= 100) {
        progressBar.classList.add('completed'); // Green
    } else if (percent >= 80) {
        progressBar.classList.add('normal'); // Blue (almost done)
    } else {
        // Orange/yellow for in-progress (0-79%)
        progressBar.style.background = "linear-gradient(90deg, #ffc107, #ff9800)";
        progressBar.style.transition = "width 0.5s ease"; // Smooth animation
    }

    // Force immediate style recalculation
    void progressBar.offsetWidth;

    progressBar.style.width = Math.min(percent, 100) + "%";
    progressPercent.textContent = roundedPercent + "%";
    progressText.innerHTML = `<span class="spinner"></span>${message}`;
    statusElement.textContent = "";

    // Debug log to verify function is being called
    console.log(`[UI] Progress: ${roundedPercent}% - ${message.substring(0, 50)}`);
}

// Show completion (GREEN bar)
function showCompletion(message) {
    progressContainer.style.display = "block";
    progressBar.classList.remove('indeterminate', 'normal');
    progressBar.classList.add('completed');
    progressBar.style.width = "100%";
    progressPercent.textContent = "✓ Fertig";
    progressText.innerHTML = message;
    statusElement.textContent = "";
}

// Hide progress indicator
function hideProgress() {
    progressContainer.style.display = "none";
    progressBar.classList.remove('indeterminate');
    progressBar.style.width = "0";
    progressPercent.textContent = "";
    progressText.textContent = "";
}

// Disable/enable buttons during export
function disableButtons(disabled) {
    exportSelectedButton.disabled = disabled;
    exportAllButton.disabled = disabled;
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
