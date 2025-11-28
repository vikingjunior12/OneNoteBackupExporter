import './style.css';
import './app.css';

import { GetNotebooks, ExportNotebook, ExportAllNotebooks, GetOneNoteVersion, BrowseFolder, GetDefaultDownloadsPath } from '../wailsjs/go/main/App';

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
                </div>
                <div class="progress-container" id="progress-container" style="display: none;">
                    <div class="progress-bar-container">
                        <div class="progress-bar" id="progress-bar"></div>
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
const statusElement = document.getElementById("status");
const progressContainer = document.getElementById("progress-container");
const progressBar = document.getElementById("progress-bar");
const progressText = document.getElementById("progress-text");

let notebooks = [];
let selectedNotebooks = new Set();

// Track if listeners are already set up (prevents duplicates on HMR reload)
let listenersInitialized = false;

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

        await exportNotebooks(Array.from(selectedNotebooks), destPath);
    });

    // Export all notebooks
    exportAllButton.addEventListener('click', async () => {
        const destPath = destPathElement.value.trim();

        if (!destPath) {
            statusElement.textContent = "Bitte geben Sie einen Zielordner an";
            statusElement.className = "status error";
            return;
        }

        try {
            // Show indeterminate progress (we don't know how long it will take)
            showIndeterminateProgress("Exportiere alle Notizbücher... OneNote arbeitet im Hintergrund, bitte warten.");
            disableButtons(true);

            const result = await ExportAllNotebooks(destPath);

            hideProgress();

            if (result.success) {
                statusElement.textContent = result.message;
                statusElement.className = "status success";
            } else {
                statusElement.textContent = result.message;
                statusElement.className = "status error";
            }

            disableButtons(false);
        } catch (err) {
            hideProgress();
            console.error(err);
            statusElement.textContent = "Fehler beim Exportieren: " + err.message;
            statusElement.className = "status error";
            disableButtons(false);
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

        for (let i = 0; i < notebookIds.length; i++) {
            const notebookId = notebookIds[i];
            const notebook = notebooks.find(nb => nb.id === notebookId);
            const notebookName = notebook ? notebook.name : 'Unbekannt';

            // Show indeterminate progress while exporting (OneNote works asynchronously)
            showIndeterminateProgress(
                `Exportiere ${i + 1}/${notebookIds.length}: ${notebookName}...\n` +
                `OneNote schreibt im Hintergrund, dies kann mehrere Minuten dauern.`
            );

            try {
                const result = await ExportNotebook(notebookId, destPath);

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

            // Show completed progress after export finishes
            const progress = ((i + 1) / notebookIds.length) * 100;
            showProgress(
                `${i + 1}/${notebookIds.length} Notizbücher abgeschlossen`,
                progress
            );

            // Brief pause to show the progress update
            await new Promise(resolve => setTimeout(resolve, 500));
        }

        hideProgress();

        // Show final result
        const finalMessage = `Export abgeschlossen: ${successCount} erfolgreich, ${failCount} fehlgeschlagen\n\n${messages.join('\n')}`;
        statusElement.textContent = finalMessage;
        statusElement.className = failCount === 0 ? "status success" : "status warning";

        disableButtons(false);

    } catch (err) {
        hideProgress();
        console.error(err);
        statusElement.textContent = "Fehler beim Exportieren: " + err.message;
        statusElement.className = "status error";
        disableButtons(false);
    }
}

// Show progress indicator with percentage
function showProgress(message, percent) {
    progressContainer.style.display = "block";
    progressBar.classList.remove('indeterminate');
    progressBar.style.width = percent + "%";
    progressBar.textContent = Math.round(percent) + "%";
    progressText.innerHTML = `<span class="spinner"></span>${message}`;
    statusElement.textContent = "";
}

// Show indeterminate progress (for unknown duration operations like OneNote writing)
function showIndeterminateProgress(message) {
    progressContainer.style.display = "block";
    progressBar.classList.add('indeterminate');
    progressBar.style.width = "100%";
    progressBar.textContent = "";
    progressText.innerHTML = `<span class="spinner"></span>${message}`;
    statusElement.textContent = "";
}

// Hide progress indicator
function hideProgress() {
    progressContainer.style.display = "none";
    progressBar.classList.remove('indeterminate');
    progressBar.style.width = "0";
    progressBar.textContent = "";
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
