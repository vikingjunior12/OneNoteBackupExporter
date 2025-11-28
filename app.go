package main

import (
	"context"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"

	wruntime "github.com/wailsapp/wails/v2/pkg/runtime"
)

// App struct
type App struct {
	ctx    context.Context
	helper *OneNoteHelper
}

// FileItem represents a file or directory in the OneNote backup
type FileItem struct {
	Name     string     `json:"name"`
	Path     string     `json:"path"`
	IsDir    bool       `json:"isDir"`
	Children []FileItem `json:"children,omitempty"`
	Level    int        `json:"level"`
}

// ExportResult represents the result of an export operation
type ExportResult struct {
	Success      bool   `json:"success"`
	Message      string `json:"message"`
	ExportedPath string `json:"exportedPath,omitempty"`
}

// NotebookInfo represents a OneNote notebook
type NotebookInfo struct {
	ID                string `json:"id"`
	Name              string `json:"name"`
	Path              string `json:"path"`
	LastModified      string `json:"lastModified"`
	IsCurrentlyViewed bool   `json:"isCurrentlyViewed"`
}

// VersionInfo represents version information about the helper and OneNote
type VersionInfo struct {
	Version          string `json:"version"`
	OneNoteInstalled bool   `json:"oneNoteInstalled"`
	OneNoteVersion   string `json:"oneNoteVersion"`
}

// NewApp creates a new App application struct
func NewApp() *App {
	return &App{}
}

// startup is called when the app starts. The context is saved
// so we can call the runtime methods
func (a *App) startup(ctx context.Context) {
	a.ctx = ctx

	// Initialize OneNote helper (don't fail if not available)
	helper, err := NewOneNoteHelper()
	if err != nil {
		fmt.Printf("Warning: OneNote Helper not available: %v\n", err)
	} else {
		a.helper = helper
	}
}

// GetOneNoteBackupPath returns the path to the OneNote backup folder
func (a *App) GetOneNoteBackupPath() string {
	// Get the user's home directory
	homeDir, err := os.UserHomeDir()
	if err != nil {
		fmt.Printf("Error getting user home directory: %v\n", err)
		return ""
	}

	// Construct the path to the OneNote backup folder
	oneNotePath := filepath.Join(homeDir, "AppData", "Local", "Microsoft", "OneNote", "16.0", "Sicherung")
	return oneNotePath
}

// GetBackupContents returns the contents of the OneNote backup folder
func (a *App) GetBackupContents() []FileItem {
	oneNotePath := a.GetOneNoteBackupPath()
	if oneNotePath == "" {
		return []FileItem{}
	}

	// Get the root contents
	rootItems, err := os.ReadDir(oneNotePath)
	if err != nil {
		fmt.Printf("Error reading directory: %v\n", err)
		return []FileItem{}
	}

	// Convert to FileItem structs
	var items []FileItem
	for _, item := range rootItems {
		isDir := item.IsDir()
		itemPath := filepath.Join(oneNotePath, item.Name())

		fileItem := FileItem{
			Name:  item.Name(),
			Path:  itemPath,
			IsDir: isDir,
			Level: 0,
		}

		// If it's a directory, get its children
		if isDir {
			fileItem.Children = a.getChildItems(itemPath, 1)
		}

		items = append(items, fileItem)
	}

	return items
}

// getChildItems recursively gets child items for a directory
func (a *App) getChildItems(dirPath string, level int) []FileItem {
	items, err := os.ReadDir(dirPath)
	if err != nil {
		return []FileItem{}
	}

	var children []FileItem
	for _, item := range items {
		isDir := item.IsDir()
		itemPath := filepath.Join(dirPath, item.Name())

		fileItem := FileItem{
			Name:  item.Name(),
			Path:  itemPath,
			IsDir: isDir,
			Level: level,
		}

		// If it's a directory, get its children
		if isDir {
			fileItem.Children = a.getChildItems(itemPath, level+1)
		}

		children = append(children, fileItem)
	}

	return children
}

// ExportBackup exports the OneNote backup to the specified destination
func (a *App) ExportBackup(destPath string) ExportResult {
	oneNotePath := a.GetOneNoteBackupPath()
	if oneNotePath == "" {
		return ExportResult{Success: false, Message: "Konnte den OneNote Backup-Pfad nicht ermitteln"}
	}

	// Create destination directory if it doesn't exist
	err := os.MkdirAll(destPath, 0755)
	if err != nil {
		return ExportResult{Success: false, Message: fmt.Sprintf("Fehler beim Erstellen des Zielverzeichnisses: %v", err)}
	}

	// Copy the files
	err = a.copyDirectory(oneNotePath, destPath)
	if err != nil {
		return ExportResult{Success: false, Message: fmt.Sprintf("Fehler beim Kopieren der Dateien: %v", err)}
	}

	// Open the folder in explorer
	a.openFolder(destPath)

	return ExportResult{Success: true, Message: "Export erfolgreich abgeschlossen!"}
}

// copyDirectory recursively copies a directory tree
func (a *App) copyDirectory(src, dst string) error {
	// Get properties of source directory
	srcInfo, err := os.Stat(src)
	if err != nil {
		return err
	}

	// Create the destination directory
	err = os.MkdirAll(dst, srcInfo.Mode())
	if err != nil {
		return err
	}

	items, err := os.ReadDir(src)
	if err != nil {
		return err
	}

	for _, item := range items {
		srcPath := filepath.Join(src, item.Name())
		dstPath := filepath.Join(dst, item.Name())

		if item.IsDir() {
			// Recursively copy subdirectory
			err = a.copyDirectory(srcPath, dstPath)
			if err != nil {
				return err
			}
		} else {
			// Copy file
			err = a.copyFile(srcPath, dstPath)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

// copyFile copies a single file
func (a *App) copyFile(src, dst string) error {
	// Open source file
	srcFile, err := os.Open(src)
	if err != nil {
		return err
	}
	defer srcFile.Close()

	// Create destination file, overwriting if it exists
	dstFile, err := os.Create(dst)
	if err != nil {
		return err
	}
	defer dstFile.Close()

	// Copy the contents
	_, err = io.Copy(dstFile, srcFile)
	if err != nil {
		return err
	}

	// Get source file mode
	srcInfo, err := os.Stat(src)
	if err != nil {
		return err
	}

	// Set the same permissions
	return os.Chmod(dst, srcInfo.Mode())
}

// openFolder opens the specified folder in the file explorer
func (a *App) openFolder(path string) error {
	var cmd *exec.Cmd

	switch runtime.GOOS {
	case "windows":
		cmd = exec.Command("explorer", path)
	case "darwin":
		cmd = exec.Command("open", path)
	case "linux":
		cmd = exec.Command("xdg-open", path)
	default:
		return fmt.Errorf("unsupported platform")
	}

	return cmd.Start()
}

// GetDefaultDownloadsPath returns the path to the user's Downloads folder
func (a *App) GetDefaultDownloadsPath() string {
	homeDir, err := os.UserHomeDir()
	if err != nil {
		return ""
	}
	return filepath.Join(homeDir, "Downloads")
}

// BrowseFolder opens a native Windows folder selection dialog
func (a *App) BrowseFolder() string {
	// Wir verwenden die Wails-Runtime-Funktion für den Ordnerauswahldialog
	selectedDir, err := wruntime.OpenDirectoryDialog(a.ctx, wruntime.OpenDialogOptions{
		Title:            "Zielordner auswählen",
		DefaultDirectory: a.GetDefaultDownloadsPath(),
	})

	if err != nil {
		fmt.Printf("Error opening directory dialog: %v\n", err)
		return ""
	}

	return selectedDir
}

// GetBackupSize returns the total size of the OneNote backup folder in bytes
func (a *App) GetBackupSize() int64 {
	oneNotePath := a.GetOneNoteBackupPath()
	if oneNotePath == "" {
		return 0
	}

	var totalSize int64 = 0

	err := filepath.Walk(oneNotePath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() {
			totalSize += info.Size()
		}
		return nil
	})

	if err != nil {
		fmt.Printf("Error calculating backup size: %v\n", err)
		return 0
	}

	return totalSize
}

// FormatSize formats a size in bytes to a human-readable string (KB, MB, GB)
func (a *App) FormatSize(size int64) string {
	const (
		KB = 1024
		MB = 1024 * KB
		GB = 1024 * MB
	)

	switch {
	case size >= GB:
		return fmt.Sprintf("%.2f GB", float64(size)/float64(GB))
	case size >= MB:
		return fmt.Sprintf("%.2f MB", float64(size)/float64(MB))
	case size >= KB:
		return fmt.Sprintf("%.2f KB", float64(size)/float64(KB))
	default:
		return fmt.Sprintf("%d Bytes", size)
	}
}

// ===== NEW: OneNote COM Integration Methods =====

// GetOneNoteVersion returns version information about OneNote
func (a *App) GetOneNoteVersion() (*VersionInfo, error) {
	if a.helper == nil {
		return nil, fmt.Errorf("OneNote Helper ist nicht verfügbar. Bitte stellen Sie sicher, dass OneNote Desktop installiert ist und das Helper-Programm kompiliert wurde.")
	}

	return a.helper.GetVersion()
}

// GetNotebooks returns all OneNote notebooks via COM API
func (a *App) GetNotebooks() ([]NotebookInfo, error) {
	fmt.Println("DEBUG: GetNotebooks called")

	if a.helper == nil {
		errMsg := "OneNote Helper ist nicht verfügbar. Bitte stellen Sie sicher, dass OneNote Desktop installiert ist und das Helper-Programm kompiliert wurde."
		fmt.Printf("ERROR: %s\n", errMsg)
		return nil, fmt.Errorf(errMsg)
	}

	fmt.Println("DEBUG: Calling helper.GetNotebooks()")
	notebooks, err := a.helper.GetNotebooks()
	if err != nil {
		fmt.Printf("ERROR: GetNotebooks failed: %v\n", err)
		return nil, err
	}

	fmt.Printf("DEBUG: Got %d notebooks\n", len(notebooks))
	return notebooks, nil
}

// ExportNotebook exports a single notebook to .onepkg format
func (a *App) ExportNotebook(notebookID, destinationPath string) (*ExportResult, error) {
	if a.helper == nil {
		return &ExportResult{
			Success: false,
			Message: "OneNote Helper ist nicht verfügbar",
		}, fmt.Errorf("OneNote Helper ist nicht verfügbar")
	}

	result, err := a.helper.ExportNotebook(notebookID, destinationPath)
	if err != nil {
		return &ExportResult{
			Success: false,
			Message: err.Error(),
		}, err
	}

	// Open the folder in explorer if successful
	if result.Success {
		a.openFolder(destinationPath)
	}

	return result, nil
}

// ExportAllNotebooks exports all notebooks to the specified destination
func (a *App) ExportAllNotebooks(destinationPath string) (*ExportResult, error) {
	if a.helper == nil {
		return &ExportResult{
			Success: false,
			Message: "OneNote Helper ist nicht verfügbar",
		}, fmt.Errorf("OneNote Helper ist nicht verfügbar")
	}

	result, err := a.helper.ExportAllNotebooks(destinationPath)
	if err != nil {
		return &ExportResult{
			Success: false,
			Message: err.Error(),
		}, err
	}

	// Open the folder in explorer if successful
	if result.Success {
		a.openFolder(destinationPath)
	}

	return result, nil
}
