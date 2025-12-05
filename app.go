package main

import (
	"context"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"syscall"

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
		return ExportResult{Success: false, Message: "Could not determine OneNote backup path"}
	}

	// Create destination directory if it doesn't exist
	err := os.MkdirAll(destPath, 0755)
	if err != nil {
		return ExportResult{Success: false, Message: fmt.Sprintf("Error creating destination directory: %v", err)}
	}

	// Copy the files
	err = a.copyDirectory(oneNotePath, destPath)
	if err != nil {
		return ExportResult{Success: false, Message: fmt.Sprintf("Error copying files: %v", err)}
	}

	// Open the folder in explorer
	a.openFolder(destPath)

	return ExportResult{Success: true, Message: "Export completed successfully!"}
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
	// Use Wails runtime function for folder selection dialog
	selectedDir, err := wruntime.OpenDirectoryDialog(a.ctx, wruntime.OpenDialogOptions{
		Title:            "Select Destination Folder",
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
		return nil, fmt.Errorf("OneNote Helper is not available. Please ensure that OneNote Desktop is installed and the Helper program is compiled.")
	}

	return a.helper.GetVersion()
}

// GetNotebooks returns all OneNote notebooks via COM API
func (a *App) GetNotebooks() ([]NotebookInfo, error) {
	fmt.Println("DEBUG: GetNotebooks called")

	if a.helper == nil {
		errMsg := "OneNote Helper is not available. Please ensure that OneNote Desktop is installed and the Helper program is compiled."
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
			Message: "OneNote Helper is not available",
		}, fmt.Errorf("OneNote Helper is not available")
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
// Runs asynchronously in background, sending real-time progress events to frontend
func (a *App) ExportAllNotebooks(destinationPath string) (*ExportResult, error) {
	if a.helper == nil {
		return &ExportResult{
			Success: false,
			Message: "OneNote Helper is not available",
		}, fmt.Errorf("OneNote Helper is not available")
	}

	// Start export in a separate goroutine to not block the frontend
	// This allows events to be sent in real-time while export is running
	go func() {
		fmt.Println("DEBUG: Starting async export...")

		// Progress callback that parses stderr output from C# helper and sends events to frontend
		progressCallback := func(line string) {
			// Emit the raw line as a progress update
			fmt.Fprintf(os.Stderr, "[Progress] %s\n", line)
			wruntime.EventsEmit(a.ctx, "export-progress", map[string]interface{}{
				"message": line,
				"type":    "status",
			})
		}

		// Call C# helper with progress streaming
		result, err := a.helper.ExportAllNotebooks(destinationPath, progressCallback)

		fmt.Println("DEBUG: Export finished, sending completion event...")

		// Send completion event to frontend
		if err != nil {
			wruntime.EventsEmit(a.ctx, "export-complete", map[string]interface{}{
				"success": false,
				"message": err.Error(),
			})
		} else {
			// Open the folder in explorer if successful
			if result.Success {
				a.openFolder(destinationPath)
			}

			wruntime.EventsEmit(a.ctx, "export-complete", map[string]interface{}{
				"success":      result.Success,
				"message":      result.Message,
				"exportedPath": result.ExportedPath,
			})
		}
	}()

	// Return immediately so frontend doesn't block waiting for response
	return &ExportResult{
		Success: true,
		Message: "Export started...",
	}, nil
}

// CancelExport cancels a running export by killing both OneNoteHelper.exe and ONENOTE.EXE processes
func (a *App) CancelExport() (*ExportResult, error) {
	fmt.Println("DEBUG: CancelExport called - killing processes...")

	killedProcesses := []string{}
	var lastError error

	// Kill OneNoteHelper.exe
	if err := killProcessByName("OneNoteHelper.exe"); err != nil {
		fmt.Printf("Warning: Failed to kill OneNoteHelper.exe: %v\n", err)
		lastError = err
	} else {
		killedProcesses = append(killedProcesses, "OneNoteHelper.exe")
		fmt.Println("✓ Killed OneNoteHelper.exe")
	}

	// Kill ONENOTE.EXE
	if err := killProcessByName("ONENOTE.EXE"); err != nil {
		fmt.Printf("Warning: Failed to kill ONENOTE.EXE: %v\n", err)
		lastError = err
	} else {
		killedProcesses = append(killedProcesses, "ONENOTE.EXE")
		fmt.Println("✓ Killed ONENOTE.EXE")
	}

	// Emit event to frontend to notify cancellation
	wruntime.EventsEmit(a.ctx, "export-cancelled", map[string]interface{}{
		"message": "Export was cancelled",
	})

	if len(killedProcesses) == 0 && lastError != nil {
		return &ExportResult{
			Success: false,
			Message: fmt.Sprintf("Error terminating processes: %v", lastError),
		}, lastError
	}

	message := fmt.Sprintf("Export cancelled. Terminated processes: %v", killedProcesses)
	if lastError != nil {
		message += fmt.Sprintf("\nNote: Some processes could not be terminated: %v", lastError)
	}

	return &ExportResult{
		Success: true,
		Message: message,
	}, nil
}

// killProcessByName kills all processes with the given name (Windows-specific)
func killProcessByName(processName string) error {
	// Use taskkill command on Windows
	cmd := exec.Command("taskkill", "/F", "/IM", processName)
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow:    true,
		CreationFlags: 0x08000000, // CREATE_NO_WINDOW
	}

	output, err := cmd.CombinedOutput()
	if err != nil {
		// Check if error is "process not found" (exit code 128)
		if exitErr, ok := err.(*exec.ExitError); ok {
			if exitErr.ExitCode() == 128 {
				// Process not found - not an error, it's just not running
				fmt.Printf("Process %s not found (not running)\n", processName)
				return nil
			}
		}
		return fmt.Errorf("taskkill failed: %w, output: %s", err, string(output))
	}

	return nil
}
