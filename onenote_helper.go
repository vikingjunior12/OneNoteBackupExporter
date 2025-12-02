package main

import (
	"bufio"
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"sync"
	"syscall"
)

// JSON-RPC structures
type jsonRpcRequest struct {
	Method string                 `json:"method"`
	Params map[string]interface{} `json:"params,omitempty"`
	ID     int                    `json:"id"`
}

type jsonRpcResponse struct {
	Result json.RawMessage `json:"result,omitempty"`
	Error  *jsonRpcError   `json:"error,omitempty"`
	ID     int             `json:"id"`
}

type jsonRpcError struct {
	Code    int    `json:"code"`
	Message string `json:"message"`
}

// OneNoteHelper manages communication with the C# helper program
type OneNoteHelper struct {
	helperPath string
	mu         sync.Mutex
	requestID  int
}

// NewOneNoteHelper creates a new helper instance
func NewOneNoteHelper() (*OneNoteHelper, error) {
	// Find the helper executable
	// First check in OneNoteHelper/bin/Release/net6.0-windows/
	exePath, err := os.Executable()
	if err != nil {
		return nil, fmt.Errorf("konnte Programmpfad nicht ermitteln: %w", err)
	}

	exeDir := filepath.Dir(exePath)

	// Try different possible locations
	possiblePaths := []string{
		filepath.Join(exeDir, "OneNoteHelper", "bin", "Release", "net8.0-windows", "OneNoteHelper.exe"),
		filepath.Join(exeDir, "OneNoteHelper", "OneNoteHelper.exe"),
		filepath.Join(exeDir, "..", "OneNoteHelper", "bin", "Release", "net8.0-windows", "OneNoteHelper.exe"),
		"OneNoteHelper/bin/Release/net8.0-windows/OneNoteHelper.exe", // Development path
	}

	var helperPath string
	fmt.Printf("Searching for OneNoteHelper.exe...\n")
	fmt.Printf("Exe directory: %s\n", exeDir)
	for _, path := range possiblePaths {
		fmt.Printf("  Trying: %s\n", path)
		if _, err := os.Stat(path); err == nil {
			helperPath = path
			fmt.Printf("  ✓ Found at: %s\n", path)
			break
		}
	}

	if helperPath == "" {
		fmt.Printf("ERROR: OneNoteHelper.exe not found in any location\n")
		return nil, fmt.Errorf("OneNoteHelper.exe nicht gefunden. Bitte zuerst das C# Helper-Programm kompilieren (cd OneNoteHelper && dotnet build -c Release)")
	}

	return &OneNoteHelper{
		helperPath: helperPath,
		requestID:  1,
	}, nil
}

// call executes a JSON-RPC call to the helper program
func (h *OneNoteHelper) call(method string, params map[string]interface{}) (json.RawMessage, error) {
	h.mu.Lock()
	reqID := h.requestID
	h.requestID++
	h.mu.Unlock()

	// Create request
	request := jsonRpcRequest{
		Method: method,
		Params: params,
		ID:     reqID,
	}

	requestJSON, err := json.Marshal(request)
	if err != nil {
		return nil, fmt.Errorf("fehler beim Erstellen der Anfrage: %w", err)
	}

	// Execute helper program
	cmd := exec.Command(h.helperPath)
	cmd.Stdin = bytes.NewReader(requestJSON)

	// Hide the console window (Windows only)
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow:    true,
		CreationFlags: 0x08000000, // CREATE_NO_WINDOW
	}

	var stdout bytes.Buffer
	cmd.Stdout = &stdout
	// Forward stderr directly to console for debugging (C# helper writes diagnostics there)
	cmd.Stderr = os.Stderr

	err = cmd.Run()
	if err != nil {
		return nil, fmt.Errorf("fehler beim Ausführen des Helpers: %w", err)
	}

	// Parse response
	var response jsonRpcResponse
	if err := json.Unmarshal(stdout.Bytes(), &response); err != nil {
		return nil, fmt.Errorf("fehler beim Parsen der Antwort: %w\nOutput: %s", err, stdout.String())
	}

	// Check for RPC error
	if response.Error != nil {
		return nil, fmt.Errorf("RPC-Fehler %d: %s", response.Error.Code, response.Error.Message)
	}

	return response.Result, nil
}

// GetVersion returns version info from the helper
func (h *OneNoteHelper) GetVersion() (*VersionInfo, error) {
	result, err := h.call("GetVersion", nil)
	if err != nil {
		return nil, err
	}

	var versionInfo VersionInfo
	if err := json.Unmarshal(result, &versionInfo); err != nil {
		return nil, fmt.Errorf("fehler beim Parsen der Versionsinformationen: %w", err)
	}

	return &versionInfo, nil
}

// GetNotebooks returns all OneNote notebooks
func (h *OneNoteHelper) GetNotebooks() ([]NotebookInfo, error) {
	result, err := h.call("GetNotebooks", nil)
	if err != nil {
		return nil, err
	}

	var notebooks []NotebookInfo
	if err := json.Unmarshal(result, &notebooks); err != nil {
		return nil, fmt.Errorf("fehler beim Parsen der Notizbücher: %w", err)
	}

	return notebooks, nil
}

// ExportNotebook exports a single notebook to .onepkg format
func (h *OneNoteHelper) ExportNotebook(notebookID, destinationPath string) (*ExportResult, error) {
	params := map[string]interface{}{
		"notebookId":      notebookID,
		"destinationPath": destinationPath,
	}

	result, err := h.call("ExportNotebook", params)
	if err != nil {
		return nil, err
	}

	var exportResult ExportResult
	if err := json.Unmarshal(result, &exportResult); err != nil {
		return nil, fmt.Errorf("fehler beim Parsen des Exportergebnisses: %w", err)
	}

	return &exportResult, nil
}

// ExportAllNotebooks exports all notebooks with real-time progress streaming
func (h *OneNoteHelper) ExportAllNotebooks(destinationPath string, progressCallback func(string)) (*ExportResult, error) {
	params := map[string]interface{}{
		"destinationPath": destinationPath,
	}

	h.mu.Lock()
	reqID := h.requestID
	h.requestID++
	h.mu.Unlock()

	// Create request
	request := jsonRpcRequest{
		Method: "ExportAllNotebooks",
		Params: params,
		ID:     reqID,
	}

	requestJSON, err := json.Marshal(request)
	if err != nil {
		return nil, fmt.Errorf("fehler beim Erstellen der Anfrage: %w", err)
	}

	// Execute helper program
	cmd := exec.Command(h.helperPath)
	cmd.Stdin = bytes.NewReader(requestJSON)

	// Hide the console window (Windows only)
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow:    true,
		CreationFlags: 0x08000000, // CREATE_NO_WINDOW
	}

	var stdout bytes.Buffer
	cmd.Stdout = &stdout

	// Capture stderr for real-time progress updates
	stderrPipe, err := cmd.StderrPipe()
	if err != nil {
		return nil, fmt.Errorf("fehler beim Erstellen der stderr-Pipe: %w", err)
	}

	// Start the command
	if err := cmd.Start(); err != nil {
		return nil, fmt.Errorf("fehler beim Starten des Helpers: %w", err)
	}

	// Read stderr in real-time and send to callback
	go func() {
		scanner := bufio.NewScanner(stderrPipe)
		for scanner.Scan() {
			line := scanner.Text()
			if progressCallback != nil {
				progressCallback(line)
			}
			// Also print to console for debugging
			fmt.Fprintf(os.Stderr, "%s\n", line)
		}
	}()

	// Wait for command to complete
	err = cmd.Wait()
	if err != nil {
		return nil, fmt.Errorf("fehler beim Ausführen des Helpers: %w", err)
	}

	// Parse response
	var response jsonRpcResponse
	if err := json.Unmarshal(stdout.Bytes(), &response); err != nil {
		return nil, fmt.Errorf("fehler beim Parsen der Antwort: %w\nOutput: %s", err, stdout.String())
	}

	// Check for RPC error
	if response.Error != nil {
		return nil, fmt.Errorf("RPC-Fehler %d: %s", response.Error.Code, response.Error.Message)
	}

	var exportResult ExportResult
	if err := json.Unmarshal(response.Result, &exportResult); err != nil {
		return nil, fmt.Errorf("fehler beim Parsen des Exportergebnisses: %w", err)
	}

	return &exportResult, nil
}
