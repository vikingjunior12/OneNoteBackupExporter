# OneNote Helper

C# console program for communicating with OneNote Desktop via the COM API.

## Prerequisites

- .NET 6.0 SDK or higher
- OneNote Desktop 2016 (OneNote for Windows)
- Windows (x86 or x64)

## Build

```bash
dotnet build -c Release
```

The compiled program will be located at: `bin/Release/net6.0-windows/OneNoteHelper.exe`

## Usage

The program communicates via JSON-RPC over stdin/stdout with the Go backend.

### Supported Methods

#### GetVersion
Returns version information and OneNote status.

Request:
```json
{"method": "GetVersion", "id": 1}
```

#### GetNotebooks
Lists all OneNote notebooks.

Request:
```json
{"method": "GetNotebooks", "id": 2}
```

#### ExportNotebook
Exports a single notebook as a .onepkg file.

Request:
```json
{
  "method": "ExportNotebook",
  "params": {
    "notebookId": "{NOTEBOOK-ID}",
    "destinationPath": "C:\\Export"
  },
  "id": 3
}
```

#### ExportAllNotebooks
Exports all notebooks to a folder.

Request:
```json
{
  "method": "ExportAllNotebooks",
  "params": {
    "destinationPath": "C:\\Export"
  },
  "id": 4
}
```

## Error Handling

The program returns JSON-RPC compliant error responses:

- `-32700`: Parse error (invalid JSON)
- `-32600`: Invalid Request
- `-32601`: Method not found
- `-32602`: Invalid params
- `-32603`: Internal error
- `-32000`: OneNote-specific error

## Notes

- The program is compiled as x86 to be compatible with most OneNote installations
- OneNote must be installed, but does not necessarily need to be running
- Locked or password-protected sections cannot be exported
- Large notebooks may take several minutes to export
