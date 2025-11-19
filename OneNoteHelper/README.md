# OneNote Helper

C# Konsolen-Programm zur Kommunikation mit OneNote Desktop über die COM-API.

## Voraussetzungen

- .NET 6.0 SDK oder höher
- OneNote Desktop 2016 (OneNote für Windows)
- Windows (x86 oder x64)

## Build

```bash
dotnet build -c Release
```

Das kompilierte Programm liegt dann unter: `bin/Release/net6.0-windows/OneNoteHelper.exe`

## Verwendung

Das Programm kommuniziert über JSON-RPC über stdin/stdout mit dem Go-Backend.

### Unterstützte Methoden

#### GetVersion
Gibt Versionsinformationen und OneNote-Status zurück.

Request:
```json
{"method": "GetVersion", "id": 1}
```

#### GetNotebooks
Listet alle OneNote-Notizbücher auf.

Request:
```json
{"method": "GetNotebooks", "id": 2}
```

#### ExportNotebook
Exportiert ein einzelnes Notizbuch als .onepkg-Datei.

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
Exportiert alle Notizbücher in einen Ordner.

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

## Fehlerbehandlung

Das Programm gibt JSON-RPC-konforme Error-Responses zurück:

- `-32700`: Parse error (ungültiges JSON)
- `-32600`: Invalid Request
- `-32601`: Method not found
- `-32602`: Invalid params
- `-32603`: Internal error
- `-32000`: OneNote-spezifischer Fehler

## Hinweise

- Das Programm ist als x86 kompiliert, um mit den meisten OneNote-Installationen kompatibel zu sein
- OneNote muss installiert sein, aber nicht unbedingt laufen
- Gesperrte oder kennwortgeschützte Abschnitte können nicht exportiert werden
- Große Notizbücher können mehrere Minuten für den Export benötigen
