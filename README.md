# OneNote Backup Exporter

Eine Desktop-Anwendung zum Exportieren von Microsoft OneNote-Notizbüchern als `.onepkg`-Dateien über die OneNote COM-API.

## Features

- ✅ Direkter Export aus OneNote Desktop via COM-API
- ✅ Export einzelner Notizbücher oder aller auf einmal
- ✅ Vollständige Synchronisierung vor dem Export
- ✅ Progress-Tracking und detaillierte Statusmeldungen
- ✅ Moderne, benutzerfreundliche Oberfläche
- ✅ Funktioniert mit mylms, Klassenotibücher (SharePoint)

## Voraussetzungen

### Zur Laufzeit
- Windows 10/11
- OneNote Desktop 2016 ("OneNote für Windows")
- .NET 6.0 Runtime (normalerweise vorinstalliert)

**Wichtig:** Die App funktioniert nur mit OneNote Desktop (2016 oder 365), nicht mit der UWP-App (Windows 10 OneNote) oder OneNote für Web/Mac.

## Installation 
setup.exe in relase runterladen.
setup.exe ausführen


## Verwendung

1. Anwendung starten
2. Die App erkennt automatisch OneNote Desktop und listet alle Notizbücher auf
3. Notizbücher auswählen (einzeln oder alle)
4. Zielordner wählen oder anpassen
5. "Ausgewählte exportieren" oder "Alle exportieren" klicken
6. Der Export-Ordner öffnet sich automatisch nach erfolgreichem Export

## Architektur

Die Anwendung nutzt eine hybride Architektur:

- **Frontend:** JavaScript/Vite (UI)
- **Backend:** Go/Wails (Anwendungslogik)
- **COM Helper:** C# (OneNote-Integration)

Der C# Helper kommuniziert über JSON-RPC (stdin/stdout) mit dem Go-Backend, welches wiederum über Wails-Bindings mit dem JavaScript-Frontend verbunden ist.

## Bekannte Einschränkungen

- Kennwortgeschützte Abschnitte werden übersprungen
- Sehr große Notizbücher (>500MB) können Timeouts verursachen
- Erfordert OneNote Desktop (keine UWP-Version)
- Nur Windows-kompatibel

## Fehlerbehebung


### Export schlägt fehl
- Prüfen Sie Schreibberechtigungen im Zielordner
- Stellen Sie sicher, dass genügend Speicherplatz verfügbar ist
- OneNotes vorher in der OneNote app vollständig zu Sychronisieren (SharePoint, OneDrive)

## Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert - siehe [LICENSE](LICENSE) für Details.

Copyright © 2025 JLI Software
