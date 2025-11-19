# Build script for OneNote Backup Exporter

Write-Host "Building C# Helper..."
cd OneNoteHelper
dotnet build -c Release
cd ..

Write-Host "Building Wails app..."
wails build

Write-Host "Copying OneNote Helper..."
mkdir -Force build\bin\OneNoteHelper | Out-Null
copy OneNoteHelper\bin\Release\net8.0-windows\*.* build\bin\OneNoteHelper\

Write-Host ""
Write-Host "✓ Build complete! App is ready at: build\bin\OneNoteBackupExporter.exe"
