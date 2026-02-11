# Build script for OneNote Backup Exporter

Write-Host "Building C# Helper (self-contained with .NET runtime)..."
cd OneNoteHelper
dotnet publish -c Release -r win-x64 --self-contained true
cd ..

Write-Host "Building Wails app..."
wails build

Write-Host "Copying OneNote Helper..."
mkdir -Force build\bin\OneNoteHelper | Out-Null
copy OneNoteHelper\bin\Release\net8.0-windows\win-x64\publish\*.* build\bin\OneNoteHelper\

Write-Host ""
Write-Host "✓ Build complete! App is ready at: build\bin\OneNoteBackupExporter.exe"
Write-Host ""
Write-Host "Note: The C# Helper includes the .NET 8.0 runtime, so users don't need to install it."
