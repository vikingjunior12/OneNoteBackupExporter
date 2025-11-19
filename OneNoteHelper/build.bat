@echo off
echo Building OneNoteHelper...
dotnet build -c Release

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
    echo Output: bin\Release\net6.0-windows\OneNoteHelper.exe
) else (
    echo.
    echo Build failed!
    exit /b %ERRORLEVEL%
)
