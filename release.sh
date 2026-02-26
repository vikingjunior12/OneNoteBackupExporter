# release.sh – GitHub Release v1.0.0 erstellen und Setup-EXE hochladen
# Voraussetzungen: git, gh (GitHub CLI), authentifiziert via `gh auth login`

set -euo pipefail

VERSION="1.0.0"
TAG="v${VERSION}"
SETUP_EXE="OneNoteBackupExporter_Setup_${VERSION}.exe"
REPO="vikingjunior12/OneNoteBackupExporter"

# ── Prüfungen ────────────────────────────────────────────────────────────────
if ! command -v gh &>/dev/null; then
  echo "FEHLER: 'gh' (GitHub CLI) ist nicht installiert."
  echo "  -> pacman -S github-cli"
  exit 1
fi

if [[ ! -f "$SETUP_EXE" ]]; then
  echo "FEHLER: Setup-Datei nicht gefunden: $SETUP_EXE"
  echo "  -> Erst build.ps1 + Inno Setup ausführen"
  exit 1
fi

# ── Tag erstellen (falls noch nicht vorhanden) ────────────────────────────────
if git rev-parse "$TAG" &>/dev/null; then
  echo ">> Tag $TAG existiert bereits, wird übersprungen."
else
  echo ">> Erstelle Tag $TAG ..."
  git tag -a "$TAG" -m "Release $TAG"
  git push origin "$TAG"
fi

# ── GitHub Release erstellen + EXE hochladen ─────────────────────────────────
echo ">> Erstelle GitHub Release $TAG ..."
gh release create "$TAG" "$SETUP_EXE" \
  --repo "$REPO" \
  --title "OneNoteBackupExporter $TAG" \
  --notes "## OneNoteBackupExporter $TAG

### Vollständige Neuentwicklung als C#/WPF-Anwendung

Diese Version ist eine komplette Neuentwicklung in **reinem C# (WPF)**.
Die frühere Go/Wails-Architektur wurde vollständig ersetzt für bessere Performance, einfachere Installation und direkte COM-Integration ohne Subprocess-Layer.

### Highlights
- Direkter Zugriff auf OneNote via COM API (kein Subprocess mehr)
- Export zu **.onepkg**, **.xps**, **.pdf** und lokalem Backup
- Fortschrittsanzeige mit Abbruch-Funktion
- Self-contained – kein .NET vorinstallieren nötig
- Nur für **Windows x64** (OneNote Desktop erforderlich)

### Installation
1. \`$SETUP_EXE\` herunterladen
2. Setup ausführen
3. OneNote Desktop muss installiert und mindestens einmal geöffnet worden sein

### Systemvoraussetzungen
- Windows 10/11 x64
- Microsoft OneNote Desktop (Microsoft 365 oder 2019/2021)"

echo ""
echo "Release erfolgreich: https://github.com/${REPO}/releases/tag/${TAG}"
