# release.sh – GitHub Release erstellen und Setup-EXE hochladen
# Voraussetzungen: git, gh (GitHub CLI), authentifiziert via `gh auth login`

set -euo pipefail

VERSION="1.1.0"
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

### New Features

- **Section export**: Notebooks can now be expanded to show individual sections — select and export only the sections you need
- Sections from nested section groups are fully supported (shown as \`📂 Group / 📄 Section\`)
- Encrypted and recycle-bin sections are automatically skipped
- Mixed export: combine whole-notebook exports and section-level exports in a single run
- Sections are saved into a notebook-named subfolder to keep the output organised (e.g. \`MyNotebook/Section1.onepkg\`)
- Export button text dynamically reflects the selection (e.g. \`Export 1 notebook(s) + 3 section(s)\`)

### Installation
1. Download \`$SETUP_EXE\`
2. Run the installer
3. OneNote Desktop must be installed and opened at least once

### System Requirements
- Windows 10/11 x64
- Microsoft OneNote Desktop (Microsoft 365 or 2019/2021)"

echo ""
echo "Release erfolgreich: https://github.com/${REPO}/releases/tag/${TAG}"
