#!/usr/bin/env bash
# git-push.sh – stage, commit und push auf main
set -euo pipefail

# Commit-Message als Argument oder interaktiv abfragen
if [[ $# -gt 0 ]]; then
  MSG="$*"
else
  read -rp "Commit-Message: " MSG
fi

if [[ -z "$MSG" ]]; then
  echo "FEHLER: Keine Commit-Message angegeben."
  exit 1
fi

BRANCH=$(git rev-parse --abbrev-ref HEAD)

git add .
git commit -m "$MSG"
git push origin "$BRANCH"

echo ""
echo "Gepusht: $(git rev-parse --short HEAD) auf $BRANCH – $MSG"
