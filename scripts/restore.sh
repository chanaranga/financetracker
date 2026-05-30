#!/bin/bash
# Usage: ./restore.sh [finance-YYYY-MM-DD.db]
# If no argument given, lists available backups and prompts.

ICLOUD_PATH="iclouddrive:Backups/financetracker"
DB_DEST="$HOME/financetracker/backend/data/finance.db"
BACKUP_DIR="$HOME/financetracker/backups"

echo "Available backups in iCloud:"
rclone ls "$ICLOUD_PATH" 2>/dev/null | sort -r || echo "(none found — check rclone config)"
echo ""
echo "Local backups:"
ls -lt "$BACKUP_DIR"/finance-*.db 2>/dev/null | awk '{print $NF}' | xargs -I{} basename {} || echo "(none)"
echo ""

FILE="${1:-}"
if [ -z "$FILE" ]; then
  read -p "Enter backup filename (e.g. finance-2026-05-30.db): " FILE
fi

# Try local first, then iCloud
LOCAL="$BACKUP_DIR/$FILE"
if [ ! -f "$LOCAL" ]; then
  echo "Not found locally, downloading from iCloud..."
  rclone copy "$ICLOUD_PATH/$FILE" "$BACKUP_DIR/"
fi

if [ ! -f "$LOCAL" ]; then
  echo "ERROR: $FILE not found in local backups or iCloud."
  exit 1
fi

echo "Stopping financetracker service..."
sudo systemctl stop financetracker

echo "Restoring $FILE ..."
cp "$DB_DEST" "$DB_DEST.pre-restore-$(date +%Y%m%d%H%M%S)"
cp "$LOCAL" "$DB_DEST"

echo "Restarting financetracker service..."
sudo systemctl start financetracker

echo "Restore complete."
