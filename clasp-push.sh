#!/usr/bin/env bash
# Push the Apps Script project to every target spreadsheet listed in clasp-targets.json.
# .clasp.json is temporarily rewritten for each target, then restored on exit.
set -euo pipefail

DIR="$(dirname "$0")"
TARGETS_FILE="$DIR/clasp-targets.json"  # {"name": "scriptId", ...}
CLASP_JSON="$DIR/.clasp.json"

# Save the scriptId that was in .clasp.json before we started.
original_id=$(jq -r '.scriptId' "$CLASP_JSON")

# Always put .clasp.json back so the primary project stays the default.
restore() {
  jq --arg id "$original_id" '.scriptId = $id' "$CLASP_JSON" > "$CLASP_JSON.tmp" && mv "$CLASP_JSON.tmp" "$CLASP_JSON"
}
trap restore EXIT

# Emit "key=scriptId" lines from the JSON object so the while-loop can split on '='.
while IFS="=" read -r key id; do
  echo "Pushing to $key ($id)..."
  jq --arg id "$id" '.scriptId = $id' "$CLASP_JSON" > "$CLASP_JSON.tmp" && mv "$CLASP_JSON.tmp" "$CLASP_JSON"
  clasp push --force
  echo "Done: $key"
done < <(jq -r 'to_entries[] | "\(.key)=\(.value)"' "$TARGETS_FILE")

echo "All deployments complete."
