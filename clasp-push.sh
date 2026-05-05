#!/usr/bin/env bash
set -euo pipefail

DIR="$(dirname "$0")"
TARGETS_FILE="$DIR/clasp-targets.json"
CLASP_JSON="$DIR/.clasp.json"

original_id=$(jq -r '.scriptId' "$CLASP_JSON")

restore() {
  jq --arg id "$original_id" '.scriptId = $id' "$CLASP_JSON" > "$CLASP_JSON.tmp" && mv "$CLASP_JSON.tmp" "$CLASP_JSON"
}
trap restore EXIT

while IFS="=" read -r key id; do
  echo "Pushing to $key ($id)..."
  jq --arg id "$id" '.scriptId = $id' "$CLASP_JSON" > "$CLASP_JSON.tmp" && mv "$CLASP_JSON.tmp" "$CLASP_JSON"
  clasp push --force
  echo "Done: $key"
done < <(jq -r 'to_entries[] | "\(.key)=\(.value)"' "$TARGETS_FILE")

echo "All deployments complete."
