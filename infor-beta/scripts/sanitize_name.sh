#!/usr/bin/env bash
# Sanitize a string for safe use as a filename component.
#
# Rules:
#   1. Replace every run of non-alphanumeric characters with a single hyphen.
#   2. Trim leading and trailing hyphens.
#   3. Collapse multiple consecutive hyphens (handled by rule 1, but explicit).
#
# Preserves case so "Rogers" stays "Rogers", not "rogers".
#
# Usage:
#   SANITIZED=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/sanitize_name.sh" "Rogers Communications Inc.")
#   # SANITIZED = "Rogers-Communications-Inc"
#
#   bash sanitize_name.sh "NasdaqGS:MSFT"   # -> NasdaqGS-MSFT
#   bash sanitize_name.sh "Dye & Durham"    # -> Dye-Durham
#   bash sanitize_name.sh "  spaces  "      # -> spaces

set -u

INPUT="${1:-}"
if [ -z "$INPUT" ]; then
  echo "usage: sanitize_name.sh <string>" >&2
  exit 2
fi

# Step 1: replace any character outside [A-Za-z0-9] with a single hyphen
# Step 2: collapse multiple hyphens into one (defensive)
# Step 3: strip leading + trailing hyphens
printf '%s' "$INPUT" \
  | sed -E 's/[^A-Za-z0-9]+/-/g; s/^-+//; s/-+$//'
echo
