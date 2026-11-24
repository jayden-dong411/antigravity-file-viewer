#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PLUGIN_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

find_node() {
  if command -v node >/dev/null 2>&1; then
    command -v node
    return 0
  fi
  for candidate in \
    "${HOME}/.local/bin/node" \
    "/opt/homebrew/bin/node" \
    "/usr/local/bin/node" \
    "${HOME}/.volta/bin/node"
  do
    if [ -x "$candidate" ]; then
      echo "$candidate"
      return 0
    fi
  done
  echo "Error: Node.js (>=20) not found in PATH." >&2
  exit 1
}

NODE_BIN="$(find_node)"
exec "$NODE_BIN" "${PLUGIN_ROOT}/src/sidecar-server.mjs" "$@"
