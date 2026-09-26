#!/bin/bash
# Compatibility entrypoint; implementation lives in mirror_skill.py.
set -euo pipefail
exec python3 "$(cd "$(dirname "$0")" && pwd)/mirror_skill.py" "$@"
