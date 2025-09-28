#!/usr/bin/env bash
# Wrapper to call the Python sync script from tools/sync
DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
python3 "$DIR/sync_vba_to_excel.py" "$@"
