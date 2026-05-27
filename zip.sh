#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" && git ls-files --cached --others --exclude-standard | zip "$SCRIPT_DIR/../mainyan-auto.zip" -@
