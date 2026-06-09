#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
rm -f "$SCRIPT_DIR/../mainyan-auto.zip"
cd "$SCRIPT_DIR" && git ls-files -z --cached --others --exclude-standard | xargs -0 zip "$SCRIPT_DIR/../mainyan-auto.zip"
