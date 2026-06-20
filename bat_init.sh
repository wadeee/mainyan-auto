#!/bin/bash
# 将项目内所有 .bat 文件的换行符转换为 CRLF

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

count=0
while IFS= read -r -d '' file; do
    sed -i -e 's/\r$//' -e 's/$/\r/' "$file"
    echo "  CRLF: $file"
    ((count++))
done < <(find "$SCRIPT_DIR" -name "*.bat" -type f -print0)

echo "Done. $count .bat files converted to CRLF."
