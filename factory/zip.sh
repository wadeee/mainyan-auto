#!/bin/bash
# 打包 factory 目录（git 已跟踪 + 未被忽略的新文件）到 ../mainyan-auto.zip
# 业务功能的 bat 不参与打包，例如：
#   xxx+1日.bat / xxx-2日.bat / xxx+0周.bat / xxx-1月.bat
#   xxx_2026-06-18.bat / xxx_2026-6.bat
# 环境/安装类 bat（如 auto-install.bat）仍然打包。

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
OUT_ZIP="$SCRIPT_DIR/../mainyan-auto.zip"

is_business_bat() {
    case "$1" in
        # 带天数/周数/月数偏移，如 兔司家门店订购商品统计表+1日.bat
        *[+-][0-9]*日.bat | *[+-][0-9]*周.bat | *[+-][0-9]*月.bat) return 0 ;;
        # 带日期后缀，如 兔司家门店订购商品统计表_2026-06-18.bat
        *_[0-9][0-9][0-9][0-9]-[0-9]*.bat) return 0 ;;
    esac
    return 1
}

cd "$SCRIPT_DIR" || exit 1

files=()
skipped=0
while IFS= read -r -d '' path; do
    if is_business_bat "$(basename "$path")"; then
        skipped=$((skipped + 1))
        continue
    fi
    files+=("$path")
done < <(git ls-files -z --cached --others --exclude-standard)

if [ "${#files[@]}" -eq 0 ]; then
    echo "没有可打包的文件（git ls-files 为空？）" >&2
    exit 1
fi

rm -f "$OUT_ZIP"
printf '%s\0' "${files[@]}" | xargs -0 zip "$OUT_ZIP"

echo "已打包 ${#files[@]} 个文件，排除 ${skipped} 个业务功能 bat → $OUT_ZIP" >&2
