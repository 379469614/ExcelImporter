#!/usr/bin/env bash
# 生成数据配置与文案配置的可携带发布包。
set -euo pipefail
cd "$(dirname "$0")"

RELEASE_DIR="发布包"
VENV_DIR=".venv-pack"
SOURCE_FILES=("导表路径规划.py" "导表入口.py" "导表核心.py" "导表工具集.py" "导表输出器.py" "嵌套解析器.py" "导表工具_图形界面.py" "导出数据配置.py" "导出文案配置.py" "数据配置图形入口.py" "文案配置图形入口.py" "数据配置便携入口.py" "文案配置便携入口.py" "生成独立脚本.py" "发布说明.md")

记录() { printf '[打包] %s\n' "$*"; }
失败() { printf '[错误] %s\n' "$*" >&2; exit 1; }

准备源码包() {
    local PACKAGE_NAME="$1"
    local COMMAND_ENTRY="$2"
    local PACKAGE_DIR="$RELEASE_DIR/$PACKAGE_NAME"
    mkdir -p "$PACKAGE_DIR"
    cp "${SOURCE_FILES[@]}" "$PACKAGE_DIR/"
    cp -R sxl "$PACKAGE_DIR/sxl"
    cp "$COMMAND_ENTRY" "$PACKAGE_DIR/${PACKAGE_NAME}_AI.py"
    cp "$0" "$PACKAGE_DIR/打包工具.sh"
}

打包便携脚本() {
    "$VENV_DIR/bin/python" 生成独立脚本.py
}

打包Linux() {
    local PACKAGE_NAME="$1"
    local GUI_ENTRY="$2"
    local PACKAGE_DIR="$RELEASE_DIR/$PACKAGE_NAME"
    "$VENV_DIR/bin/python" -m PyInstaller --noconfirm --clean -F -w --name "${PACKAGE_NAME}Ubuntu" --paths . "$GUI_ENTRY"
    test -x "dist/${PACKAGE_NAME}Ubuntu" || 失败 "未生成 Linux 二进制"
    cp "dist/${PACKAGE_NAME}Ubuntu" "$PACKAGE_DIR/"
}

打包Windows() {
    local PACKAGE_NAME="$1"
    local GUI_ENTRY="$2"
    local PACKAGE_DIR="$RELEASE_DIR/$PACKAGE_NAME"
    command -v wine >/dev/null || 失败 "未找到 wine"
    local WINDOWS_PYTHON
    WINDOWS_PYTHON="$(wine cmd /c "where python" 2>/dev/null | tr -d '\r' | head -1 | sed 's|\\|/|g')"
    test -n "$WINDOWS_PYTHON" || 失败 "wine 中未找到 Windows Python"
    wine "$WINDOWS_PYTHON" -m PyInstaller --noconfirm --clean -F -w --name "$PACKAGE_NAME" --paths . "$GUI_ENTRY"
    test -s "dist/${PACKAGE_NAME}.exe" || 失败 "未生成 Windows 二进制"
    cp "dist/${PACKAGE_NAME}.exe" "$PACKAGE_DIR/"
}

test -x "$VENV_DIR/bin/python" || 失败 "缺少 $VENV_DIR/bin/python"
"$VENV_DIR/bin/python" -m PyInstaller --version >/dev/null || 失败 "缺少 PyInstaller"

rm -rf "$RELEASE_DIR"
准备源码包 "数据配置导表工具" "导出数据配置.py"
准备源码包 "文案配置导表工具" "导出文案配置.py"
打包便携脚本
记录 "打包 Linux 二进制"
打包Linux "数据配置导表工具" "数据配置图形入口.py"
打包Linux "文案配置导表工具" "文案配置图形入口.py"
记录 "打包 Windows 二进制"
打包Windows "数据配置导表工具" "数据配置图形入口.py"
打包Windows "文案配置导表工具" "文案配置图形入口.py"
记录 "完成：$RELEASE_DIR"
