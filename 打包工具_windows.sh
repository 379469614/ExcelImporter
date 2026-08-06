#!/usr/bin/env bash
# 导表工具 Windows 打包脚本（Linux + Wine）
#
# 流程：探测 wine 下的 Windows Python -> 安装 PyQt6 + PyInstaller -> PyInstaller 打包
#       -> 输出 dist/导表工具Windows.exe，供其他开发伙伴在 Windows 上直接双击使用。
# 幂等：重复执行不会重复安装依赖。
#
# 用法：
#   ./打包工具_windows.sh            打包 Windows exe
#   ./打包工具_windows.sh --命令行   打包后用 wine 跑内置 --命令行 自检（验证导出链路），结束时清理临时文件
# 注意：自检需要项目内有 测试配置/ 目录；固定路径(5配置文件/)只在本脚本自检时以复制形式临时存在。
set -euo pipefail
cd "$(dirname "$0")"

APP_NAME="导表工具"
ENTRY_SCRIPT="导表工具_图形界面.py"

# 自检用：把 测试配置 复制到程序写死的固定来源路径，验证后删除
FIXED_SRC_NAME="5配置文件"
FIXED_OUT_DIR="BouncyPinball"

log() { printf '\033[1;36m[打包Win]\033[0m %s\n' "$*"; }
err() { printf '\033[1;31m[错误]\033[0m %s\n' "$*" >&2; }

# ---------------------------------------------------------------- 1. 探测 wine 与 Windows Python
if ! command -v wine >/dev/null 2>&1; then
    err "未找到 wine，请先安装 wine（sudo apt install wine）"
    exit 1
fi

WIN_PY_WINPATH="$(wine cmd /c "where python" 2>/dev/null | tr -d '\r' | head -1)"
if [ -z "$WIN_PY_WINPATH" ]; then
    err "未在 wine 环境中找到 Windows 版 Python，请先在 wine 中安装 Python（如官方 python-3.x 安装包）"
    exit 1
fi

# wine 使用正斜杠形式的 Windows 路径调用更稳
WIN_PY="$(printf '%s' "$WIN_PY_WINPATH" | sed 's|\\|/|g')"
log "使用 Windows Python：$WIN_PY (经 wine)"
if ! wine "$WIN_PY" -c "import sys; print(sys.version)" >/dev/null 2>&1; then
    err "Windows Python 无法通过 wine 运行：$WIN_PY"
    exit 1
fi

# ---------------------------------------------------------------- 2. 安装依赖（幂等）
# 注1：官方 PyPI 源在 wine 下大文件下载易卡死，使用清华镜像加速。
# 注2：PyQt6 6.11.x 的 Windows wheel 存在上游缺陷——Qt6Core.dll 硬依赖 icuuc.dll
#      但未随包分发，导致任何环境（含真实 Windows）import QtCore 即崩溃。
#      故固定使用 6.9.1 + Qt6 6.9.2（经实测无此问题）。
PIP_INDEX="https://pypi.tuna.tsinghua.edu.cn/simple"
if wine "$WIN_PY" -c "import PyQt6, PyInstaller" >/dev/null 2>&1 \
   && wine "$WIN_PY" -c "import PyQt6.QtCore" >/dev/null 2>&1; then
    log "PyQt6 / PyInstaller 已安装且 QtCore 可导入，跳过安装"
else
    log "安装 PyQt6 6.9.1、PyInstaller（首次约需数分钟，含约 70MB Qt 运行库；走清华镜像）..."
    wine "$WIN_PY" -m pip install --no-warn-script-location -i "$PIP_INDEX" \
        "pyqt6==6.9.1" "pyqt6-qt6==6.9.2" pyinstaller
fi

# ---------------------------------------------------------------- 3. 打包
log "PyInstaller 打包 $APP_NAME ..."
wine "$WIN_PY" -m PyInstaller --noconfirm --clean -F -w \
    --name "$APP_NAME" \
    --paths . \
    "$ENTRY_SCRIPT"

# 检查产物
ARTIFACT="dist/$APP_NAME.exe"
if [ ! -s "$ARTIFACT" ]; then
    err "打包失败：未生成可执行文件 $ARTIFACT"
    exit 1
fi
log "打包完成：$(ls -lh "$ARTIFACT" | awk '{print $5}')  $ARTIFACT"

# ---------------------------------------------------------------- 4. 可选自检
if [ "${1:-}" = "--命令行" ]; then
    log "自检：以 测试配置 走内置 --命令行 模式验证导出链路..."
    if [ ! -d "测试配置" ]; then
        err "自检需要 测试配置/ 目录，已跳过"
        exit 1
    fi
    # Windows exe 在 wine 下经 Z: 盘映射访问真实路径；Linux 软链对 Windows 程序不可见，故复制真实目录
    rm -rf "$FIXED_SRC_NAME" "$FIXED_OUT_DIR"
    cp -r 测试配置 "$FIXED_SRC_NAME"
    trap 'rm -rf "$FIXED_SRC_NAME" "$FIXED_OUT_DIR"' EXIT
    if wine "$ARTIFACT" --命令行; then
        log "自检通过：导出全部成功"
    else
        err "自检失败：导出存在错误（见上方输出）"
        exit 1
    fi
else
    log "完成（提示：加 --命令行 参数可同时执行自检验证）"
fi