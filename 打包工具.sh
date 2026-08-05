#!/usr/bin/env bash
# 导表工具 一键打包脚本（Linux）
#
# 流程：自举虚拟环境(.venv-pack) -> 安装 PyInstaller -> PyInstaller 打包 -> 输出 dist/导表工具
# 幂等：重复执行不会重复建 venv、不会重复装 PyInstaller。
#
# 用法：
#   ./打包工具.sh            打包
#   ./打包工具.sh --命令行   打包后用测试配置跑一次内置自检（验证导出链路），结束时清理临时文件
# 注意：自检需要项目内有 测试配置/ 目录；固定路径(5配置文件/)只在本脚本自检时以软链形式临时存在。
set -euo pipefail
cd "$(dirname "$0")"

APP_NAME="导表工具"
VENV_DIR=".venv-pack"
ENTRY_SCRIPT="导表工具_图形界面.py"

# 自检用：把 测试配置 软链到程序写死的固定来源路径，验证后删除
FIXED_SRC_NAME="5配置文件"
FIXED_OUT_DIR="BouncyPinball"

log() { printf '\033[1;36m[打包]\033[0m %s\n' "$*"; }
err() { printf '\033[1;31m[错误]\033[0m %s\n' "$*" >&2; }

# ---------------------------------------------------------------- 1. 自举虚拟环境
if [ ! -x "$VENV_DIR/bin/python" ]; then
    log "创建虚拟环境 $VENV_DIR（继承系统 PyQt6）..."
    python3 -m venv --system-site-packages "$VENV_DIR"
else
    log "虚拟环境已存在，跳过创建"
fi

# ---------------------------------------------------------------- 2. 安装 PyInstaller
if ! "$VENV_DIR/bin/python" -m PyInstaller --version >/dev/null 2>&1; then
    log "安装 PyInstaller..."
    "$VENV_DIR/bin/pip" install pyinstaller
else
    log "PyInstaller 已安装，跳过安装"
fi

# ---------------------------------------------------------------- 3. 打包
log "PyInstaller 打包 $APP_NAME ..."
"$VENV_DIR/bin/python" -m PyInstaller --noconfirm --clean -F -w \
    --name "$APP_NAME" \
    --paths . \
    "$ENTRY_SCRIPT"

# 检查产物
ARTIFACT="dist/$APP_NAME"
if [ ! -x "$ARTIFACT" ]; then
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
    # 程序写死固定相对路径(上一级目录)；打包产物在 dist/ 下，用软链把测试配置挂到固定来源路径
    ln -sfn 测试配置 "$FIXED_SRC_NAME"
    trap 'rm -f "$FIXED_SRC_NAME"; rm -rf "$FIXED_OUT_DIR"' EXIT
    if "$ARTIFACT" --命令行; then
        log "自检通过：导出全部成功"
    else
        err "自检失败：导出存在错误（见上方输出）"
        exit 1
    fi
else
    log "完成（提示：加 --命令行 参数可同时执行自检验证）"
fi