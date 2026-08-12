#!/usr/bin/env bash
# 导表工具 一键打包脚本（Linux + Windows + AI 调用脚本）
#
# 流程（依次执行）：
#   Linux  ：自举虚拟环境(.venv-pack) -> 安装 PyInstaller -> PyInstaller 打包 -> 输出 dist/导表工具Ubuntu
#   Windows：探测 wine 下的 Windows Python -> 安装 PyQt6 + PyInstaller -> PyInstaller 打包
#            -> 输出 dist/导表工具.exe，供其他开发伙伴在 Windows 上直接双击使用
#   AI脚本 ：内嵌生成逻辑（仅标准库，原 生成AI调用脚本.py 已内嵌，不再依赖独立脚本）
#            -> 输出 dist/导表工具_AI.py
# 幂等：重复执行不会重复建 venv、不会重复装依赖；同名产物由 --noconfirm --clean 直接覆盖导出。
#
# 用法：
#   ./打包工具.sh            依次打出 Linux 可执行文件、Windows exe 与 AI 调用脚本
#   ./打包工具.sh --命令行   打包后用测试配置分别跑两平台内置自检（验证导出链路），结束时清理临时文件
# 注意：自检需要项目内有 测试配置/ 目录；固定路径(5配置文件/)只在本脚本自检时以软链(Linux)/复制(Win)形式临时存在。
set -euo pipefail
cd "$(dirname "$0")"

MODE="${1:-}"

LINUX_APP_NAME="导表工具Ubuntu"
WIN_APP_NAME="导表工具"
VENV_DIR=".venv-pack"
ENTRY_SCRIPT="导表工具_图形界面.py"

# 自检用：把 测试配置 挂到程序写死的固定来源路径，验证后删除
FIXED_SRC_NAME="5配置文件"
FIXED_OUT_DIR="BouncyPinball"

log()  { printf '\033[1;36m[打包]\033[0m %s\n' "$*"; }
logw() { printf '\033[1;36m[打包Win]\033[0m %s\n' "$*"; }
err()  { printf '\033[1;31m[错误]\033[0m %s\n' "$*" >&2; }

# ---------------------------------------------------------------- Linux 打包（本机）
build_linux() {
    # 1. 自举虚拟环境
    if [ ! -x "$VENV_DIR/bin/python" ]; then
        log "创建虚拟环境 $VENV_DIR（继承系统 PyQt6）..."
        python3 -m venv --system-site-packages "$VENV_DIR"
    else
        log "虚拟环境已存在，跳过创建"
    fi

    # 2. 安装 PyInstaller
    if ! "$VENV_DIR/bin/python" -m PyInstaller --version >/dev/null 2>&1; then
        log "安装 PyInstaller..."
        "$VENV_DIR/bin/pip" install pyinstaller
    else
        log "PyInstaller 已安装，跳过安装"
    fi

    # 3. 打包
    log "PyInstaller 打包 $LINUX_APP_NAME ..."
    "$VENV_DIR/bin/python" -m PyInstaller --noconfirm --clean -F -w \
        --name "$LINUX_APP_NAME" \
        --paths . \
        "$ENTRY_SCRIPT"

    # 检查产物
    ARTIFACT="dist/$LINUX_APP_NAME"
    if [ ! -x "$ARTIFACT" ]; then
        err "Linux 打包失败：未生成可执行文件 $ARTIFACT"
        exit 1
    fi
    log "Linux 打包完成：$(ls -lh "$ARTIFACT" | awk '{print $5}')  $ARTIFACT"

    # 4. 可选自检
    if [ "$MODE" = "--命令行" ]; then
        log "自检：以 测试配置 走内置 --命令行 模式验证导出链路..."
        if [ ! -d "测试配置" ]; then
            err "自检需要 测试配置/ 目录，已跳过"
            exit 1
        fi
        # 程序写死固定相对路径(上一级目录)；打包产物在 dist/ 下，用软链把测试配置挂到固定来源路径
        ln -sfn 测试配置 "$FIXED_SRC_NAME"
        trap 'rm -f "$FIXED_SRC_NAME"; rm -rf "$FIXED_OUT_DIR"' EXIT
        if "$ARTIFACT" --命令行; then
            log "Linux 自检通过：导出全部成功"
        else
            err "Linux 自检失败：导出存在错误（见上方输出）"
            exit 1
        fi
    else
        log "Linux 完成（提示：加 --命令行 参数可同时执行自检验证）"
    fi
}

# ---------------------------------------------------------------- Windows 打包（wine）
build_windows() {
    # 1. 探测 wine 与 Windows Python
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
    logw "使用 Windows Python：$WIN_PY (经 wine)"
    if ! wine "$WIN_PY" -c "import sys; print(sys.version)" >/dev/null 2>&1; then
        err "Windows Python 无法通过 wine 运行：$WIN_PY"
        exit 1
    fi

    # 2. 安装依赖（幂等）
    # 注1：官方 PyPI 源在 wine 下大文件下载易卡死，使用清华镜像加速。
    # 注2：PyQt6 6.11.x 的 Windows wheel 存在上游缺陷——Qt6Core.dll 硬依赖 icuuc.dll
    #      但未随包分发，导致任何环境（含真实 Windows）import QtCore 即崩溃。
    #      故固定使用 6.9.1 + Qt6 6.9.2（经实测无此问题）。
    PIP_INDEX="https://pypi.tuna.tsinghua.edu.cn/simple"
    if wine "$WIN_PY" -c "import PyQt6, PyInstaller" >/dev/null 2>&1 \
       && wine "$WIN_PY" -c "import PyQt6.QtCore" >/dev/null 2>&1; then
        logw "PyQt6 / PyInstaller 已安装且 QtCore 可导入，跳过安装"
    else
        logw "安装 PyQt6 6.9.1、PyInstaller（首次约需数分钟，含约 70MB Qt 运行库；走清华镜像）..."
        wine "$WIN_PY" -m pip install --no-warn-script-location -i "$PIP_INDEX" \
            "pyqt6==6.9.1" "pyqt6-qt6==6.9.2" pyinstaller
    fi

    # 3. 打包
    logw "PyInstaller 打包 $WIN_APP_NAME ..."
    wine "$WIN_PY" -m PyInstaller --noconfirm --clean -F -w \
        --name "$WIN_APP_NAME" \
        --paths . \
        "$ENTRY_SCRIPT"

    # 检查产物
    ARTIFACT="dist/$WIN_APP_NAME.exe"
    if [ ! -s "$ARTIFACT" ]; then
        err "Windows 打包失败：未生成可执行文件 $ARTIFACT"
        exit 1
    fi
    logw "Windows 打包完成：$(ls -lh "$ARTIFACT" | awk '{print $5}')  $ARTIFACT"

    # 4. 可选自检
    if [ "$MODE" = "--命令行" ]; then
        logw "自检：以 测试配置 走内置 --命令行 模式验证导出链路..."
        if [ ! -d "测试配置" ]; then
            err "自检需要 测试配置/ 目录，已跳过"
            exit 1
        fi
        # Windows exe 在 wine 下经 Z: 盘映射访问真实路径；Linux 软链对 Windows 程序不可见，故复制真实目录
        rm -rf "$FIXED_SRC_NAME" "$FIXED_OUT_DIR"
        cp -r 测试配置 "$FIXED_SRC_NAME"
        trap 'rm -rf "$FIXED_SRC_NAME" "$FIXED_OUT_DIR"' EXIT
        if wine "$ARTIFACT" --命令行; then
            logw "Windows 自检通过：导出全部成功"
        else
            err "Windows 自检失败：导出存在错误（见上方输出）"
            exit 1
        fi
    else
        logw "Windows 完成（提示：加 --命令行 参数可同时执行自检验证）"
    fi
}

# ---------------------------------------------------------------- AI 调用脚本（本机，仅标准库）
# 生成逻辑内嵌于此（原 生成AI调用脚本.py）：heredoc 直接交给 python 执行，
# 输出 dist/导表工具_AI.py，不依赖任何独立脚本文件。
build_ai_script() {
    AI_ARTIFACT="dist/导表工具_AI.py"
    log "生成 AI 调用脚本（内嵌全部导出引擎源码与 sxl 依赖）..."
    python3 - <<'PY'
import base64, os, zlib

工程目录 = os.getcwd()  # 打包脚本已 cd 到工程目录；内嵌后不再有 __file__
输出目录 = os.path.join(工程目录, "dist")
源码文件列表 = [
    "嵌套解析器.py",
    "导表输出器.py",
    "导表工具集.py",
    "导表核心.py",
    "导表入口.py",
    os.path.join("sxl", "__init__.py"),
    os.path.join("sxl", "sxl.py"),
]
模板 = r'''#!/usr/bin/env python3
"""导表工具 —— AI 可直接调用的自包含单文件脚本

本文件由 生成AI调用脚本.py 自动生成：内嵌了全部导出引擎源码与 sxl 依赖源码，
复制到任意目录即可独立使用，仅依赖 Python 3.10+ 标准库，无需安装第三方包、
无需携带任何工程文件。文件本身可以随意重命名。

命令行参数与 导表入口.py 完全一致，并额外支持 --json：
    -p  输入 excel 文件，用 , ; | 分隔
    -f  输出文件夹（缺省：上一级目录/BouncyPinball/数据配置）
    -e  导出格式：json / xml / lua / ycl
    -s  签名（控制列/表是否导出）
    -t  导出文件后缀
    -r  对象字段分隔符，默认 ;
    -m  多进程数量，默认 cpu 数
    -c  将表结构写入该 json 文件，供外部生成读取代码
    -h  打印帮助
    --json  在 stdout 末尾输出一行机器可读的 JSON 结果摘要（供 AI 等外部程序解析）

示例：
    python3 导表工具_AI.py -p "hero.xlsx,mount.xlsx" -f out -e json
    python3 导表工具_AI.py -p "text.xlsx" -f out -e lua --json
"""
import base64
import multiprocessing
import os
import shutil
import sys
import tempfile
import zlib

# 内嵌源码表：{ 相对路径: zlib 压缩 + base64 编码的源码 }。由生成脚本写入，请勿手工修改。
_内嵌文件 = {
__内嵌文件__
}


def _引导依赖() -> str:
    """将内嵌源码解压到临时目录并加入 sys.path，返回临时目录路径。"""
    临时目录 = tempfile.mkdtemp(prefix="导表工具_", suffix="")
    for 相对路径, 编码数据 in _内嵌文件.items():
        目标路径 = os.path.join(临时目录, 相对路径)
        目录 = os.path.dirname(目标路径)
        if 目录:
            os.makedirs(目录, exist_ok=True)
        with open(目标路径, "wb") as 文件:
            文件.write(zlib.decompress(base64.b64decode(编码数据)))
    sys.path.insert(0, 临时目录)
    return 临时目录


def 主函数() -> None:
    multiprocessing.freeze_support()  # Windows 下 multiprocessing 需要
    # 统一以 UTF-8 容错编码输出，避免 Windows 重定向 stdout 时中文/符号崩溃。
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    临时目录 = _引导依赖()
    try:
        import 导表入口
        # 注入本脚本所在目录，使默认导出路径与图形界面保持一致（上一级/BouncyPinball/数据配置）。
        导表入口.脚本所在目录 = os.path.dirname(os.path.abspath(__file__))
        导表入口.主函数()
    finally:
        # 仅主进程清理临时目录；SystemExit（含错误退出）也会执行到此处。
        if __name__ == "__main__":
            shutil.rmtree(临时目录, ignore_errors=True)


if __name__ == "__main__":
    主函数()
'''
条目列表 = []
for 相对路径 in 源码文件列表:
    完整路径 = os.path.join(工程目录, 相对路径)
    with open(完整路径, "rb") as 文件:
        源码 = 文件.read()
    编码数据 = base64.b64encode(zlib.compress(源码)).decode("ascii")
    键 = 相对路径.replace(os.sep, "/")
    条目列表.append(f"    {键!r}: {编码数据!r},")
内容 = 模板.replace("__内嵌文件__", "\n".join(条目列表))
输出路径 = os.path.join(输出目录, "导表工具_AI.py")
os.makedirs(输出目录, exist_ok=True)
with open(输出路径, "w", encoding="utf-8", newline="\n") as 文件:
    文件.write(内容)
print(f"已生成：{输出路径}（{os.path.getsize(输出路径)} 字节）")
PY
    if [ ! -s "$AI_ARTIFACT" ]; then
        err "AI 脚本生成失败：未生成 $AI_ARTIFACT"
        exit 1
    fi
    chmod +x "$AI_ARTIFACT"
    log "AI 脚本生成完成：$(ls -lh "$AI_ARTIFACT" | awk '{print $5}')  $AI_ARTIFACT"
}

# ---------------------------------------------------------------- 主流程：依次打包两个平台 + AI 脚本
build_linux
build_windows
build_ai_script

log "全部完成：dist/导表工具Ubuntu（Linux）、dist/导表工具.exe（Windows）与 dist/导表工具_AI.py（AI 调用脚本）均已生成"