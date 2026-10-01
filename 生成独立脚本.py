"""将现有导表模块合并为两个只依赖标准库和外部 sxl 的纯源码脚本。"""
import ast
from pathlib import Path


源码目录 = Path(__file__).resolve().parent
公共模块 = (
    "嵌套解析器", "导表工具集", "导表输出器", "导表核心", "导表入口", "导表路径规划",
)
本地模块 = set(公共模块)


def 合并模块(模块名: str) -> str:
    """按语法节点移除本地模块引用，保留原源码的格式和注释。"""
    源码 = (源码目录 / f"{模块名}.py").read_text(encoding="utf-8-sig")
    行列表 = 源码.splitlines(keepends=True)
    行偏移 = [0]
    for 行 in 行列表:
        行偏移.append(行偏移[-1] + len(行))

    def 位置(行号: int, 字节列号: int) -> int:
        return 行偏移[行号 - 1] + len(行列表[行号 - 1].encode("utf-8")[:字节列号].decode("utf-8"))

    def 范围(节点) -> tuple[int, int]:
        return 位置(节点.lineno, 节点.col_offset), 位置(节点.end_lineno, 节点.end_col_offset)

    语法树 = ast.parse(源码)
    替换列表 = []
    删除节点 = []
    for 节点 in 语法树.body:
        if isinstance(节点, ast.Import):
            if all(别名.name in 本地模块 or 别名.name == "sxl" for 别名 in 节点.names):
                删除节点.append(节点)
        elif isinstance(节点, ast.ImportFrom) and 节点.module in 本地模块:
            删除节点.append(节点)
        elif 模块名 == "导表入口" and (
            isinstance(节点, ast.FunctionDef) and 节点.name == "主函数"
            or isinstance(节点, ast.If)
        ):
            # 发布入口只接受项目目录；不用通用入口的参数解析与启动逻辑。
            删除节点.append(节点)

    删除范围 = [范围(节点) for 节点 in 删除节点]
    替换列表.extend((开始, 结束, "") for 开始, 结束 in 删除范围)
    for 节点 in ast.walk(语法树):
        if isinstance(节点, ast.Attribute) and isinstance(节点.value, ast.Name) and 节点.value.id in 本地模块:
            开始, 结束 = 范围(节点)
            if not any(删除开始 <= 开始 < 删除结束 for 删除开始, 删除结束 in 删除范围):
                替换列表.append((开始, 结束, 节点.attr))
        elif 模块名.startswith("导出") and isinstance(节点, ast.Assign):
            if any(isinstance(目标, ast.Name) and 目标.id == "项目目录" for 目标 in 节点.targets):
                开始, 结束 = 范围(节点.value)
                替换列表.append((开始, 结束, "参数.项目目录 or 获取工具所在项目目录(__file__)"))

    for 开始, 结束, 文本 in sorted(替换列表, reverse=True):
        源码 = 源码[:开始] + 文本 + 源码[结束:]
    源码 = 源码.replace("项目目录；未指定时使用当前工作目录", "项目目录；未指定时从脚本所在位置向上查找")
    return f"# ===== {模块名} =====\n{源码.strip()}\n"


def 主函数() -> None:
    文件头 = '''#!/usr/bin/env python3
"""独立导表脚本：需要 Python 3.10+ 和第三方库 sxl。

本文件由生成独立脚本.py 合并现有源码生成，不依赖其他项目脚本。
导表核心源自 YANG Huan 的 proton 项目，遵循 Apache License 2.0。
"""
import sys

try:
    import sxl
except ModuleNotFoundError as 异常:
    if 异常.name != "sxl":
        raise
    print("缺少 SXL Python 第三方库，请先在当前 Python 环境中安装。", file=sys.stderr)
    print(f'安装命令："{sys.executable}" -m pip install sxl', file=sys.stderr)
    sys.exit(1)

'''
    公共源码 = "\n\n".join(合并模块(模块名) for 模块名 in 公共模块)
    for 类型 in ("数据", "文案"):
        输出路径 = 源码目录 / "发布包" / f"{类型}配置导表工具" / "导表工具.py"
        源码 = 文件头 + 公共源码 + "\n\n" + 合并模块(f"导出{类型}配置")
        # 写入前验证语法，避免留下不可执行的发布文件。
        ast.parse(源码)
        输出路径.parent.mkdir(parents=True, exist_ok=True)
        输出路径.write_text(源码, encoding="utf-8", newline="\n")
        print(f"已生成：{输出路径}")


if __name__ == "__main__":
    主函数()
