#!/usr/bin/env python3
"""生成AI调用脚本

读取工程内导出引擎的全部源码（含 sxl 依赖），压缩编码后生成一个单文件、
自包含、可独立复制的 AI 调用脚本（默认输出到工程目录下的 导表工具_AI.py）。

用法：
    python3 生成AI调用脚本.py                 # 生成 ./导表工具_AI.py
    python3 生成AI调用脚本.py /tmp/xx.py      # 生成到指定路径
"""
import base64
import os
import sys
import zlib

工程目录 = os.path.dirname(os.path.abspath(__file__))

# 需要内嵌进单文件脚本的源码文件（工程内相对路径）
源码文件列表 = [
    "嵌套解析器.py",
    "导表输出器.py",
    "导表工具集.py",
    "导表核心.py",
    "导表入口.py",
    os.path.join("sxl", "__init__.py"),
    os.path.join("sxl", "sxl.py"),
]

# 单文件脚本模板；__内嵌文件__ 占位符会被替换为编码后的源码条目。
模板 = r'''#!/usr/bin/env python3
"""导表工具 —— AI 可直接调用的自包含单文件脚本

本文件由 生成AI调用脚本.py 自动生成：内嵌了全部导出引擎源码与 sxl 依赖源码，
复制到任意目录即可独立使用，仅依赖 Python 3.10+ 标准库，无需安装第三方包、
无需携带任何工程文件。文件本身可以随意重命名。

命令行参数与 导表入口.py 完全一致，并额外支持 --json：
    -p  输入 excel 文件，用 , ; | 分隔
    -f  输出文件夹
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
        导表入口.主函数()
    finally:
        # 仅主进程清理临时目录；SystemExit（含错误退出）也会执行到此处。
        if __name__ == "__main__":
            shutil.rmtree(临时目录, ignore_errors=True)


if __name__ == "__main__":
    主函数()
'''


def 主函数() -> None:
    输出路径 = sys.argv[1] if len(sys.argv) > 1 else os.path.join(工程目录, "导表工具_AI.py")
    条目列表 = []
    for 相对路径 in 源码文件列表:
        完整路径 = os.path.join(工程目录, 相对路径)
        with open(完整路径, "rb") as 文件:
            源码 = 文件.read()
        编码数据 = base64.b64encode(zlib.compress(源码)).decode("ascii")
        键 = 相对路径.replace(os.sep, "/")
        条目列表.append(f"    {键!r}: {编码数据!r},")
    内容 = 模板.replace("__内嵌文件__", "\n".join(条目列表))
    输出路径 = os.path.abspath(输出路径)
    with open(输出路径, "w", encoding="utf-8", newline="\n") as 文件:
        文件.write(内容)
    print(f"已生成：{输出路径}（{os.path.getsize(输出路径)} 字节）")


if __name__ == "__main__":
    主函数()
