"""导出项目的数据配置表。"""
import argparse as 参数解析_库
import os as 操作系统_库
import sys as 系统_库

import 导表路径规划
from 导表入口 import 导出上下文, 导出多个文件


def 主函数() -> int:
    参数解析器 = 参数解析_库.ArgumentParser(description="导出 4数据配置表格/ 中的数据配置")
    参数解析器.add_argument("--项目目录", help="项目目录；未指定时使用当前工作目录")
    参数 = 参数解析器.parse_args()
    项目目录 = 参数.项目目录 or 操作系统_库.path.dirname(
        操作系统_库.path.dirname(操作系统_库.path.abspath(__file__))
    )
    输入目录, 输出目录 = 导表路径规划.获取导表目录(项目目录)
    配置表列表 = 导表路径规划.收集配置表(输入目录)
    if not 配置表列表:
        raise ValueError(f"数据配置目录没有可导出的 xlsx 文件：{输入目录}")

    上下文 = 导出上下文()
    上下文.路径 = ",".join(配置表列表)
    上下文.文件夹 = 输出目录
    上下文.多进程数量 = 1
    print(f"数据配置输入目录：{输入目录}")
    print(f"数据配置输出目录：{输出目录}")
    导出多个文件(上下文)
    return 0


if __name__ == "__main__":
    try:
        系统_库.exit(主函数())
    except ValueError as 异常:
        print(f"导出失败：{异常}", file=系统_库.stderr)
        系统_库.exit(1)
