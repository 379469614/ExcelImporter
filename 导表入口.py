"""导表入口

提供导出参数上下文、多文件批量导出、异常收集与命令行参数解析入口。
支持以脚本方式运行本文件完成 Excel 批量转 JSON/XML/Lua/YCL。
"""
import getopt
import json
import multiprocessing
import os
import re
import sys
import traceback

import 导表工具集
from 导表核心 import 导出器


class 导出上下文:
    """保存命令行与调用方传入的导出参数，供导出器解析与输出使用。

    usage python 导表入口.py [-p filelist] [-f outfolder] [-e format]
    Arguments
    -p      : input excel files, use , or ; or space to separate
    -f      : out folder
    -e      : format, json or xml or lua or ycl

    Options
    -s      ：sign, controls whether the column is exported, defalut all export
    -t      : suffix, export file suffix
    -r      : the separator of object field, default is ; you can use it to change
    -m      : use the count of multiprocesses to export, default is cpu count
    -c      : a file path, save the excel structure to json
              the external program uses this file to automatically generate the read code
    -h      : print this help message and exit

    https://github.com/yanghuan/proton
    """

    def __init__(self):
        self.路径: str | None = None
        self.文件夹: str = "."
        self.格式: str = "json"
        self.签名: str | None = None
        self.扩展名: str | None = None
        self.对象分隔符: str = ";"
        self.代码生成器: str | None = None
        self.多进程数量: int | None = None


# 构建默认上下文并封装导出器单文件导出，异常时返回堆栈文本。
def 导出单个文件(上下文: 导出上下文, 路径: str):
    try:
        return 导出器(上下文).导出(路径)
    except Exception as 异常:
        return traceback.format_exc()


# 多进程包装函数，供进程池映射调用。
def 导出打包(参数: tuple) -> list | str:
    return 导出单个文件(参数[0], 参数[1])


# 批量导出多个文件：解析路径列表、分发导出、收集错误与模式信息。
def 导出多个文件(上下文: 导出上下文) -> None:
    路径列表: list[str] = []
    for 路径 in re.split(r"[,;|]+", 上下文.路径.strip()):
        if 路径:
            if not os.path.isfile(路径):
                raise ValueError(f"{路径} 不存在")
            elif 路径 in 路径列表:
                raise ValueError(f"{路径} 已存在")
            路径列表.append(路径)

    错误列表: list = []
    模式列表: list = []

    def 追加结果(结果) -> None:
        if type(结果) is str:
            错误列表.append(结果)
        else:
            模式列表.extend(结果)

    if 上下文.多进程数量 is None or 上下文.多进程数量 > 1:
        with multiprocessing.Pool(上下文.多进程数量) as 进程池:
            for 结果 in 进程池.map(导出打包, [(上下文, 路径) for 路径 in 路径列表]):
                追加结果(结果)
    else:
        for 路径 in 路径列表:
            结果 = 导出单个文件(上下文, 路径)
            追加结果(结果)

    if 模式列表:
        if 上下文.代码生成器:
            模式json字符串 = json.dumps(模式列表, ensure_ascii=False, indent=2)
            目录 = os.path.dirname(上下文.代码生成器)
            if 目录 and not os.path.isdir(目录):
                os.makedirs(目录)
            with open(上下文.代码生成器, "w", encoding="utf-8", newline="\n") as 文件:
                文件.write(模式json字符串.rstrip("\n"))

        已导出文件列表: list = []
        for 模式 in 模式列表:
            导出文件 = 模式["exportfile"]
            已有记录 = next((记录 for 记录 in 已导出文件列表 if 记录["exportfile"] == 导出文件), False)
            if 已有记录:
                错误列表.append("%s 在 %s 已定义于 %s" % (模式["root"], 模式["path"], 已有记录["path"]))
                os.remove(导出文件)
            else:
                已导出文件列表.append(模式)

    if 错误列表:
        print("\n\n".join(错误列表))
        sys.exit(-1)

    print("导出完成成功！！！")


# 解析命令行参数并填充上下文，随后执行批量导出。
def 主函数() -> None:
    print("argv:", sys.argv)
    选项列表, 参数列表 = getopt.getopt(sys.argv[1:], "p:f:e:s:t:r:m:c:h")

    上下文 = 导出上下文()
    上下文.路径 = None
    上下文.文件夹 = "."
    上下文.格式 = "json"
    上下文.签名 = None
    上下文.扩展名 = None
    上下文.对象分隔符 = ";"
    上下文.代码生成器 = None
    上下文.多进程数量 = None

    for 选项, 值 in 选项列表:
        if 选项 == "-p":
            上下文.路径 = 值
        elif 选项 == "-f":
            上下文.文件夹 = 值
        elif 选项 == "-e":
            上下文.格式 = 值.lower()
        elif 选项 == "-s":
            上下文.签名 = 值
        elif 选项 == "-t":
            上下文.扩展名 = 值
        elif 选项 == "-r":
            上下文.对象分隔符 = 值
        elif 选项 == "-m":
            上下文.多进程数量 = int(值) if 值 is not None else None
        elif 选项 == "-c":
            上下文.代码生成器 = 值
        elif 选项 == "-h":
            print(导出上下文.__doc__)
            sys.exit()

    if not 上下文.路径:
        print(导出上下文.__doc__)
        sys.exit(2)

    导出多个文件(上下文)


if __name__ == "__main__":
    主函数()