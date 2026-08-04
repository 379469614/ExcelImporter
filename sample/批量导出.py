"""批量导出示例

演示如何按客户端 / 服务器配置拆分导出 Excel 文件，并调用代码生成器。
原有代码保留平台判断与依赖工具路径，供示例工程复制后直接使用。
"""
import os
import platform
import shutil
import sys
import traceback

# 需要导出的公共配置文件（客户端、服务器都需要）
导出公共文件 = [
    "hero.xlsx",
    "mount.xlsx",
]

# 客户端额外需要导出的配置文件（仅客户端需要）
导出客户端专属 = [
    "text.xlsx",
]

# 服务器端额外需要导出的配置文件（仅服务器需要）
导出服务端专属 = [
]

# 以下内容请勿修改

导出脚本路径 = "../导表入口.py"
Python命令 = "tools\\py37\\py37.exe " if platform.system() == "Windows" else "python "


class 导出异常(Exception):
    pass


# 构造命令行并调用导表入口执行导出，失败时抛出导出异常。
def 导出文件(文件列表: list, 格式: str, 签名: str, 输出文件夹: str, 后缀: str, 模式文件: str) -> None:
    命令行 = ' -p "' + ",".join(文件列表) + '" -f ' + 输出文件夹 + " -e " + 格式 + " -s " + 签名
    if 后缀:
        命令行 += " -t " + 后缀
    if 模式文件:
        命令行 += " -c " + 模式文件
    命令行 = Python命令 + 导出脚本路径 + 命令行
    返回码 = os.system(命令行)
    if 返回码 != 0:
        raise 导出异常("导出 Excel 失败，请查看输出信息")


# 调用代码生成器处理模式文件，并清理临时模式文件。
def 生成代码(模式文件: str, 输出文件夹: str, 命名空间: str, 后缀: str) -> None:
    if os.path.exists(模式文件):
        命令行 = "tools\\CSharpGeneratorForProton\\CSharpGeneratorForProton.exe " + "-n " + 命名空间 + " -f " + 输出文件夹 + " -p " + 模式文件
        if 后缀:
            命令行 += " -t " + 后缀
        返回码 = os.system(命令行)
        os.remove(模式文件)
        if 返回码 != 0:
            raise 导出异常("代码生成失败，请查看输出信息")


# 导出服务器端配置：公共表加服务端专属表生成 JSON，并调用代码生成器。
def 导出服务器() -> None:
    导出文件(导出公共文件 + 导出服务端专属, "json", "server", "config_server", "Config", "schemaserver.json")
    生成代码("schemaserver.json", "config_server/ConfigGenerator/Template", "Ice.Project.Config", "Template")


# 导出客户端配置：公共表加客户端专属表生成 Lua。
def 导出客户端() -> None:
    导出文件(导出公共文件 + 导出客户端专属, "lua", "client", "config_client", "Template", None)


def 主函数() -> int:
    try:
        导出服务器()
        导出客户端()
        print("所有操作完成成功")
        return 0
    except 导出异常 as 异常:
        print(异常)
        print("发生错误，请查看日志，按回车键退出")
        input()
        return 1
    except Exception:
        traceback.print_exc()
        print("发生错误，请查看日志，按回车键退出")
        input()
        return 1


if __name__ == "__main__":
    sys.exit(主函数())