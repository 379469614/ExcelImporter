"""文案配置便携导表工具入口。"""
import sys as 系统_库

import 导表路径规划


if "--项目目录" not in 系统_库.argv:
    系统_库.argv.extend(
        ["--项目目录", 导表路径规划.获取工具所在项目目录(系统_库.argv[0])]
    )

import 导出文案配置 as 入口


系统_库.exit(入口.主函数())
