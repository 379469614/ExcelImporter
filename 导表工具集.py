"""导表工具集

提供单元格值格式化、导出标记识别、签名匹配、文件新旧比较、模式信息组装等
通用工具函数，以及绑定类型数据结构，供导表核心引擎与图形界面共同使用。
"""
import os
import re
import string

import sxl
import 嵌套解析器


class 绑定类型:
    """标记基础类型对应的外部引用键，携带来源标记与字段名，供代码生成时查找定义。"""

    def __init__(self, 类型: str):
        self.类型名: str = 类型
        self.标记: str = ""
        self.字段: str = ""

    def __eq__(self, 其他) -> bool:
        return self.类型名 == 其他


# 将解析出的值填入父容器：列表追加，字典按名称赋值；模式阶段校验标识符合法性。
def 填充值(父容器: list | dict, 名称: str, 值, 是模式: bool) -> None:
    if isinstance(父容器, list):
        父容器.append(值)
    else:
        if 是模式 and not re.match(r"^_|[a-zA-Z\u4e00-\u9fa5]\w*$", 名称):
            raise ValueError(f"{名称} 是不合法的标识符")
        父容器[名称] = 值


# 在信息列表中查找指定名称首次出现的索引，未找到时返回 -1。
def 获取索引(信息列表: list, 名称: str) -> int:
    return next((索引 for 索引, 项 in enumerate(信息列表) if 项 == 名称), -1)


# 将单元格值转换为字符串，空值转换为空字符串。
def 获取单元格值(值) -> str:
    return str(值) if 值 is not None else ""


# 将基础类型转换为模式信息列表，描述存在时追加描述，绑定类型则取其类型名。
def 获取模式信息(类型, 描述: str | None) -> list:
    if isinstance(类型, 绑定类型):
        类型 = 类型.类型名
    return [类型, 描述] if 描述 else [类型]


# 从以"|"开头的工作表名称中识别导出标记，无匹配标记时返回 False。
def 获取导出标记(工作表名称: str) -> str | bool:
    匹配 = re.match(r"\|[" + string.whitespace + r"]*(\w+)", 工作表名称)
    return 匹配.group(1) if 匹配 else False


# 判断单元格签名是否与当前导出签名匹配，未指定签名时视为全部匹配。
def 是否为签名匹配(签名参数: str | None, 签名: str) -> bool:
    if 签名参数 is None:
        return True
    return True if [子项 for 子项 in re.split(r"[/\\, :]", 签名) if 子项 in 签名参数] else False


# 组合根名称与格式生成导出文件名，并拼接到导出文件夹下。
def 生成导出文件路径(根名称: str, 格式: str, 导出文件夹: str) -> str:
    文件名 = 根名称 + "." + 格式
    return os.path.join(导出文件夹, 文件名)


# 将字段声明文本拆分为类型与字段名的二元组。
def 分割空白字段(文本: str) -> tuple[str, str]:
    return 嵌套解析器.分割字段声明(文本)


# 将 1 起始列索引转换为 Excel 列字母：1→A、27→AA。
def 获取列字母(列号: int) -> str:
    字母 = ""
    while 列号 > 0:
        列号, 余数 = divmod(列号 - 1, 26)
        字母 = chr(余数 + ord("A")) + 字母
    return 字母


# 预扫各 xlsx 带 | 标记的 sheet，返回（冲突标记集合, 冲突提示列表）。
# 仅跨文件同名标记视为冲突；无 | 标记的 sheet 完全不参与判断。
def 查找同名标记冲突(xlsx路径列表: list) -> tuple:
    标记到来源 = {}
    顺序 = []
    for 路径 in xlsx路径列表:
        try:
            工作簿 = sxl.Workbook(路径)
        except Exception:
            continue
        文件名 = os.path.basename(路径)
        for 名称 in 工作簿.sheets:
            if isinstance(名称, str):
                标记 = 获取导出标记(名称)
                if 标记:
                    if 标记 not in 标记到来源:
                        标记到来源[标记] = set()
                        顺序.append(标记)
                    标记到来源[标记].add(文件名)

    冲突标记 = {标记 for 标记 in 顺序 if len(标记到来源[标记]) > 1}
    冲突提示列表 = [
        f"{'和'.join(sorted(标记到来源[标记]))}存在同名sheet：|{标记}，无法导出，请修改sheet名字后重新导出"
        for 标记 in 顺序
        if 标记 in 冲突标记
    ]
    return 冲突标记, 冲突提示列表