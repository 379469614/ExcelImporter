"""导表核心引擎

实现单个 Excel 工作簿的导出流程：遍历工作表、识别导出标记、解析数据表与
配置表、构建导出对象，并按上下文指定的格式保存输出文件。Apache License 2.0
版权归原 proton 项目作者 YANG Huan 所有，本文件为其中文规范重写版本。
"""
import codecs
import collections
import json
import os
import re
import string

import sxl
import 嵌套解析器
import 导表工具集
import 导表输出器


class 约束:
    """保存标记与字段名的约束信息，供导出过程中的条件判断使用。"""

    def __init__(self, 标记: str, 字段: str):
        self.标记: str = 标记
        self.字段: str = 字段


class 导出器:
    """按上下文配置导出工作簿，维护导出记录列表并统一执行保存。"""

    配置表标题: tuple[str, ...] = ("name", "value", "type", "sign", "description")
    最大空行数: int = 3

    def __init__(self, 上下文):
        self.上下文 = 上下文
        self.记录列表: list[导表输出器.导出记录] = []

    # 将类型特定转义符还原为原始字符：字符串内的换行、逗号与对象分隔符。
    def 检查字符串转义(self, 类型: str, 值: str) -> str:
        return 值 if not 值 or "string" not in 类型 else 值.replace("\\n", "\n").replace("\\,", "\0").replace("\\" + self.上下文.对象分隔符, "\a")

    # 还原类型转义时的占位符，恢复为逗号与对象分隔符。
    def 还原字符串转义(self, 字符串: str) -> str:
        return 字符串.replace("\0", ",").replace("\a", self.上下文.对象分隔符)

    # 根据是否禁用复数命名返回条目名称或名称加 s 的列表形式。
    def 复数化名称(self, 名称: str) -> str:
        return 名称 if self.上下文.禁用复数 else 名称 + "s"

    # 将字段类型文本解析为内部类型标识：列表、对象、基础类型或绑定类型。
    def 获取类型(self, 类型: str):
        if 类型[-2] == "[" and 类型[-1] == "]":
            return "list"
        if 类型[0] == "{" and 类型[-1] == "}":
            return "obj"
        if 类型 in ("int", "double", "string", "bool", "long", "float"):
            return 类型

        匹配 = re.search(r"(int|string|long)[" + string.whitespace + r"]*\((\S+)\.(\S+)\)", 类型)
        if 匹配:
            绑定 = 导表工具集.绑定类型(匹配.group(1))
            绑定.标记 = 匹配.group(2)
            绑定.字段 = 匹配.group(3)
            return 绑定

        raise ValueError(f"{类型} 不是合法类型")

    # 将列表类型字段解析为列表对象，模式阶段仅生成模式信息，数据阶段逐项解析。
    def 构建列表表达式(self, 父容器, 类型: str, 名称: str, 值, 是模式: bool) -> None:
        基础类型: str = 类型[:-2]
        列表: list = []
        if 是模式:
            self.构建表达式(列表, 基础类型, 名称, None, 是模式)
            列表 = 导表工具集.获取模式信息(列表[0], 值)
        else:
            值列表 = 嵌套解析器.分割列表值(值)
            for 单项值 in 值列表:
                self.构建表达式(列表, 基础类型, 名称, 单项值, False, True)

        导表工具集.填充值(父容器, self.复数化名称(名称), 列表, 是模式)

    # 将对象类型字段解析为有序字典对象，模式阶段仅生成模式信息，数据阶段逐字段解析。
    def 构建对象表达式(self, 父容器, 类型: str, 名称: str, 值, 是模式: bool) -> None:
        对象 = collections.OrderedDict()
        字段类型列表 = 嵌套解析器.分割对象类型字段(类型, self.上下文.对象分隔符)

        if 是模式:
            for 索引 in range(0, len(字段类型列表)):
                字段类型, 字段名称 = 导表工具集.分割空白字段(字段类型列表[索引])
                self.构建表达式(对象, 字段类型, 字段名称, None, 是模式)
            对象 = 导表工具集.获取模式信息(对象, 值)
        else:
            字段值列表 = 嵌套解析器.分割对象值(值, self.上下文.对象分隔符)
            for 索引 in range(0, len(字段类型列表)):
                if 索引 < len(字段值列表):
                    字段类型, 字段名称 = 导表工具集.分割空白字段(字段类型列表[索引])
                    self.构建表达式(对象, 字段类型, 字段名称, 字段值列表[索引], False, True)

        导表工具集.填充值(父容器, 名称, 对象, 是模式)

    # 将基础类型字段解析为对应 Python 值，模式阶段生成模式信息，数据阶段做类型转换。
    def 构建基础表达式(self, 父容器, 类型: str, 名称: str, 值, 是模式: bool, 在对象内: bool) -> None:
        类型名 = self.获取类型(类型)
        if 是模式:
            值 = 导表工具集.获取模式信息(类型名, 值)
        else:
            if 类型名 != "string" and 值.isspace():
                return

            if 类型名 == "int" or 类型名 == "long":
                值 = int(float(值))
            elif 类型名 == "double" or 类型名 == "float":
                值 = float(值)
            elif 类型名 == "string":
                if 值.endswith(".0"):
                    try:
                        值 = str(int(float(值)))
                    except ValueError:
                        值 = self.还原字符串转义(str(值))
                else:
                    值 = self.还原字符串转义(str(值))
                if 在对象内 and len(值) > 0 and 值[0] == "\n":
                    值 = 值[1:]
            elif 类型名 == "bool":
                try:
                    值 = int(float(值))
                    值 = False if 值 == 0 else True
                except ValueError:
                    值 = 值.lower()
                    if 值 in ("false", "no", "off"):
                        值 = False
                    elif 值 in ("true", "yes", "on"):
                        值 = True
                    else:
                        raise ValueError(f"{值} 是不合法的布尔值")

        导表工具集.填充值(父容器, 名称, 值, 是模式)

    # 按类型分发到列表、对象或基础表达式构建，统一表达式入口。
    def 构建表达式(self, 父容器, 类型: str, 名称: str, 值, 是模式: bool = False, 在对象内: bool = False) -> None:
        类型名 = self.获取类型(类型)
        if 类型名 == "list":
            self.构建列表表达式(父容器, 类型, 名称, 值, 是模式)
        elif 类型名 == "obj":
            self.构建对象表达式(父容器, 类型, 名称, 值, 是模式)
        else:
            self.构建基础表达式(父容器, 类型, 名称, 值, 是模式, 在对象内)

    # 根据是否条目模式组合根名称，列表条目追加复数后缀与扩展名。
    def 获取根名称(self, 导出标记: str, 是条目: bool) -> str:
        根名称 = self.复数化名称(导出标记) if 是条目 else 导出标记
        return 根名称 + (self.上下文.扩展名 or "")

    # 导出单个工作簿：遍历字符串键名工作表，识别标记并收集导出记录。
    def 导出(self, 路径: str) -> list[dict]:
        self.路径 = 路径
        数据 = sxl.Workbook(self.路径)
        合并输出 = None

        for 工作表名称 in [名称 for 名称 in 数据.sheets if type(名称) is str]:
            self.工作表名称 = 工作表名称
            导出标记 = 导表工具集.获取导出标记(工作表名称)
            if 导出标记:
                工作表 = 数据.sheets[工作表名称]
                是否合并 = 工作表名称.endswith("<<")
                配置标题信息 = self.获取配置表标题信息(工作表)
                if not 配置标题信息:
                    根名称 = self.获取根名称(导出标记, not 是否合并)
                    条目名称 = 导出标记
                else:
                    根名称 = self.获取根名称(导出标记, False)
                    条目名称 = None

                if not 合并输出:
                    self.检查工作表名重复(self.路径, 工作表名称, 根名称)
                    导出文件 = 导表工具集.生成导出文件路径(根名称, self.上下文.格式, self.上下文.文件夹)

                    if 导表工具集.是否过期(self.路径, 导出文件):
                        if 条目名称:
                            导出对象 = self.导出条目工作表(工作表)
                        else:
                            导出对象 = self.导出配置表(工作表, 配置标题信息)

                        if 是否合并:
                            if not 条目名称:
                                合并输出 = 导出对象
                            else:
                                合并输出 = (collections.OrderedDict(), collections.OrderedDict())
                                条目键 = self.复数化名称(条目名称)
                                合并输出[0][条目键] = [[导出对象[0]]]
                                条目名称 = None
                                导出对象 = 合并输出
                                对象 = 导出对象[1]
                                if 对象:
                                    合并输出[1][条目键] = 对象

                        self.记录列表.append(导表输出器.导出记录(self.路径, 工作表, 导出文件, 根名称, 条目名称, 导出对象, 导出标记))
                    else:
                        print(f"{self.路径} 未发生变化")
                        break
                else:
                    if 条目名称:
                        导出对象 = self.导出条目工作表(工作表)
                        合并输出[0][self.复数化名称(条目名称)] = [[导出对象[0]]]
                        对象 = 导出对象[1]
                        if 对象:
                            合并输出[1][self.复数化名称(条目名称)] = 对象
                    else:
                        导出对象 = self.导出配置表(工作表, 配置标题信息)
                        合并输出[0].update(导出对象[0])
                        对象 = 导出对象[1]
                        if 对象:
                            合并输出[1].update(对象)

        return self.保存全部()

    # 读取配置表首行标题并定位各列索引，缺少必要列时返回 None。
    def 获取配置表标题信息(self, 工作表):
        标题行 = 工作表.head(1)[0]

        名称索引 = 导表工具集.获取索引(标题行, self.配置表标题[0])
        值索引 = 导表工具集.获取索引(标题行, self.配置表标题[1])
        类型索引 = 导表工具集.获取索引(标题行, self.配置表标题[2])
        签名索引 = 导表工具集.获取索引(标题行, self.配置表标题[3])
        描述索引 = 导表工具集.获取索引(标题行, self.配置表标题[4])

        if 名称索引 != -1 and 值索引 != -1 and 类型索引 != -1:
            return (名称索引, 值索引, 类型索引, 签名索引, 描述索引)
        else:
            return None

    # 导出条目数据表：解析前四行标题信息，逐行构建条目对象并收集模式。
    def 导出条目工作表(self, 工作表) -> tuple[collections.OrderedDict, list]:
        行迭代 = iter(工作表.rows)
        描述行 = next(行迭代)
        类型行 = next(行迭代)
        名称行 = next(行迭代)
        签名行 = next(行迭代)

        列数 = len(类型行)
        标题信息列表 = []
        模式对象 = collections.OrderedDict()

        try:
            for 列索引 in range(列数):
                类型 = 导表工具集.获取单元格值(类型行[列索引]).strip()
                名称 = 导表工具集.获取单元格值(名称行[列索引]).strip()
                签名匹配 = 导表工具集.是否为签名匹配(self.上下文.签名, 导表工具集.获取单元格值(签名行[列索引]).strip())
                标题信息列表.append((类型, 名称, 签名匹配))

                if self.上下文.代码生成器:
                    if 类型 and 名称 and 签名匹配:
                        self.构建表达式(模式对象, 类型, 名称, 描述行[列索引], True)

        except Exception as 异常:
            异常.args += (f"{工作表.name} 存在标题错误，{列索引 + 1} 列 {类型} {名称} 于 {self.路径} 出错", "")
            raise 异常

        列表 = []
        存在导出 = next((信息 for 信息 in 标题信息列表 if 信息[0] and 信息[1] and 信息[2]), False)
        if 存在导出:
            try:
                空行计数 = 0
                self.行索引 = 3
                for 行 in 行迭代:
                    self.行索引 += 1

                    条目 = collections.OrderedDict()
                    首列文本 = 导表工具集.获取单元格值(行[0]).strip()
                    if not 首列文本:
                        空行计数 += 1
                        if 空行计数 >= self.最大空行数:
                            break

                    if not 首列文本 or 首列文本[0] == "#":
                        continue

                    跳过标记索引 = None
                    if 首列文本[0] == "!":
                        下一位置 = 首列文本.find("!", 1)
                        if 下一位置 >= 2:
                            签名标记 = 首列文本[1:下一位置]
                            if 导表工具集.是否为签名匹配(self.上下文.签名, 签名标记.strip()):
                                continue
                            else:
                                跳过标记索引 = len(签名标记) + 2

                    for self.列索引 in range(列数):
                        签名匹配 = 标题信息列表[self.列索引][2]
                        if 签名匹配:
                            类型 = 标题信息列表[self.列索引][0]
                            名称 = 标题信息列表[self.列索引][1]
                            值 = 导表工具集.获取单元格值(行[self.列索引])

                            if 跳过标记索引 and self.列索引 == 0:
                                值 = 值.lstrip()[跳过标记索引:]

                            if 类型 and 名称 and 值:
                                self.构建表达式(条目, 类型, 名称, self.检查字符串转义(类型, 值))
                        空行计数 = 0

                    if 条目:
                        列表.append(条目)

            except Exception as 异常:
                异常.args += (f"{工作表.name} 在 {self.行索引 + 1} 行 {self.列索引 + 1}（{名称}）处于 {self.路径} 出错", "")
                raise 异常

        return (模式对象, 列表)

    # 导出配置表：逐行解析名称、值、类型与签名，构建配置对象并收集模式。
    def 导出配置表(self, 工作表, 标题索引: tuple) -> tuple[collections.OrderedDict, collections.OrderedDict]:
        行迭代 = iter(工作表.rows)
        next(行迭代)

        名称索引 = 标题索引[0]
        值索引 = 标题索引[1]
        类型索引 = 标题索引[2]
        签名索引 = 标题索引[3]
        描述索引 = 标题索引[4]

        模式对象 = collections.OrderedDict()
        对象 = collections.OrderedDict()

        try:
            空行计数 = 0
            self.行索引 = 0
            for 行 in 行迭代:
                self.行索引 += 1
                名称 = 导表工具集.获取单元格值(行[名称索引]).strip()
                值 = 导表工具集.获取单元格值(行[值索引])
                类型 = 导表工具集.获取单元格值(行[类型索引]).strip()
                描述 = 导表工具集.获取单元格值(行[描述索引]).strip()

                if 签名索引 > 0:
                    签名 = 导表工具集.获取单元格值(行[签名索引]).strip()
                    if not 导表工具集.是否为签名匹配(self.上下文.签名, 签名):
                        continue

                if not 名称 and not 值 and not 类型:
                    空行计数 += 1
                    if 空行计数 >= self.最大空行数:
                        break
                    continue

                if 名称 and 类型:
                    if 名称[0] != "#":
                        if self.上下文.代码生成器:
                            self.构建表达式(模式对象, 类型, 名称, 描述, True)
                        if 值:
                            self.构建表达式(对象, 类型, 名称, self.检查字符串转义(类型, 值))
                    空行计数 = 0

        except Exception as 异常:
            异常.args += (f"{工作表.name} 在 {self.行索引 + 1} 行（{类型}，{名称}，{值}）于 {self.路径} 出错", "")
            raise 异常

        return (模式对象, 对象)

    # 遍历导出记录执行保存，并在启用代码生成器时收集模式摘要列表。
    def 保存全部(self) -> list[dict]:
        模式列表 = []
        for 记录 in self.记录列表:
            if 记录.对象:
                self.保存(记录)

                if self.上下文.代码生成器:
                    模式列表.append({
                        "path": 记录.路径,
                        "exportfile": 记录.导出文件,
                        "root": 记录.根名称,
                        "item": 记录.条目名称 or 记录.导出标记,
                        "schema": 记录.模式,
                    })

        return 模式列表

    # 按上下文格式将记录对象写入导出文件，JSON/XML/Lua/YCL 四种格式分支。
    def 保存(self, 记录: 导表输出器.导出记录) -> None:
        if not 记录.对象:
            return

        if not os.path.isdir(self.上下文.文件夹):
            os.makedirs(self.上下文.文件夹)

        if self.上下文.格式 == "json":
            json字符串 = json.dumps(记录.对象, ensure_ascii=False, indent=2)
            with codecs.open(记录.导出文件, "w", "utf-8") as 文件:
                文件.write(json字符串)
            print("保存 %s 从 %s 到 %s" % (记录.导出文件, 记录.工作表.name, 记录.路径))

        elif self.上下文.格式 == "xml":
            if 记录.条目名称:
                记录.对象 = {self.复数化名称(记录.条目名称): 记录.对象}
            导表输出器.保存为XML(记录, self.上下文.禁用复数)

        elif self.上下文.格式 == "lua":
            lua字符串 = "".join(导表输出器.转为Lua(记录.对象))
            with codecs.open(记录.导出文件, "w", "utf-8") as 文件:
                文件.write("return ")
                文件.write(lua字符串)
            print("保存 %s 从 %s 到 %s" % (记录.导出文件, 记录.工作表.name, 记录.路径))

        elif self.上下文.格式 == "ycl":
            生成器 = 导表输出器.转为YCL(记录.对象)
            next(生成器)
            ycl字符串 = "".join(生成器)
            with codecs.open(记录.导出文件, "w", "utf-8") as 文件:
                文件.write(ycl字符串)
            print("保存 %s 从 %s 到 %s" % (记录.导出文件, 记录.工作表.name, 记录.路径))

    # 检查已收集记录中是否存在同名根名称，重复时抛错提示来源文件。
    def 检查工作表名重复(self, 路径: str, 工作表名称: str, 根名称: str) -> None:
        已有记录 = next((记录 for 记录 in self.记录列表 if 记录.根名称 == 根名称), False)
        if 已有记录:
            raise ValueError(f"{根名称} 在 {路径} 已定义于 {已有记录.路径}")