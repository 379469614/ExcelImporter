"""导表输出器

提供导出记录的承载结构，以及 XML、Lua、YCL 三种格式的序列化输出能力，
供导表核心引擎保存最终导出结果。
"""
import json
from io import StringIO
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ElementTree


class 导出记录:
    """保存单个工作表导出的路径、工作表引用、输出文件、根名称、条目名称、对象与导出标记。"""

    def __init__(self, 路径: str, 工作表, 导出文件: str, 根名称: str, 条目名称: str | None, 对象, 导出标记: str):
        self.路径: str = 路径
        self.工作表: object = 工作表
        self.导出文件: str = 导出文件
        self.根名称: str = 根名称
        self.条目名称: str | None = 条目名称
        self.导出标记: str = 导出标记
        self.设置对象(对象)

    # 将导出对象拆分为模式与数据两部分存入记录，空对象时两者均置为 None。
    def 设置对象(self, 对象) -> None:
        self.模式 = 对象[0] if 对象 else None
        self.对象 = 对象[1] if 对象 else None


# 生成长度为指定缩进层级的换行加空白前缀。
def 生成换行(缩进层级: int) -> str:
    return "\n" + "  " * 缩进层级


# 将基础类型的值作为属性或子元素写入 XML 父节点。
def 构建基础节点(父节点, 名称: str, 值) -> None:
    值 = str(值)
    if 父节点.tag == 名称:
        元素 = ElementTree.Element(名称)
        元素.text = 值
        父节点.append(元素)
    else:
        父节点.set(名称, 值)


# 将列表值构建为 XML 子节点序列，每条数据递归调用构建节点。
def 构建列表节点(父节点, 名称: str, 列表) -> None:
    元素 = ElementTree.Element(名称)
    父节点.append(元素)
    for 值 in 列表:
        构建节点(元素, 名称, 值)


# 将字典对象构建为 XML 子节点序列，每个键值递归调用构建节点。
def 构建对象节点(父节点, 名称: str, 对象) -> None:
    元素 = ElementTree.Element(名称)
    父节点.append(元素)
    for 键, 值 in 对象.items():
        构建节点(元素, 键, 值)


# 依据值类型分发到基础、列表或对象三种 XML 构建方式。
def 构建节点(父节点, 名称: str, 值) -> None:
    if isinstance(值, int) or isinstance(值, float) or isinstance(值, str):
        构建基础节点(父节点, 名称, 值)
    elif isinstance(值, list):
        构建列表节点(父节点, 名称, 值)
    elif isinstance(值, dict):
        构建对象节点(父节点, 名称, 值)


# 将导出记录的根对象序列化写入 XML 文件，并打印保存信息。
def 保存为XML(记录: 导出记录) -> None:
    书 = ElementTree.ElementTree()
    书.append = lambda 元素: 书._setroot(元素)
    构建节点(书, 记录.根名称, 记录.对象)

    xml字符串 = ElementTree.tostring(书.getroot(), "utf-8")
    文档 = minidom.parseString(xml字符串)
    缓冲 = StringIO()
    文档.writexml(缓冲, "", "  ", "\n", "utf-8")
    with open(记录.导出文件, "w", encoding="utf-8", newline="\n") as 文件:
        文件.write(缓冲.getvalue().rstrip("\n"))

    print("保存 %s 从 %s 到 %s" % (记录.导出文件, 记录.工作表.name, 记录.路径))


# 将任意值转换为 Lua 字面量文本生成器，字典按键输出 键 = 值 形式。
def 转为Lua(对象, 缩进层级: int = 1):
    if isinstance(对象, int) or isinstance(对象, float) or isinstance(对象, str):
        yield json.dumps(对象, ensure_ascii=False)
    else:
        yield "{"
        是列表 = isinstance(对象, list)
        是首个 = True
        for 项 in 对象:
            if 是首个:
                是首个 = False
            else:
                yield ","
            yield 生成换行(缩进层级)
            if not 是列表:
                键 = 项
                项 = 对象[键]
                yield 键
                yield " = "
            for 片段 in 转为Lua(项, 缩进层级 + 1):
                yield 片段
        yield 生成换行(缩进层级 - 1)
        yield "}"


# 将任意值转换为 YCL 文本生成器，字典按键输出 键 = 值 形式。
def 转为YCL(对象, 缩进层级: int = 0):
    是列表 = isinstance(对象, list)
    for 项 in 对象:
        yield 生成换行(缩进层级)
        if not 是列表:
            键 = 项
            项 = 对象[键]
            yield 键
        if isinstance(项, int) or isinstance(项, float) or isinstance(项, str):
            if not 是列表:
                yield " = "
            yield json.dumps(项, ensure_ascii=False)
        else:
            if not 是列表:
                yield " "
            yield "{"
            for 片段 in 转为YCL(项, 缩进层级 + 1):
                yield 片段
            yield 生成换行(缩进层级)
            yield "}"
