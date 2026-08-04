"""嵌套解析器

提供按顶层分隔符切割文本、剥离容器包围字符、拆分字段声明等解析能力，
供导表核心引擎解析 Excel 单元格中的列表、对象与字段类型声明表达式。
"""
import string


_常量_开括号到闭括号映射: dict[str, str] = {
    "{": "}",
    "[": "]",
    "(": ")",
}
_常量_闭括号集合: set[str] = set(_常量_开括号到闭括号映射.values())


# 按顶层分隔符切割文本，仅当分隔符位于所有括号嵌套之外时生效，并支持反斜杠转义。
def 分割顶层(文本: str | None, 分隔符: str, 是否跳过空值: bool = False) -> list[str]:
    if 文本 is None:
        return []
    if not 分隔符:
        raise ValueError("分隔符不能为空")

    子串列表: list[str] = []
    括号栈: list[str] = []
    起始位置: int = 0
    索引: int = 0
    文本长度: int = len(文本)

    while 索引 < 文本长度:
        当前字符: str = 文本[索引]

        if 当前字符 == "\\" and 索引 + 1 < 文本长度:
            索引 += 2
            continue

        if 当前字符 in _常量_开括号到闭括号映射:
            括号栈.append(_常量_开括号到闭括号映射[当前字符])
            索引 += 1
            continue

        if 当前字符 in _常量_闭括号集合:
            if not 括号栈 or 当前字符 != 括号栈[-1]:
                raise ValueError(f"{文本} 不是合法的嵌套表达式")
            括号栈.pop()
            索引 += 1
            continue

        if not 括号栈 and 文本.startswith(分隔符, 索引):
            子串: str = 文本[起始位置:索引]
            if not 是否跳过空值 or 子串:
                子串列表.append(子串)
            索引 += len(分隔符)
            起始位置 = 索引
            continue

        索引 += 1

    if 括号栈:
        raise ValueError(f"{文本} 不是合法的嵌套表达式")

    末尾子串: str = 文本[起始位置:]
    if not 是否跳过空值 or 末尾子串:
        子串列表.append(末尾子串)

    return 子串列表


# 剥离文本首尾的成对容器字符，首尾不匹配时原样返回输入文本。
def 剥离容器(文本: str, 起始字符: str, 结束字符: str) -> str:
    值: str = 文本.strip()
    if len(值) >= 2 and 值[0] == 起始字符 and 值[-1] == 结束字符:
        return 值[1:-1]
    return 值


# 按逗号切割方括号包裹的列表值，并跳过空项。
def 分割列表值(值: str) -> list[str]:
    return 分割顶层(剥离容器(值, "[", "]"), ",", True)


# 按指定分隔符切割花括号包裹的对象类型字段声明，并跳过空项。
def 分割对象类型字段(类型: str, 分隔符: str) -> list[str]:
    return 分割顶层(剥离容器(类型, "{", "}"), 分隔符, True)


# 按指定分隔符切割花括号包裹的对象值，并跳过空项。
def 分割对象值(值: str, 分隔符: str) -> list[str]:
    return 分割顶层(剥离容器(值, "{", "}"), 分隔符, True)


# 拆分字段声明为类型与字段名，类型部分以空白连接还原，返回 (类型, 字段名) 二元组。
def 分割字段声明(文本: str) -> tuple[str, str]:
    声明文本: str = 文本.strip()
    if not 声明文本:
        raise ValueError("字段声明不能为空")

    子串列表: list[str] = []
    括号栈: list[str] = []
    起始位置: int | None = None
    索引: int = 0
    文本长度: int = len(声明文本)

    while 索引 < 文本长度:
        当前字符: str = 声明文本[索引]

        if 当前字符 == "\\" and 索引 + 1 < 文本长度:
            if 起始位置 is None:
                起始位置 = 索引
            索引 += 2
            continue

        if 当前字符 in _常量_开括号到闭括号映射:
            if 起始位置 is None:
                起始位置 = 索引
            括号栈.append(_常量_开括号到闭括号映射[当前字符])
            索引 += 1
            continue

        if 当前字符 in _常量_闭括号集合:
            if not 括号栈 or 当前字符 != 括号栈[-1]:
                raise ValueError(f"{声明文本} 不是合法的字段声明")
            括号栈.pop()
            索引 += 1
            continue

        if 当前字符 in string.whitespace and not 括号栈:
            if 起始位置 is not None:
                子串列表.append(声明文本[起始位置:索引])
                起始位置 = None
            索引 += 1
            continue

        if 起始位置 is None:
            起始位置 = 索引
        索引 += 1

    if 括号栈:
        raise ValueError(f"{声明文本} 不是合法的字段声明")

    if 起始位置 is not None:
        子串列表.append(声明文本[起始位置:])

    if len(子串列表) < 2:
        raise ValueError(f"{声明文本} 不是合法的字段声明")

    return (" ".join(子串列表[:-1]), 子串列表[-1])