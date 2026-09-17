"""导表路径规划。

按照项目目录与 Godot 工程目录的约定，定位数据表输入目录及导出目录。
"""
import os as 操作系统_库


常量_数据表目录名 = "4数据配置表格"
常量_文案表目录名 = "文案表格"
常量_数据配置目录名 = "数据配置"
常量_工程标识文件名 = "project.godot"


def 获取项目目录(项目目录参数: str | None = None) -> str:
    """返回规范化后的项目目录；未传入时使用当前工作目录。"""
    项目目录 = 项目目录参数 or 操作系统_库.getcwd()
    项目目录 = 操作系统_库.path.abspath(项目目录)
    if not 操作系统_库.path.isdir(项目目录):
        raise ValueError(f"项目目录不存在：{项目目录}")
    return 项目目录


def 获取工具所在项目目录(工具路径: str) -> str:
    """从便携导表工具所在位置向上定位项目目录。"""
    当前目录 = 操作系统_库.path.dirname(操作系统_库.path.abspath(工具路径))
    while True:
        数据表目录 = 操作系统_库.path.join(当前目录, 常量_数据表目录名)
        if 操作系统_库.path.isdir(数据表目录):
            try:
                获取工程目录(当前目录)
                return 当前目录
            except ValueError:
                pass

        上级目录 = 操作系统_库.path.dirname(当前目录)
        if 上级目录 == 当前目录:
            break
        当前目录 = 上级目录

    raise ValueError(
        "无法从导表工具所在位置定位项目目录；"
        f"请传入 --项目目录，工具路径：{操作系统_库.path.abspath(工具路径)}"
    )


def 获取工程目录(项目目录: str) -> str:
    """在项目目录或其直属子目录中定位唯一的 Godot 工程目录。"""
    候选目录 = []
    if 操作系统_库.path.isfile(操作系统_库.path.join(项目目录, 常量_工程标识文件名)):
        候选目录.append(项目目录)

    for 名称 in sorted(操作系统_库.listdir(项目目录)):
        路径 = 操作系统_库.path.join(项目目录, 名称)
        if 操作系统_库.path.isdir(路径) and 操作系统_库.path.isfile(
            操作系统_库.path.join(路径, 常量_工程标识文件名)
        ):
            候选目录.append(路径)

    if len(候选目录) != 1:
        候选文本 = "、".join(候选目录) or "无"
        raise ValueError(
            f"项目目录中应存在唯一的 Godot 工程目录（含 {常量_工程标识文件名}），"
            f"当前候选：{候选文本}"
        )
    return 候选目录[0]


def 获取导表目录(项目目录参数: str | None = None, 是文案: bool = False) -> tuple[str, str]:
    """返回指定类型配置表的输入目录与对应工程输出目录。"""
    项目目录 = 获取项目目录(项目目录参数)
    工程目录 = 获取工程目录(项目目录)
    输入目录 = 操作系统_库.path.join(项目目录, 常量_数据表目录名)
    输出目录 = 操作系统_库.path.join(工程目录, 常量_数据配置目录名)
    if 是文案:
        输入目录 = 操作系统_库.path.join(输入目录, 常量_文案表目录名)
        输出目录 = 操作系统_库.path.join(输出目录, 常量_文案表目录名)
    return 输入目录, 输出目录


def 收集配置表(输入目录: str) -> list[str]:
    """收集输入目录顶层的有效 Excel 表，忽略 Office 临时文件。"""
    if not 操作系统_库.path.isdir(输入目录):
        raise ValueError(f"配置表目录不存在：{输入目录}")
    return [
        操作系统_库.path.join(输入目录, 文件名)
        for 文件名 in sorted(操作系统_库.listdir(输入目录))
        if 文件名.lower().endswith(".xlsx") and not 文件名.startswith((".~", "~$"))
    ]
