#!/usr/bin/env python3
"""Python 3.10+；只读检查，仅使用标准库，不导出或修改任何表格。

默认检查本文件所在目录和文案表格目录；--目录可指定另一根目录。
--json 输出机器可读摘要；错误退出码为 1，只有提醒时为 0。
检查基于 xlsx-generation 规范及当前导表工具的转换行为。
不能恢复已经丢失的前导零，也不替代 Excel 或消费方的端到端验证。
"""
import argparse
from decimal import Decimal, InvalidOperation
import json
import math
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
中文 = re.compile(r'[\u3400-\u4dbf\u4e00-\u9fff\U00020000-\U0002fa1f]+')
问题 = []


def 报告(级别, 文件, 表, 格, 原因):
    问题.append(dict(级别=级别, 文件=str(文件), 工作表=表, 单元格=格, 原因=原因))


def 列名(列):
    名 = ''
    while 列:
        列, 余数 = divmod(列 - 1, 26)
        名 = chr(65 + 余数) + 名
    return 名


def 切分(文本, 分隔符, 跳空=False):
    栈, 结果, 起始, i = [], [], 0, 0
    配对 = {'[': ']', '{': '}', '(': ')'}
    while i < len(文本):
        字 = 文本[i]
        if 字 == '\\' and i + 1 < len(文本):
            i += 2
            continue
        if 字 in 配对:
            栈.append(配对[字])
        elif 字 in ']})':
            if not 栈 or 栈.pop() != 字:
                raise ValueError('括号不匹配；文本中的括号需成对或转义')
        elif not 栈 and 字 == 分隔符:
            结果.append(文本[起始:i])
            起始 = i + 1
        i += 1
    if 栈:
        raise ValueError('括号未闭合')
    结果.append(文本[起始:])
    return [项 for 项 in 结果 if 项] if 跳空 else 结果


def 去括号(值, 左, 右, 提醒=None):
    清理值 = 值.strip()
    if 清理值.startswith(左) and 清理值.endswith(右):
        return 清理值[1:-1]
    # 保持导表工具的解析语义，但在 strip 丢弃内容之前报告换行变化。
    前缀 = 值[:len(值) - len(值.lstrip())]
    后缀 = 值[len(值.rstrip()):]
    if 提醒 is not None and '\n' in 前缀 + 后缀:
        提醒('数组/对象值首尾的换行会在导表解析前被删除')
    return 清理值


def 解析类型(类型):
    if 类型 in ('int', 'long', 'float', 'string', 'bool'):
        return (类型, None)
    if 类型.endswith('[]'):
        return ('数组', 解析类型(类型[:-2]))
    if 类型.startswith('{') and 类型.endswith('}'):
        字段, 名称 = [], set()
        for 声明 in 切分(类型[1:-1], ';', True):
            部分 = 声明.rsplit(None, 1)
            if len(部分) != 2:
                raise ValueError('对象字段必须为“类型 字段名”')
            子类型, 名 = 部分
            if 名 in 名称:
                raise ValueError('对象字段名重复：' + 名)
            名称.add(名)
            字段.append((名, 解析类型(子类型)))
        if not 字段:
            raise ValueError('对象类型未声明字段')
        return ('对象', 字段)
    匹配 = re.fullmatch(r'(int|string|long)\s*\(([^\s().]+)\.([^\s().]+)\)', 类型)
    if 匹配:
        return (匹配[1], (匹配[2], 匹配[3]))
    raise ValueError('类型不合法或基本类型未使用小写：' + 类型)


def 检查值(类型, 值, 提醒, 内部=False):
    种类, 内容 = 类型
    # string 转义先于分割生效，与当前导表工具一致。
    if 种类 == '数组':
        if not 值:
            return
        原文 = 去括号(值, '[', ']', 提醒)
        项 = 切分(原文, ',')
        if any(子项 == '' for 子项 in 项):
            提醒('数组中的空项会被跳过，不会生成占位默认值')
        for 子项 in 项:
            if 子项:
                检查值(内容, 子项, 提醒, True)
        return
    if 种类 == '对象':
        if not 值 or 值.isspace():
            return
        项 = 切分(去括号(值, '{', '}', 提醒), ';')
        if any(子项 == '' for 子项 in 项):
            提醒('对象中的空项会被跳过，可能导致字段位置移动')
        项 = [子项 for 子项 in 项 if 子项]
        if len(项) > len(内容):
            raise ValueError('对象值数量超过声明字段数，多余值会被丢弃')
        for i, (_, 子类型) in enumerate(内容):
            检查值(子类型, 项[i] if i < len(项) else '', 提醒, True)
        return
    if 种类 == 'string':
        if 值.endswith('.0'):
            try:
                float(值)
            except ValueError:
                pass
            else:
                提醒('数字形字符串以 .0 结尾，导表工具会改写内容，可能丢失前导零或精度')
        if 内部 and 值.startswith('\n'):
            提醒('数组/对象内字符串开头的一个换行会被导表工具删除')
        return
    if not 值 or 值.isspace():
        return
    if 种类 == 'bool' and 值.lower() in ('true', 'false', 'yes', 'no', 'on', 'off'):
        return
    try:
        数 = float(值)
        精确 = Decimal(值)
    except (ValueError, InvalidOperation):
        raise ValueError('数据不能转换为 ' + 种类) from None
    if not math.isfinite(数):
        raise ValueError('不允许 NaN、Infinity 或浮点溢出')
    if 种类 in ('int', 'long', 'bool'):
        if 精确 != 精确.to_integral_value():
            提醒('导表工具会先截断小数，再转换为整数或布尔值')
        if int(数) != int(精确):
            raise ValueError('先转浮点再转整数会丢失精度')


def 读取表格(文件):
    """直接只读 ZIP/XML，保留公式和存储类型，不需要第三方依赖。"""
    with ZipFile(文件) as 包:
        坏文件 = 包.testzip()
        if 坏文件:
            raise ValueError('压缩包校验失败：' + 坏文件)
        共享 = []
        if 'xl/sharedStrings.xml' in 包.namelist():
            for 项 in ET.fromstring(包.read('xl/sharedStrings.xml')).findall(NS + 'si'):
                共享.append(''.join(t.text or '' for t in 项.iter(NS + 't')))
        日期样式 = set()
        if 'xl/styles.xml' in 包.namelist():
            样式 = ET.fromstring(包.read('xl/styles.xml'))
            自定 = {int(x.get('numFmtId')): x.get('formatCode', '')
                    for x in 样式.findall('./' + NS + 'numFmts/' + NS + 'numFmt')}
            for i, x in enumerate(样式.findall('./' + NS + 'cellXfs/' + NS + 'xf')):
                编号 = int(x.get('numFmtId', '0'))
                格式 = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', '', 自定.get(编号, ''))
                if 编号 in set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59)) or re.search('[ymdhs]', 格式, re.I):
                    日期样式.add(i)
        关系 = {x.get('Id'): x.get('Target') for x in ET.fromstring(包.read('xl/_rels/workbook.xml.rels'))}
        for 表 in ET.fromstring(包.read('xl/workbook.xml')).findall('./' + NS + 'sheets/' + NS + 'sheet'):
            目标 = 关系[表.get(REL + 'id')]
            路径 = 目标.lstrip('/') if 目标.startswith('/') else posixpath.normpath('xl/' + 目标)
            根 = ET.fromstring(包.read(路径))
            单元格 = {}
            for c in 根.findall('./' + NS + 'sheetData/' + NS + 'row/' + NS + 'c'):
                地址, 类型 = c.get('r'), c.get('t', 'n')
                if not 地址:
                    raise ValueError('单元格缺少地址')
                v = c.find(NS + 'v')
                值 = v.text if v is not None and v.text is not None else ''
                if 类型 == 's':
                    值 = 共享[int(值)]
                elif 类型 == 'inlineStr':
                    值 = ''.join(t.text or '' for t in c.iter(NS + 't'))
                elif 类型 == 'b':
                    值 = 'True' if 值 == '1' else 'False'
                if c.find(NS + 'f') is not None:
                    类型 = '公式'
                elif 类型 == 'd' or int(c.get('s', '0')) in 日期样式:
                    类型 = '日期'
                匹配 = re.fullmatch(r'([A-Z]+)([1-9][0-9]*)', 地址)
                if not 匹配:
                    raise ValueError('单元格地址不合法：' + 地址)
                列 = 0
                for 字 in 匹配[1]:
                    列 = 列 * 26 + ord(字) - 64
                单元格[(int(匹配[2]), 列)] = (值, 类型)
            合并范围 = [项.get('ref', '') for 项 in 根.findall('./' + NS + 'mergeCells/' + NS + 'mergeCell')]
            yield 表.get('name', ''), 单元格, 合并范围


def 检查工作表(文件, 表, 格子, 合并):
    def 值(r, c):
        return 格子.get((r, c), ('', 'n'))[0]

    def 错(r, c, 文):
        报告('错误', 文件, 表, f'{列名(c)}{r}', 文)

    def 提醒(r, c, 文):
        报告('提醒', 文件, 表, f'{列名(c)}{r}', 文)

    for 范围 in 合并:
        报告('提醒', 文件, 表, 范围, '存在合并单元格，非左上角单元格可能被读取为空')
    最大行 = max((r for r, c in 格子), default=0)
    最大列 = max((c for r, c in 格子 if r <= 2), default=0)
    表头 = [值(1, c) for c in range(1, 最大列 + 1)]
    配置表 = all(名 in 表头 for 名 in ('name', 'value', 'type'))
    for (r, c), (_, 存储) in 格子.items():
        if r <= (1 if 配置表 else 2) and 存储 in ('公式', '日期', 'e'):
            错(r, c, '表头必须为明确的文本，不能使用公式、日期或错误值')
    列类型, 键位置 = {}, {}
    if 配置表:
        for 名 in ('name', 'value', 'type', 'sign', 'description'):
            位置 = [c for c, 标题 in enumerate(表头, 1) if 标题 == 名]
            for c in 位置[1:]:
                错(1, c, f'配置表标题与 {列名(位置[0])}1 重复：{名}')
        名列, 值列, 类型列 = [表头.index(名) + 1 for 名 in ('name', 'value', 'type')]
        if 'description' not in 表头:
            提醒(1, 1, '缺少 description，当前导表工具会将最后一列误读为描述')
        if 表头 and 表头[0] == 'sign':
            提醒(1, 1, 'sign 位于第一列时当前导表工具不执行签名筛选')
        起始 = 2
    else:
        if 值(1, 1) != 'ID' or 值(2, 1) != 'key':
            错误行 = 1 if 值(1, 1) != 'ID' else 2
            错(错误行, 1, '数据表必须 A1=ID、A2=key；配置表必须含 name/value/type')
            return
        名称位置 = {}
        for c in range(2, 最大列 + 1):
            名, 类型 = 值(1, c).strip(), 值(2, c).strip()
            if 名 and not 中文.fullmatch(名):
                错(1, c, '数据表字段名必须为中文')
            if not 类型 or 类型.lower() == 'none':
                if 类型 and 类型 != 'none':
                    错(2, c, 'none 必须小写')
                if not 类型 and any(值(r, c) for r in range(3, 最大行 + 1)):
                    提醒(2, c, '此列有数据但没有类型，不会导出')
                continue
            if not 名:
                提醒(1, c, '此列有类型但没有字段名，不会导出')
                continue
            if 名 in 名称位置:
                错(1, c, f'字段名与 {列名(名称位置[名])}1 重复：{名}')
            名称位置[名] = c
            try:
                列类型[c] = (解析类型(类型), 类型)
            except ValueError as 异常:
                错(2, c, str(异常))
        if not 列类型:
            提醒(2, 1, '没有可导出的有效字段')
        起始 = 3
    空行数, 截断 = 0, False
    for r in range(起始, 最大行 + 1):
        键列 = 名列 if 配置表 else 1
        键 = 值(r, 键列).strip()
        # 先检查存储类型，避免无缓存公式和 #DIV/0! 被当作空行或注释跳过。
        键类型异常 = 格子.get((r, 键列), ('', 'n'))[1] in ('公式', '日期', 'e', 'b')
        if 键类型异常:
            错(r, 键列, 'ID 或配置名称不能使用公式、日期、错误值或布尔值')
        空行 = not (键 or 值(r, 值列) or 值(r, 类型列)) if 配置表 else not 键
        if 空行:
            空行数 += 1
            if 空行数 >= 3:
                截断 = True
            if not 配置表 and not 键类型异常 and any(v for (行, c), (v, _) in 格子.items() if 行 == r and c > 1):
                提醒(r, 1, 'ID 为空，此行其他数据不会导出')
            continue
        if 键.startswith('#'):
            continue
        if 截断:
            错(r, 1, '前面已累计三行空行，导表工具不会读取这里的数据')
        if not 配置表 and 键.startswith('!'):
            if re.match(r'![^!]+!.+', 键):
                提醒(r, 1, '签名行在未指定签名时会被跳过')
                continue
            错(r, 1, '签名行应为 !签名!key')
        if 配置表:
            类型 = 值(r, 类型列).strip()
            if not 键 or not 类型:
                提醒(r, 名列 if not 键 else 类型列, '缺少名称或类型，此配置项不会导出')
                continue
            try:
                当前类型 = {值列: (解析类型(类型), 类型)}
            except ValueError as 异常:
                错(r, 类型列, str(异常))
                continue
            if 'sign' in 表头 and 值(r, 表头.index('sign') + 1):
                提醒(r, 表头.index('sign') + 1, '配置项带签名；不同签名运行时可能筛选掉此项')
        else:
            当前类型 = 列类型
            if 键.endswith('.0'):
                try:
                    数 = float(键)
                    if not math.isfinite(数):
                        raise ValueError('ID 溢出')
                    转换键 = str(int(数))
                    if Decimal(键) != Decimal(转换键):
                        错(r, 1, 'ID 归一化时发生精度丢失')
                    键 = 转换键
                except (ValueError, InvalidOperation):
                    pass
        if 键 in 键位置:
            错(r, 名列 if 配置表 else 1, f'导出键与第 {键位置[键]} 行重复：{键}')
        键位置[键] = r
        空行数 = 0
        for c, (类型对象, 类型文本) in 当前类型.items():
            数据, 存储 = 格子.get((r, c), ('', 'n'))
            if 存储 in ('公式', '日期', 'e'):
                错(r, c, '导出数据不应存为公式、日期或 Excel 错误值；应按声明显式写入')
                continue
            if 类型对象[0] == 'string' and 数据 and 存储 not in ('s', 'inlineStr', 'str'):
                提醒(r, c, 'string 字段实际存为数字或布尔，前导零或原始文本可能已丢失')
            try:
                if 'string' in 类型文本:
                    数据 = 数据.replace('\\n', '\n').replace('\\,', '\0').replace('\\;', '\a')
                检查值(类型对象, 数据, lambda 文, r=r, c=c: 提醒(r, c, 文))
            except (ValueError, OverflowError, RecursionError) as 异常:
                错(r, c, str(异常))


def 主函数():
    参数器 = argparse.ArgumentParser(description=__doc__)
    参数器.add_argument('--目录', type=Path, default=Path(__file__).resolve().parent)
    参数器.add_argument('--json', action='store_true', help='仅输出 JSON 摘要')
    参数 = 参数器.parse_args()
    根目录 = 参数.目录.resolve()
    汇总, 标记位置 = [], {}
    for 分组, 目录 in (('数据配置', 根目录), ('文案配置', 根目录 / '文案表格')):
        if not 目录.is_dir():
            报告('错误', 目录, '', '', '检查目录不存在')
            汇总.append(dict(分类=分组, 文件数=0, 导出表数=0))
            continue
        文件列表 = sorted(p for p in 目录.iterdir() if p.is_file() and p.suffix.lower() == '.xlsx' and not p.name.startswith(('~$', '.~')))
        表数 = 0
        if not 文件列表:
            报告('提醒', 目录, '', '', '目录内没有有效 xlsx 文件')
        for 文件 in 文件列表:
            try:
                有效表 = 0
                for 表, 格子, 合并 in 读取表格(文件):
                    if not 表.startswith('|'):
                        报告('提醒', 文件, 表, '', '工作表没有 | 标记，不参与导出检查')
                        continue
                    有效表 += 1
                    表数 += 1
                    标记 = 表[1:]
                    if not 中文.fullmatch(标记):
                        报告('错误', 文件, 表, '', '导出表名必须为 | 加纯中文名称')
                    # 全局唯一是技能约束，即使数据和文案最终输出目录不同也检查。
                    if 标记 in 标记位置:
                        报告('错误', 文件, 表, '', '导出标记重复，首次来源：' + 标记位置[标记])
                    标记位置[标记] = f'{文件} / {表}'
                    检查工作表(文件, 表, 格子, 合并)
                if not 有效表:
                    报告('提醒', 文件, '', '', '此文件没有任何可导出工作表')
            except (OSError, ValueError, KeyError, IndexError, ET.ParseError, RecursionError) as 异常:
                报告('错误', 文件, '', '', '读取失败：' + str(异常))
            except Exception as 异常:
                报告('错误', 文件, '', '', f'检查异常 {type(异常).__name__}：{异常}')
        汇总.append(dict(分类=分组, 文件数=len(文件列表), 导出表数=表数))
    if sum(x['文件数'] for x in 汇总) == 0:
        报告('错误', 根目录, '', '', '未找到可检查表格，不能判定通过')
    错误数 = sum(x['级别'] == '错误' for x in 问题)
    结果 = dict(通过=错误数 == 0, 错误数=错误数, 提醒数=len(问题) - 错误数, 汇总=汇总, 问题=问题,
                验证范围='只读静态检查；未运行真实导出，未验证引用目标存在或 Excel 界面')
    if 参数.json:
        print(json.dumps(结果, ensure_ascii=False))
    else:
        print('检查目录：' + str(根目录))
        for 项 in 汇总:
            print(f"{项['分类']}：{项['文件数']} 个文件，{项['导出表数']} 张导出表")
        for 项 in 问题:
            print(f"[{项['级别']}] {项['文件']} / {项['工作表']} / {项['单元格']}：{项['原因']}")
        print(f"{'通过' if 结果['通过'] else '未通过'}：{错误数} 个错误，{结果['提醒数']} 个提醒")
        print(结果['验证范围'])
    return 1 if 错误数 else 0


if __name__ == '__main__':
    sys.exit(主函数())
