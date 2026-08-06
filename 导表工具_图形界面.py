#!/usr/bin/env python3
"""导表工具图形界面（PyQt6 版）

选择来源目录与导出目录，批量将 xlsx 导出为 json。
- 深色现代主题（QSS 样式）
- 导出前预扫描同名表标记冲突（不同 xlsx 生成同名 json 相互覆盖）
- 后台线程导出，界面不卡顿
- 逐文件进度条 + 导出日志
- 保留 --命令行 自检模式，供自动化验证

注意：因 PyQt6 对非 ASCII 方法名的信号连接存在段错误 bug
（clicked.connect(self.中文方法) 会崩溃），信号槽方法与线程方法统一使用英文命名，
其余变量、属性、核心逻辑一律中文。
"""
import multiprocessing
import os
import sys
import traceback

import sxl
import 导表工具集
from 导表核心 import 导出器
from 导表入口 import 导出上下文

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
    QProgressBar,
    QPlainTextEdit,
    QFrame,
)


常量_自检参数 = "--命令行"

# 是否为 PyInstaller 打包后的可执行程序
是打包程序 = getattr(sys, "frozen", False)
if 是打包程序:
    程序目录 = os.path.dirname(sys.executable)
else:
    程序目录 = os.path.dirname(os.path.abspath(__file__))

# 路径基准：当前 py 文件或可执行文件的上一级目录
路径基准目录 = os.path.dirname(程序目录)

# 固定配置路径：来源目录与导出目录以路径基准目录（上一级目录）为基础拼接，
# 后续无论是直接执行 py 脚本亦或是打包成可执行文件，均使用该写定的相对路径。
固定来源路径 = os.path.normpath(os.path.join(路径基准目录, "5配置文件"))
固定导出路径 = os.path.normpath(os.path.join(路径基准目录, "BouncyPinball", "数据配置"))


# 将界面填写的路径解析为绝对路径：绝对路径原样返回，相对路径以脚本所在目录的上一级目录为基准拼接。
def 解析路径(输入: str) -> str:
    输入 = (输入 or "").strip()
    if not 输入:
        return 输入
    if os.path.isabs(输入):
        return os.path.normpath(输入)
    return os.path.normpath(os.path.join(路径基准目录, 输入))


# 将项目自带的 sxl 依赖目录加入搜索路径，供系统未安装 sxl 时使用。
def 引导自带依赖() -> None:
    脚本目录 = os.path.dirname(os.path.abspath(__file__))
    自带依赖目录 = os.path.join(脚本目录, "sample", "tools", "py37")
    if os.path.isdir(自带依赖目录):
        sys.path.insert(0, 自带依赖目录)


引导自带依赖()


# ---------------------------------------------------------------- 核心逻辑

# 预扫各 xlsx 的 sheet 导出标记，检测不同文件含同名标记导致生成同名 json 相互覆盖的问题。
def 扫描根节点冲突(xlsx路径列表: list) -> dict:
    标记到来源 = {}
    上下文 = 导出上下文()
    for 路径 in xlsx路径列表:
        try:
            工作簿 = sxl.Workbook(路径)
        except Exception:
            continue
        for 名称 in 工作簿.sheets:
            if isinstance(名称, str):
                标记 = 导表工具集.获取导出标记(名称)
                if 标记:
                    表 = 工作簿.sheets[名称]
                    导出实例 = 导出器(上下文)
                    是配置表 = 导出实例.获取配置表标题信息(表) is not None
                    # 数据表与配置表均直接以标记作为文件名
                    文件名 = 标记
                    标记到来源.setdefault(文件名 + ".json", []).append(路径)
    return {文件名: 来源 for 文件名, 来源 in 标记到来源.items() if len(来源) > 1}


# 按导表入口默认参数构建单个文件的导出上下文。
def 构建导出上下文(xlsx路径: str, 导出目录: str) -> 导出上下文:
    上下文 = 导出上下文()
    上下文.路径 = xlsx路径
    上下文.文件夹 = 导出目录
    上下文.格式 = "json"
    上下文.签名 = None
    上下文.扩展名 = None
    上下文.对象分隔符 = ";"
    上下文.代码生成器 = None
    上下文.多进程数量 = None
    return 上下文


# 扫描来源目录顶层所有 xlsx，逐个导出 json，返回处理结果汇总。
def 导出全部(来源目录: str, 导出目录: str, 进度回调, 明细回调=None) -> dict:
    os.makedirs(导出目录, exist_ok=True)
    xlsx列表 = sorted(
        文件名
        for 文件名 in os.listdir(来源目录)
        if 文件名.lower().endswith(".xlsx")
        and not 文件名.startswith((".~", "~$"))  # 过滤 Office 临时锁定文件
    )
    xlsx路径列表 = [os.path.join(来源目录, 文件名) for 文件名 in xlsx列表]
    冲突 = 扫描根节点冲突(xlsx路径列表)
    成功数 = 0
    错误列表 = []
    跳过列表 = []
    部分列表 = []
    key重复列表 = []
    总数 = len(xlsx路径列表)
    for 索引, xlsx路径 in enumerate(xlsx路径列表, start=1):
        文件名 = os.path.basename(xlsx路径)
        进度回调(索引, 总数, 文件名)
        try:
            导出实例 = 导出器(构建导出上下文(xlsx路径, 导出目录))
            导出实例.导出(xlsx路径)
            key重复列表.extend(导出实例.跳过提示列表)
            # 以是否实际产出导出文件为成功判据：
            # 全部表因 key 重复被跳过（零输出）的文件单独计入跳过，不入成功也不入失败。
            已导出记录 = [记录 for 记录 in 导出实例.记录列表 if 记录.对象]
            if 已导出记录:
                成功数 += 1
                json名 = ", ".join(os.path.basename(记录.导出文件) for 记录 in 已导出记录)
                if 导出实例.错误提示列表:  # 部分表因类型/数据不合法被跳过
                    跳过详情 = "\n".join(导出实例.错误提示列表)
                    部分列表.append((文件名, 跳过详情))
                    if 明细回调:
                        明细回调(文件名, json名, 跳过详情, "部分")
                elif 明细回调:
                    明细回调(文件名, json名, None, "成功")
            elif 导出实例.错误提示列表:
                跳过详情 = "\n".join(导出实例.错误提示列表)
                错误列表.append((文件名, 跳过详情))
                if 明细回调:
                    明细回调(文件名, None, 跳过详情, "失败")
            elif 导出实例.跳过提示列表:
                跳过详情 = "\n".join(导出实例.跳过提示列表)
                跳过列表.append((文件名, 跳过详情))
                if 明细回调:
                    明细回调(文件名, None, 跳过详情, "跳过")
            else:
                错误列表.append((文件名, "未识别到可导出的表"))
                if 明细回调:
                    明细回调(文件名, None, "未识别到可导出的表", "失败")
        except Exception as 异常:  # 单个文件失败不中断整体流程
            详情 = "\n".join(导出实例.错误提示列表) if 导出实例.错误提示列表 else str(异常)
            错误列表.append((文件名, 详情))
            if 明细回调:
                明细回调(文件名, None, 详情, "失败")
    return {
        "成功": 成功数,
        "失败": len(错误列表),
        "错误列表": 错误列表,
        "跳过": len(跳过列表),
        "跳过列表": 跳过列表,
        "部分": len(部分列表),
        "部分列表": 部分列表,
        "冲突": 冲突,
        "key重复列表": key重复列表,
    }


# ---------------------------------------------------------------- 导出线程

class ExportThread(QThread):
    """后台导出线程：通过信号向界面汇报进度，避免阻塞 UI

    注：类名与信号名必须为 ASCII —— PyQt6 6.11 对非 ASCII 类名/信号名
    （如中文）的 pyqtSignal 子类存在 UnicodeEncodeError 崩溃 bug。
    """

    progress_signal = pyqtSignal(int, int, str)
    detail_signal = pyqtSignal(str, str, object, str)
    finished_signal = pyqtSignal(dict)

    def __init__(self, 来源目录: str, 导出目录: str, 父对象=None):
        super().__init__(父对象)
        self.来源目录 = 来源目录
        self.导出目录 = 导出目录

    def run(self) -> None:
        try:
            结果 = 导出全部(
                self.来源目录,
                self.导出目录,
                self.progress_signal.emit,
                self.detail_signal.emit,
            )
        except Exception:
            结果 = {
                "成功": 0,
                "失败": 1,
                "错误列表": [("整体流程", traceback.format_exc())],
                "跳过": 0,
                "跳过列表": [],
                "部分": 0,
                "部分列表": [],
                "冲突": {},
                "key重复列表": [],
            }
        self.finished_signal.emit(结果)


# ---------------------------------------------------------------- 界面样式

深色样式表 = """
QMainWindow, QDialog {
    background-color: #14161c;
}
QLabel {
    color: #e8eaf0;
    font-size: 13px;
}
QLabel#标题 {
    font-size: 20px;
    font-weight: bold;
    color: #ffffff;
}
QLabel#副标题 {
    color: #8b93a7;
    font-size: 12px;
}
QLabel#状态标签 {
    color: #a8b3c5;
}
QFrame#分隔线 {
    background-color: #2e3440;
    min-height: 1px;
    max-height: 1px;
    border: none;
}
QLineEdit {
    background-color: #1f232c;
    border: 1px solid #2e3440;
    border-radius: 8px;
    padding: 8px 12px;
    color: #e8eaf0;
    selection-background-color: #4f8cff;
    font-size: 13px;
}
QLineEdit:focus {
    border: 1px solid #4f8cff;
}
QPushButton {
    background-color: #2b303b;
    color: #e8eaf0;
    border: none;
    border-radius: 8px;
    padding: 8px 18px;
    font-size: 13px;
}
QPushButton:hover {
    background-color: #3a4150;
}
QPushButton:pressed {
    background-color: #232832;
}
QPushButton#主按钮 {
    background-color: #4f8cff;
    color: #ffffff;
    font-size: 15px;
    font-weight: bold;
    padding: 12px 24px;
}
QPushButton#主按钮:hover {
    background-color: #6ba0ff;
}
QPushButton#主按钮:pressed {
    background-color: #3d78e0;
}
QPushButton#主按钮:disabled {
    background-color: #2b3550;
    color: #7a8294;
}
QProgressBar {
    background-color: #1f232c;
    border: 2px solid #2e3440;
    border-radius: 14px;
    text-align: center;
    color: #ffffff;
    font-size: 14px;
    font-weight: bold;
    min-height: 28px;
    max-height: 28px;
}
QProgressBar::chunk {
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #3d78e0, stop:0.5 #4f8cff, stop:1 #6ba0ff);
    border-radius: 12px;
}
QPlainTextEdit {
    background-color: #0f1116;
    color: #a8b3c5;
    border: 1px solid #2e3440;
    border-radius: 8px;
    font-family: "Consolas", "Noto Sans Mono CJK SC", monospace;
    font-size: 12px;
    padding: 6px;
}
QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #3a4150;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background: #4a5263;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    background: transparent;
    height: 10px;
    margin: 0;
}
QScrollBar::handle:horizontal {
    background: #3a4150;
    border-radius: 5px;
    min-width: 30px;
}
QScrollBar::handle:horizontal:hover {
    background: #4a5263;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}
"""


# ---------------------------------------------------------------- 主窗口

class 导表窗口(QMainWindow):
    """主窗口：来源目录、导出目录、进度条、日志与开始按钮

    注：槽函数使用英文命名，规避 PyQt6 对中文方法名信号连接的段错误。
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("导表工具 - Excel 转 JSON")
        self.resize(720, 640)
        self.setMinimumSize(640, 560)
        self._线程 = None
        self._构建界面()
        self.预扫冲突()

    def _构建界面(self) -> None:
        根控件 = QWidget()
        self.setCentralWidget(根控件)
        主布局 = QVBoxLayout(根控件)
        主布局.setContentsMargins(20, 18, 20, 18)
        主布局.setSpacing(10)

        # 标题区
        标题 = QLabel("导表工具")
        标题.setObjectName("标题")
        副标题 = QLabel("将目录下所有 Excel 表导出为 JSON 配置文件")
        副标题.setObjectName("副标题")
        主布局.addWidget(标题)
        主布局.addWidget(副标题)
        主布局.addSpacing(6)

        # 来源目录（固定路径，不可编辑）
        来源行 = QHBoxLayout()
        来源行.addWidget(QLabel("来源目录:"))
        self.来源输入 = QLineEdit()
        self.来源输入.setText(固定来源路径)
        self.来源输入.setReadOnly(True)
        来源行.addWidget(self.来源输入, 1)
        主布局.addLayout(来源行)

        # 导出目录（固定路径，不可编辑）
        导出行 = QHBoxLayout()
        导出行.addWidget(QLabel("导出目录:"))
        self.导出输入 = QLineEdit()
        self.导出输入.setText(固定导出路径)
        self.导出输入.setReadOnly(True)
        导出行.addWidget(self.导出输入, 1)
        主布局.addLayout(导出行)

        # 进度条与状态
        self.进度条 = QProgressBar()
        self.进度条.setTextVisible(True)
        self.进度条.setRange(0, 1)
        self.进度条.setValue(0)
        主布局.addWidget(self.进度条)
        self.状态标签 = QLabel("就绪")
        self.状态标签.setObjectName("状态标签")
        主布局.addWidget(self.状态标签)
        self.状态标签.setWordWrap(True)

        # 横向分隔线：状态文本与导出日志之间
        分隔线 = QFrame()
        分隔线.setObjectName("分隔线")
        分隔线.setFrameShape(QFrame.Shape.HLine)
        分隔线.setFrameShadow(QFrame.Shadow.Plain)
        主布局.addWidget(分隔线)

        # 日志
        日志标题 = QLabel("导出日志")
        日志标题.setObjectName("副标题")
        主布局.addWidget(日志标题)
        self.日志区 = QPlainTextEdit()
        self.日志区.setReadOnly(True)
        主布局.addWidget(self.日志区, 1)  # 占满剩余空间

        # 开始按钮
        self.按钮开始 = QPushButton("开始导出")
        self.按钮开始.setObjectName("主按钮")
        self.按钮开始.clicked.connect(self.start_export)
        主布局.addWidget(self.按钮开始)

    # ------------------------------------------------------------ 交互

    def 预扫冲突(self) -> None:
        """启动时预扫来源目录 xlsx，检测不同文件同名标记生成的 json 相互覆盖"""
        来源目录 = 解析路径(self.来源输入.text())
        if not os.path.isdir(来源目录):
            return
        xlsx列表 = sorted(
            文件名
            for 文件名 in os.listdir(来源目录)
            if 文件名.lower().endswith(".xlsx")
            and not 文件名.startswith((".~", "~$"))  # 过滤 Office 临时锁定文件
        )
        路径列表 = [os.path.join(来源目录, 文件名) for 文件名 in xlsx列表]
        冲突 = 扫描根节点冲突(路径列表)
        if 冲突:
            说明 = self.format_conflict(冲突)
            self.状态标签.setText("⚠ 检测到同名标记冲突，请检查")
            self.append_log("警告：以下 json 文件由多个 excel 生成，后导出的会覆盖先前的：\n" + 说明)
            QMessageBox.warning(self, "警告：存在重复生成的文件", 说明)
        else:
            self.状态标签.setText("就绪")

    def validate_input(self):
        来源目录 = 解析路径(self.来源输入.text())
        导出目录 = 解析路径(self.导出输入.text())
        if not 来源目录 or not os.path.isdir(来源目录):
            QMessageBox.critical(self, "错误", "来源目录不存在")
            return None
        if not 导出目录:
            QMessageBox.critical(self, "错误", "请填写导出目录")
            return None
        return 来源目录, 导出目录

    def start_export(self) -> None:
        if self._线程 is not None and self._线程.isRunning():
            return
        目录组 = self.validate_input()
        if not 目录组:
            return
        来源目录, 导出目录 = 目录组

        # 重置界面状态
        self.按钮开始.setEnabled(False)
        self.按钮开始.setText("导出中...")
        self.进度条.setRange(0, 1)
        self.进度条.setValue(0)
        self.状态标签.setText("正在扫描...")
        self.日志区.clear()

        self._线程 = ExportThread(来源目录, 导出目录, self)
        self._线程.progress_signal.connect(self.update_progress)
        self._线程.detail_signal.connect(self.update_file_detail)
        self._线程.finished_signal.connect(self.show_finished)
        self._线程.start()

    def update_progress(self, 索引: int, 总数: int, 文件名: str) -> None:
        self.进度条.setRange(0, 总数)
        self.进度条.setValue(索引)
        self.状态标签.setText(f"正在导出 {文件名} ({索引}/{总数})")

    def update_file_detail(self, 文件名: str, json名, 错误详情, 状态: str) -> None:
        """每个文件导出完成后更新日志"""
        if 状态 == "成功":
            self.append_log(f"✔ {文件名} → {json名}")
        elif 状态 == "部分":
            self.append_log(f"⚠ {文件名} → {json名}（部分表跳过）\n{错误详情}")
        elif 状态 == "跳过":
            self.append_log(f"⏭ {文件名} 全部跳过：\n{错误详情}")
        else:
            self.append_log(f"✘ {文件名} 失败：\n{错误详情}")

    @staticmethod
    def format_key_duplicate(key重复列表: list) -> str:
        """key 重复提示段落，未发生时返回空字符串"""
        if not key重复列表:
            return ""
        return "\n\n⚠ 以下表因 key 重复已跳过未导出：\n  - " + "\n  - ".join(key重复列表)

    def show_finished(self, 结果: dict) -> None:
        self.按钮开始.setEnabled(True)
        self.按钮开始.setText("开始导出")
        成功数 = 结果["成功"]
        失败数 = 结果["失败"]
        跳过数 = 结果.get("跳过", 0)
        部分数 = 结果.get("部分", 0)
        冲突 = 结果["冲突"]
        key重复段落 = self.format_key_duplicate(结果.get("key重复列表", []))
        部分段落 = self.format_partial(结果.get("部分列表", []))
        提示文本 = f"导出完成：成功 {成功数} 个，失败 {失败数} 个，跳过 {跳过数} 个，部分跳过 {部分数} 个"
        self.状态标签.setText(提示文本)

        跳过段落 = ""
        if 结果.get("跳过列表"):
            跳过段落 = "\n\n⚠ 以下文件因全部表 key 重复已跳过未导出：\n  - " + "\n  - ".join(
                f"{文件名}\n    {详情}" for 文件名, 详情 in 结果["跳过列表"]
            )
        if 冲突:
            说明 = self.format_conflict(冲突) + key重复段落 + 部分段落 + 跳过段落
            self.append_log("═══ 导出汇总 ═══\n" + 说明)
            QMessageBox.warning(self, "警告：存在重复生成的文件", 说明)
            self.状态标签.setText(提示文本 + "（存在重复覆盖警告）")
        elif 失败数 > 0:
            说明 = self.format_error(结果["错误列表"]) + key重复段落 + 部分段落 + 跳过段落
            self.append_log("═══ 导出汇总 ═══\n" + 说明)
            QMessageBox.warning(self, "部分文件导出失败", 说明)
        elif 部分数 > 0:
            说明 = 提示文本 + 部分段落 + key重复段落 + 跳过段落
            self.append_log("═══ 导出汇总 ═══\n" + 说明)
            QMessageBox.warning(self, "部分表未导出", 说明)
        elif 跳过数 > 0:
            说明 = 提示文本 + 跳过段落
            self.append_log("═══ 导出汇总 ═══\n" + 说明)
            QMessageBox.information(self, "完成", 说明)
        else:
            self.append_log("═══ 导出汇总 ═══\n" + 提示文本)
            QMessageBox.information(self, "完成", 提示文本)

    @staticmethod
    def format_partial(部分列表: list) -> str:
        """部分表因类型/数据不合法被跳过的提示段落，未发生时返回空字符串"""
        if not 部分列表:
            return ""
        return "\n\n⚠ 以下文件的部分表因类型/数据不合法已跳过：\n  - " + "\n  - ".join(
            f"{文件名}\n    {详情}" for 文件名, 详情 in 部分列表
        )

    def append_log(self, 文本: str) -> None:
        self.日志区.appendPlainText(文本)

    @staticmethod
    def format_error(错误列表: list) -> str:
        return "\n\n".join(f"{文件名}:\n{详情}" for 文件名, 详情 in 错误列表)

    @staticmethod
    def format_conflict(冲突: dict) -> str:
        行列表 = ["以下 json 文件由多个 excel 生成（表标记同名），后导出的会覆盖先前的："]
        for 文件名, 来源 in 冲突.items():
            行列表.append(f"\n{文件名}")
            for 源文件 in 来源:
                行列表.append(f"  ← {源文件}")
        return "\n".join(行列表)

    def closeEvent(self, 事件) -> None:
        """关闭窗口时确保后台线程结束"""
        if self._线程 is not None and self._线程.isRunning():
            self._线程.wait(3000)
        事件.accept()


# ---------------------------------------------------------------- 命令行模式

# 无窗口自检模式：打印进度与结果，供自动化验证，来源与导出目录写死为固定路径。
def 命令行导出() -> int:
    结果 = 导出全部(固定来源路径, 固定导出路径, lambda 索引, 总数, 文件名: print(f"[{索引}/{总数}] {文件名}"))
    print(f"完成：成功 {结果['成功']} 个，失败 {结果['失败']} 个，跳过 {结果.get('跳过', 0)} 个，部分跳过 {结果.get('部分', 0)} 个")
    if 结果["key重复列表"]:
        print("⚠ 以下表因 key 重复已跳过未导出：")
        for 提示 in 结果["key重复列表"]:
            print(f"  - {提示}")
    if 结果.get("部分列表"):
        print("⚠ 以下文件的部分表因类型/数据不合法已跳过：")
        for 文件名, 详情 in 结果["部分列表"]:
            print(f"  - {文件名}\n    {详情}")
    if 结果.get("跳过列表"):
        print("⚠ 以下文件因全部表 key 重复已跳过未导出：")
        for 文件名, 详情 in 结果["跳过列表"]:
            print(f"  - {文件名}\n    {详情}")
    if 结果["冲突"]:
        print("警告：以下 json 文件由多个 excel 生成，后导出的会覆盖先前的：")
        for 文件名, 来源 in 结果["冲突"].items():
            print(f"  {文件名} <- {来源}")
    if 结果["错误列表"]:
        for 文件名, 详情 in 结果["错误列表"]:
            print(f"失败 {文件名}:\n{详情}")
        return 1
    return 0


def 主函数() -> None:
    multiprocessing.freeze_support()  # PyInstaller 打包后 multiprocessing 需要
    if 常量_自检参数 in sys.argv:
        sys.exit(命令行导出())

    # 高分屏适配
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    应用 = QApplication(sys.argv)
    应用.setStyleSheet(深色样式表)
    # 中文字体优先
    字体 = QFont()
    字体.setFamilies(["Microsoft YaHei UI", "PingFang SC", "Noto Sans Mono CJK SC", "WenQuanYi Micro Hei", "sans-serif"])
    应用.setFont(字体)
    窗口 = 导表窗口()
    窗口.show()
    sys.exit(应用.exec())


if __name__ == "__main__":
    主函数()