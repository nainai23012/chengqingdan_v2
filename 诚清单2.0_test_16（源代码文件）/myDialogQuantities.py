# -*- coding: utf-8 -*-
import sys
import os
import ast
import re
import xlrd
from win32com.client import Dispatch  # 打开excel程序
import pickle
from pprint import pprint
import copy
from PyQt5.QtWidgets import qApp, QApplication, QMainWindow, QUndoStack, QUndoCommand, \
    QMessageBox, QSpinBox, QLabel, QTableView, QCheckBox, QAbstractItemView, QHeaderView, \
    QColorDialog, QDialog, QFileDialog, QPlainTextEdit, QInputDialog, QLineEdit, QComboBox, \
    QListWidget, QListWidgetItem

from PyQt5.QtCore import pyqtSlot, pyqtSignal, Qt, QItemSelectionModel, QStringListModel, \
    QCoreApplication, QDir
from PyQt5.QtGui import QFont, QColor, QPalette, QStandardItemModel, QStandardItem
##from PyQt5.QtSql import
##from PyQt5.QtMultimedia import
##from PyQt5.QtMultimediaWidgets import
# 载入对话框
from myDelegates import QmyFloatSpinDelegate, QmyComboBoxDelegate  # 定义代理模块
from ui_QWDialogQuantities import Ui_Dialog
from myDialogFExcel import QmyDialogFExcel
from myDialogImportGlodon import QmyDialogImportGlodon
# 载入 子对话框
from myDialogImportReinforcementQuantities import QmyDialogImportReinforcementQuantities


class QmyDialogQuantities(QDialog):
    changeActionEnable = pyqtSignal(bool)  # 用于设置主窗口的Action的Enabled
    quantitiesTableLists = pyqtSignal(dict)  # 多张表与历史操作 发送到主界面
    # quantitiesOK = pyqtSignal(bool)  # 是否点击了OK 退出的 发送到主界面
   # changeCellText = pyqtSignal(int, int, str)   #用于设置主窗口的单元格的内容

    def __init__(self, parent=None):
        super().__init__(parent)  # 调用父类构造函数，创建窗体
        self.ui = Ui_Dialog()  # 创建UI对象
        self.ui.setupUi(self)  # 构造UI界面
        # self.resize(1800, 900)
        # self.setWindowState(Qt.WindowMaximized)  # 窗口最大化显示 全屏
        self.edition = '工程量总表：包含“广联达工程量表（土建、钢筋）”与“分类表”的全部工程量数据'  # 统一的版本号
        # 设置窗口名称
        self.setWindowTitle(self.edition)
        # 窗口最大化显示 全屏
        # self.setWindowState(Qt.WindowFullScreen)
        # 窗口颜色
        # col = QColor(245,255,255)
        # self.setStyleSheet('QWidget{background-color:%s}' % col.name())
        # 隐藏标题栏
        # self.setWindowFlags(Qt.FramelessWindowHint)  # 隐藏标题
        # self.setFixedSize(self.width(),self.height())项目概况

        # 选项卡控件设置
        self.ui.tabWidget_1.tabBarClicked.connect(self.tabClicked_1)  # 选项卡被点击时触发 清单表 分析表
        self.ui.tabWidget_1.currentChanged.connect(self.tabWidget_1_currentChanged)
        self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 工程量页面

        # 公共变量
        self.quantitiesDict = {"历史操作": [], "工程量总表": [], "分类表": [], "钢筋工程量文件表": [], "钢筋工程量表": [],
                                "绘图工程量文件表": [], "绘图工程量表": []}
        # self.dictIron3 = {...}  # 放入钢筋表的 初始化内
        self.matchingdict = {}  # 匹配表字典 绘图工程量表 被点击时 初始化
        self.tabindex = '工程量总表'  # 选项卡 中文名
        self.templist = []  # 复制粘贴行用的临时列表变量
        self.spinFontSize = 10  # 字体大小
        self.ui.spinBoxSize.valueChanged.connect(self.spinBoxSize_valueChanged)
        # self.decimalPlaces = 3  # 小数位数
        self.beforestrundo = [0, 0]  # 改之前的原字符  用于记录历史操作的
        self.afterstrundo = ""  # 改字符之后的新字符  用于记录历史操作的
        self.filename = None  # 路径
        self.__NoCalFlags = (Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)  # 不计标志 复选框
        self.__NoCalTitle = ""  # 不计标志 复选框

        # 控件初始化
        self.__init_tableView_11()  # 工程量总表 初始化
        self.__init_tableView_12()  # 明细 初始化
        self.__init_tableView_2()  # 分类 初始化

        self.__init_tableView_31()  # 钢筋文件表 初始化
        self.__init_tableView_32()  # 钢筋 初始化
        self.__init_tableView_33()  # 钢筋明细表  初始化

        self.__init_tableView_41()  # 绘图工程量文件表 初始化
        self.__init_tableView_42()  # 绘图工程量表 初始化
        self.__init_tableView_43()  # 绘图工程量表 明细表 初始化

        self.__init_listView_undo()  # 历史操作 初始化
        self.__init_comboBoxIron()  # 钢筋发送清单表的 下拉复选框
        # 信号与函数连接
        self.connectAll()

        # 分割器设置
        self.ui.splitter_1.setStretchFactor(0, 10)
        self.ui.splitter_1.setStretchFactor(1, 90)

        self.ui.splitter_2.setStretchFactor(0, 70)
        self.ui.splitter_2.setStretchFactor(1, 30)

        self.ui.splitter_3.setStretchFactor(0, 70)
        self.ui.splitter_3.setStretchFactor(1, 30)

        self.ui.splitter_4.setStretchFactor(0, 70)
        self.ui.splitter_4.setStretchFactor(1, 30)
        # 控件 快捷键设置
        self.ui.QuanBtnCopy.setShortcut('ctrl+c')
        self.ui.QuanBtnPaste.setShortcut('ctrl+v')
        self.ui.QuanBtnFind.setShortcut('ctrl+f')
        self.ui.QuanBtnInsertRow.setShortcut('insert')
        self.ui.QuanBtnDelRow.setShortcut('delete')

        # 初始不可用控件

        self.ui.tableView_12.setEnabled(False)  # 明细表不可用
        self.ui.QuanBtnPaste.setEnabled(False)  # 粘贴行 复制行时触发
        self.ui.QuanBtnDelRow.setEnabled(False)  # 删除行 有选中行时触发
        self.ui.QuanBtnIronImport.setEnabled(False)  # 不可用
        self.ui.QuanBtnGlodonImport.setEnabled(False)  # 不可用
        self.ui.QuanBtnIronOK.setEnabled(False)  # 钢筋量表 发送至清单总表
        self.ui.QuanBtnGlodonOK.setEnabled(False)  # 土建量表 发送至清单总表
        self.ui.QuanBtnImportClassification.setEnabled(False)  # 分类表 导入
        # self.ui.btnOK.setEnabled(False)  # 删除行 有选中行时触发
        # self.ui.btnExit.setEnabled(False)  # 删除行 有选中行时触发

    def __del__(self):  # 析构函数
        pass
        print("QmyDialogHeaders 对象被删除了")

    # TODO  ==============event处理函数 事件==========================
    # 窗体显示时 对话框显示事件
    def showEvent(self, event):
        self.changeActionEnable.emit(False)
        super().showEvent(event)

    # 窗体关闭时询问
    def closeEvent(self, event):
        dlgTitle = "警告！"
        strInfo = "确定要退出吗？ \n   Yes 保存、发送、退出。\n   No 不保存、退出\n   Cancel 取消！"
        defaultBtn = QMessageBox.No  # 缺省按钮
        result = QMessageBox.question(self, dlgTitle, strInfo,
                                      QMessageBox.Yes | QMessageBox.Cancel | QMessageBox.No, defaultBtn)
        if result == QMessageBox.Yes:  # 按下了 yes
            self.on_QuanBtnRefresh_clicked()  # 刷新主程序
            self.changeActionEnable.emit(True)
            event.accept()  # 窗口可关闭
            super().closeEvent(event)
        elif result == QMessageBox.Cancel:  # 按下了取消
            event.ignore()  # 窗口不能被关闭
        else:  # 按下了 NO
            self.changeActionEnable.emit(True)
            super().closeEvent(event)

    def keyPressEvent(self, event):  # 重新实现了keyPressEvent()事件处理器。
        # 按住键盘事件
        # 这个事件是PyQt自带的自动运行的，当我修改后，其内容也会自动调用
        if event.key() == Qt.Key_Escape:  # 当我们按住键盘是esc按键时
            self.close()  # 关闭程序


    # TODO  ==============初始化功能函数========================
    # 非常重要 传入一个新建的数据模型 全部赋值为空白 否则无法获取text（）
    def initItemModelBlank(self, itemModel):
        rows = itemModel.rowCount()
        cols = itemModel.columnCount()
        for x in range(rows):
            for y in range(cols):
                item = QStandardItem("")
                itemModel.setItem(x, y, item)  # 初始化表 itemModel 每项目为空

    # 工程量总表
    def __init_tableView_11(self):
        self.itemModel11 = QStandardItemModel(1, 9, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel11)  # 初始化 数据模型为 空值
        headerList = ['序号', '来源', '引用编码\n项目名称', '项目名称', '项目特征描述', '计量\n单位', '工程量', '∑工程量明细表\n隐藏', '备注']
        self.itemModel11.setHorizontalHeaderLabels(headerList)  # 设置表头标题
        self.selectionModel11 = QItemSelectionModel(self.itemModel11)  # itemModel 选择模型

        self.ui.tableView_11.setModel(self.itemModel11)  # 设置数据模型
        self.ui.tableView_11.setSelectionModel(self.selectionModel11)  # 设置选择模型
        self.ui.tableView_11.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))  # 表头字符对齐 换行
        self.ui.tableView_11.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸

        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_11.setSelectionMode(oneOrMore)  # 可多选
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_11.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_11.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_11.setAlternatingRowColors(True)  # 交替行颜色
        self.ui.tableView_11.setSelectionBehavior(QAbstractItemView.SelectRows)  # 单击某个项目时,将选择整个列
        self.ui.tableView_11.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"  # #A9A9A9
            "QTableView::item{selection-background-color:#3399FF}")  # 表格颜色样式与 选中行 反色样式 lightBlue darkcyan #3399FF

        self.ui.tableView_11.setColumnWidth(0, 40)
        self.ui.tableView_11.setColumnWidth(1, 160)
        self.ui.tableView_11.setColumnWidth(2, 280)
        # self.ui.tableView_11.setColumnWidth(3, 160)
        self.ui.tableView_11.horizontalHeader().hideSection(3)
        self.ui.tableView_11.setColumnWidth(4, 120)
        self.ui.tableView_11.setColumnWidth(5, 40)
        self.ui.tableView_11.setColumnWidth(6, 100)
        # self.ui.tableView_11.setColumnWidth(7, 100)  # ∑工程量明细表
        self.ui.tableView_11.horizontalHeader().hideSection(7)
        self.ui.tableView_11.setColumnWidth(8, 300)
        # self.ui.tableView_11.horizontalHeader().hideSection(8)

        # 计量单位 下拉
        qualities = ["", "m", "m2", "m3", "T", "kg", "个", "项", "座", "樘"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_11.setItemDelegateForColumn(5, self.UnitOfMeasurement)  # 计量单位

    # 工程量总表 明细表
    def __init_tableView_12(self):
        self.itemModel12 = QStandardItemModel(6, 8, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel12)  # 初始化 数据模型为 空值
        # 计算结果temp 用于暂存去除【注释】的结果
        headerList = ['楼层号\n部位', '计算表达式\n中文括号【注释】\t\t\t\t尖括号<楼层>引用',
                      '计算结果real\n隐藏', '计算结果', '小计\n楼 层', '不计\n标志', '公式\n错误', '备注']
        self.itemModel12.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel12 = QItemSelectionModel(self.itemModel12)  # itemModel 选择模型

        self.ui.tableView_12.setModel(self.itemModel12)  # 设置数据模型
        self.ui.tableView_12.setSelectionModel(self.selectionModel12)  # 设置选择模型

        # tb=QTableView  # 表头换行
        self.ui.tableView_12.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_12.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_12.setSelectionMode(oneOrMore)  # 可多选
        #
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_12.setSelectionBehavior(itemOrRow)  # 单元格选择

        # self.ui.tableView_12.verticalHeader().setDefaultSectionSize(28)  # 缺省行高
        self.ui.tableView_12.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_12.setAlternatingRowColors(True)  # 交替行颜色

        # self.ui.tableView_12.setStyleSheet("QTableView{border:1px solid #014F84}")

        # 设置表头边框样式
        self.ui.tableView_12.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"
            "QTableView::item{selection-background-color:#3399FF}")
        self.ui.tableView_12.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        # self.ui.tableView_12.setStyleSheet("selection-background-color:lightBlue")   # 单单元格选中变色  且光标离开保留颜色

        self.ui.tableView_12.setColumnWidth(0, 160)
        self.ui.tableView_12.setColumnWidth(1, 500)
        # self.ui.tableView_12.setColumnWidth(2, 100)  # 结算结果的real  后期 隐藏
        self.ui.tableView_12.horizontalHeader().hideSection(2)
        self.ui.tableView_12.setColumnWidth(3, 100)
        self.ui.tableView_12.setColumnWidth(4, 100)
        self.ui.tableView_12.setColumnWidth(5, 40)
        self.ui.tableView_12.setColumnWidth(6, 40)
        self.ui.tableView_12.setColumnWidth(7, 300)
        for i in range(self.itemModel12.rowCount()):
            item = QStandardItem(self.__NoCalTitle)  # 最后一列
            item.setFlags(self.__NoCalFlags)
            item.setCheckable(True)  # 非锁定
            item.setCheckState(Qt.Unchecked)  # 非勾选
            self.itemModel12.setItem(i, 5, item)  # 设置最后一列的item

    # 分类表
    def __init_tableView_2(self):
        self.itemModel2 = QStandardItemModel(3, 9, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel2)  # 初始化 数据模型为 空值

        headerList = ['序号', '楼层/部位', '清单名称', '计量\n单位', '计算表达式\n中文括号【注释】',
                      '工程量', '不计\n标志', '公式\n错误', '备注']
        self.itemModel2.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel2 = QItemSelectionModel(self.itemModel2)  # itemModel 选择模型

        self.ui.tableView_2.setModel(self.itemModel2)  # 设置数据模型
        self.ui.tableView_2.setSelectionModel(self.selectionModel2)  # 设置选择模型

        # tb=QTableView  # 表头对齐方式换行
        self.ui.tableView_2.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_2.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_2.setSelectionMode(oneOrMore)  # 可多选

        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_2.setSelectionBehavior(itemOrRow)  # 单元格选择

        # self.ui.num_tableView2.verticalHeader().setDefaultSectionSize(28)  # 缺省行高
        self.ui.tableView_2.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_2.setAlternatingRowColors(True)  # 交替行颜色

        # 设置表头边框样式
        # self.ui.tableView_2.setStyleSheet(
        #     "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9}")
        # self.ui.tableView_2.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        self.ui.tableView_2.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        self.ui.tableView_2.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9}""QTableView::item{selection-background-color:#3399FF}")
        self.ui.tableView_2.setColumnWidth(0, 50)
        self.ui.tableView_2.setColumnWidth(1, 160)
        self.ui.tableView_2.setColumnWidth(2, 220)
        self.ui.tableView_2.setColumnWidth(3, 40)
        self.ui.tableView_2.setColumnWidth(4, 400)
        self.ui.tableView_2.setColumnWidth(5, 100)
        self.ui.tableView_2.setColumnWidth(6, 40)
        self.ui.tableView_2.setColumnWidth(7, 40)
        self.ui.tableView_2.setColumnWidth(8, 300)
        # self.Quantity = QmyFloatSpinDelegate(0, 1000000, 3, self)  # 用于工程量 最小值 最大值 精度
        # self.ui.tableView_2.setItemDelegateForColumn(6, self.Quantity)
        #
        # self.price = QmyFloatSpinDelegate(0, 20000, 2, self)  # 用于价格 最小值 最大值 精度
        # self.ui.tableView_2.setItemDelegateForColumn(7, self.price)
        qualities = ["", "m", "m2", "m3", "T", "kg", "个", "项"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_2.setItemDelegateForColumn(3, self.UnitOfMeasurement)  # 计量单位
        for i in range(self.itemModel2.rowCount()):
            item = QStandardItem(self.__NoCalTitle)  # 最后一列
            item.setFlags(self.__NoCalFlags)
            item.setCheckable(True)  # 非锁定
            item.setCheckState(Qt.Unchecked)  # 非勾选
            self.itemModel2.setItem(i, 6, item)  # 设置最后一列的item
        self.itemModel2.setItem(0, 2, QStandardItem("请勿使用《 》,【 】做为清单名称,\n否则无法被引用，此条可删！"))

    # 钢筋工程量文件表
    def __init_tableView_31(self):
        # 钢筋发送至清单表时  再次初始化  必须一样修改
        self.dictIron3 = {'级别总重': {},  # 钢筋种类的总重量 如 一级钢A 500t 二级钢B 600t
                          '接头个数': {},  # 除绑扎以外的接头总数 如 电渣压力焊 100个
                          '接头直径个数': {},  # 一级钢A6 500t
                          '构件总重': {},
                          '构件楼层总重': {},
                          '楼层构件总重': {},
                          '级别直径总重': {},
                          '构件纵筋总重': {},
                          '江苏2014_钢筋重量': {},
                          '江苏2014_接头个数': {}
                          }
        # self.dictIron3_dcopy = copy.deepcopy(self.dictIron3)
        self.itemModel31 = QStandardItemModel(0, 4, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel31)  # 初始化 数据模型为 空值
        headerList = ['序号', '文件名', '文件数据量', '文件数据\n隐藏']
        self.itemModel31.setHorizontalHeaderLabels(headerList)  # 设置表头标题
        self.selectionModel31 = QItemSelectionModel(self.itemModel31)  # itemModel 选择模型

        self.ui.tableView_31.setModel(self.itemModel31)  # 设置数据模型
        self.ui.tableView_31.setSelectionModel(self.selectionModel31)  # 设置选择模型
        self.ui.tableView_31.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))  # 表头字符对齐 换行
        # self.ui.tableView_31.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸

        # oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        # self.ui.tableView_31.setSelectionMode(oneOrMore)  # 可多选
        # itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        # self.ui.tableView_31.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_31.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        # self.ui.tableView_31.verticalHeader().setDefaultSectionSize(58)  # 缺省行高 设置了行高自动 就失效了
        self.ui.tableView_31.setAlternatingRowColors(True)  # 交替行颜色
        # "QTableView::item{background-color:rgb(223, 255, 255);}"
        # QWidget{background-color:rgb(245, 245, 245);}''')
        self.ui.tableView_31.setSelectionBehavior(QAbstractItemView.SelectRows)  # 单击某个项目时,将选择整个列
        self.ui.tableView_31.setStyleSheet("QHeaderView::section{background:rgb(203, 255, 255)}"
                                           "QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"  # #A9A9A9
                                           "QTableView::item{selection-background-color:#3399FF}"
                                           )
        bigcolWidth = 100
        self.ui.tableView_31.setColumnWidth(0, 50)
        self.ui.tableView_31.setColumnWidth(1, 700)
        self.ui.tableView_31.setColumnWidth(2, 100)
        # self.ui.tableView_31.setColumnWidth(3, bigcolWidth)
        self.ui.tableView_31.horizontalHeader().hideSection(3)
        # 双击事件

    # 钢筋工程量表
    def __init_tableView_32(self):
        self.itemModel32 = QStandardItemModel(1, 9, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel32)  # 初始化 数据模型为 空值
        headerList = ['序号', '来源', '引用编码\n隐藏', '报表名称', '特征描述', '计量\n单位', '工程量', '∑工程量明细表\n隐藏', '备注']
        self.itemModel32.setHorizontalHeaderLabels(headerList)  # 设置表头标题
        self.selectionModel32 = QItemSelectionModel(self.itemModel32)  # itemModel 选择模型

        self.ui.tableView_32.setModel(self.itemModel32)  # 设置数据模型
        self.ui.tableView_32.setSelectionModel(self.selectionModel32)  # 设置选择模型
        self.ui.tableView_32.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))  # 表头字符对齐 换行
        self.ui.tableView_32.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸

        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_32.setSelectionMode(oneOrMore)  # 可多选
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_32.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_32.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_32.setAlternatingRowColors(True)  # 交替行颜色
        self.ui.tableView_32.setSelectionBehavior(QAbstractItemView.SelectRows)  # 单击某个项目时,将选择整个列
        self.ui.tableView_32.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"  # #A9A9A9
            "QTableView::item{selection-background-color:#3399FF}")  # 表格颜色样式与 选中行 反色样式 lightBlue darkcyan #3399FF

        self.ui.tableView_32.setColumnWidth(0, 40)
        self.ui.tableView_32.setColumnWidth(1, 100)
        # self.ui.tableView_32.setColumnWidth(2, 100)
        self.ui.tableView_32.horizontalHeader().hideSection(2)
        self.ui.tableView_32.setColumnWidth(3, 140)
        self.ui.tableView_32.setColumnWidth(4, 300)
        self.ui.tableView_32.setColumnWidth(5, 40)
        self.ui.tableView_32.setColumnWidth(6, 100)
        # self.ui.tableView_11.setColumnWidth(7, 100)  # ∑工程量明细表
        self.ui.tableView_32.horizontalHeader().hideSection(7)
        self.ui.tableView_32.setColumnWidth(8, 300)
        # self.ui.tableView_11.horizontalHeader().hideSection(8)

        # 计量单位 下拉
        qualities = ["", "m", "m2", "m3", "T", "kg", "个", "项", "座"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_32.setItemDelegateForColumn(5, self.UnitOfMeasurement)  # 计量单位

    # 钢筋工程量明细表
    def __init_tableView_33(self):
        self.itemModel33 = QStandardItemModel(6, 8, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel33)  # 初始化 数据模型为 空值
        # 计算结果temp 用于暂存去除【注释】的结果
        headerList = ['楼层号\n或分类', '计算表达式\n中文括号【注释】\t\t\t\t尖括号<楼层>引用',
                      '计算结果real\n隐藏', '计算结果', '小计\n楼 层', '不计\n标志', '公式\n错误', '备注']
        self.itemModel33.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel33 = QItemSelectionModel(self.itemModel33)  # itemModel 选择模型

        self.ui.tableView_33.setModel(self.itemModel33)  # 设置数据模型
        self.ui.tableView_33.setSelectionModel(self.selectionModel33)  # 设置选择模型

        # tb=QTableView  # 表头换行
        self.ui.tableView_33.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_33.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_33.setSelectionMode(oneOrMore)  # 可多选
        #
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_33.setSelectionBehavior(itemOrRow)  # 单元格选择

        # self.ui.tableView_12.verticalHeader().setDefaultSectionSize(28)  # 缺省行高
        self.ui.tableView_33.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_33.setAlternatingRowColors(True)  # 交替行颜色

        # self.ui.tableView_12.setStyleSheet("QTableView{border:1px solid #014F84}")

        # 设置表头边框样式
        self.ui.tableView_33.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"
            "QTableView::item{selection-background-color:#3399FF}")
        self.ui.tableView_33.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        # self.ui.tableView_12.setStyleSheet("selection-background-color:lightBlue")   # 单单元格选中变色  且光标离开保留颜色

        self.ui.tableView_33.setColumnWidth(0, 80)
        self.ui.tableView_33.setColumnWidth(1, 400)
        # self.ui.tableView_33.setColumnWidth(2, 100)  # 结算结果的real  后期 隐藏
        self.ui.tableView_33.horizontalHeader().hideSection(2)
        self.ui.tableView_33.setColumnWidth(3, 100)
        self.ui.tableView_33.setColumnWidth(4, 100)
        self.ui.tableView_33.setColumnWidth(5, 40)
        self.ui.tableView_33.setColumnWidth(6, 40)
        self.ui.tableView_33.setColumnWidth(7, 300)
        for i in range(self.itemModel33.rowCount()):
            item = QStandardItem(self.__NoCalTitle)  # 最后一列
            item.setFlags(self.__NoCalFlags)
            item.setCheckable(True)  # 非锁定
            item.setCheckState(Qt.Unchecked)  # 非勾选
            self.itemModel33.setItem(i, 5, item)  # 设置最后一列的item

    # 绘图工程量文件表
    def __init_tableView_41(self):
        self.itemModel41 = QStandardItemModel(0, 4, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel41)  # 初始化 数据模型为 空值

        headerList = ['序号', '文件名', '文件数据量', '文件数据\n隐藏']  # dict类型 {表名简称：datas}
        self.itemModel41.setHorizontalHeaderLabels(headerList)  # 设置表头标题
        self.selectionModel41 = QItemSelectionModel(self.itemModel41)  # itemModel 选择模型

        self.ui.tableView_41.setModel(self.itemModel41)  # 设置数据模型
        self.ui.tableView_41.setSelectionModel(self.selectionModel41)  # 设置选择模型
        self.ui.tableView_41.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))  # 表头字符对齐 换行
        # self.ui.tableView_41.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸

        # oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        # self.ui.tableView_41.setSelectionMode(oneOrMore)  # 可多选
        # itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        # self.ui.tableView_41.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_41.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        # self.ui.tableView_41.verticalHeader().setDefaultSectionSize(58)  # 缺省行高 设置了行高自动 就失效了
        self.ui.tableView_41.setAlternatingRowColors(True)  # 交替行颜色

        self.ui.tableView_41.setSelectionBehavior(QAbstractItemView.SelectRows)  # 单击某个项目时,将选择整个列
        self.ui.tableView_41.setStyleSheet("QHeaderView::section{background:rgb(203, 255, 255)}"
                                           "QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"  # #A9A9A9
                                           "QTableView::item{selection-background-color:#3399FF}")
        # 表格颜色样式与 选中行 反色样式 lightBlue darkcyan #3399FF
        bigcolWidth = 100
        self.ui.tableView_41.setColumnWidth(0, 50)
        self.ui.tableView_41.setColumnWidth(1, 600)
        self.ui.tableView_41.setColumnWidth(2, 180)
        # self.ui.tableView_41.setColumnWidth(3, bigcolWidth)
        self.ui.tableView_41.horizontalHeader().hideSection(3)

    # 绘图工程量表
    def __init_tableView_42(self):
        self.itemModel42 = QStandardItemModel(1, 9, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel42)  # 初始化 数据模型为 空值
        headerList = ['序号', '来源', '引用编码\n隐藏', '报表名称', '特征描述', '计量\n单位', '工程量', '∑工程量明细表\n隐藏', '备注']
        self.itemModel42.setHorizontalHeaderLabels(headerList)  # 设置表头标题
        self.selectionModel42 = QItemSelectionModel(self.itemModel42)  # itemModel 选择模型

        self.ui.tableView_42.setModel(self.itemModel42)  # 设置数据模型
        self.ui.tableView_42.setSelectionModel(self.selectionModel42)  # 设置选择模型
        self.ui.tableView_42.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))  # 表头字符对齐 换行
        self.ui.tableView_42.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸

        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_42.setSelectionMode(oneOrMore)  # 可多选
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_42.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_42.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_42.setAlternatingRowColors(True)  # 交替行颜色
        self.ui.tableView_42.setSelectionBehavior(QAbstractItemView.SelectRows)  # 单击某个项目时,将选择整个列
        self.ui.tableView_42.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"  # #A9A9A9
            "QTableView::item{selection-background-color:#3399FF}")  # 表格颜色样式与 选中行 反色样式 lightBlue darkcyan #3399FF

        self.ui.tableView_42.setColumnWidth(0, 40)
        self.ui.tableView_42.setColumnWidth(1, 100)
        # self.ui.tableView_42.setColumnWidth(2, 100)
        self.ui.tableView_42.horizontalHeader().hideSection(2)
        self.ui.tableView_42.setColumnWidth(3, 140)
        self.ui.tableView_42.setColumnWidth(4, 300)
        self.ui.tableView_42.setColumnWidth(5, 40)
        self.ui.tableView_42.setColumnWidth(6, 100)
        # self.ui.tableView_11.setColumnWidth(7, 100)  # ∑工程量明细表
        self.ui.tableView_42.horizontalHeader().hideSection(7)
        self.ui.tableView_42.setColumnWidth(8, 300)
        # self.ui.tableView_11.horizontalHeader().hideSection(8)

        # 计量单位 下拉
        qualities = ["", "m", "m2", "m3", "T", "kg", "个", "项", "座"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_42.setItemDelegateForColumn(5, self.UnitOfMeasurement)  # 计量单位

    # 绘图工程量明细表
    def __init_tableView_43(self):
        self.itemModel43 = QStandardItemModel(6, 8, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel43)  # 初始化 数据模型为 空值
        # 计算结果temp 用于暂存去除【注释】的结果
        headerList = ['楼层号\n或分类', '计算表达式\n中文括号【注释】\t\t\t\t尖括号<楼层>引用',
                      '计算结果real\n隐藏', '计算结果', '小计\n楼 层', '不计\n标志', '公式\n错误', '备注']
        self.itemModel43.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel43 = QItemSelectionModel(self.itemModel43)  # itemModel 选择模型

        self.ui.tableView_43.setModel(self.itemModel43)  # 设置数据模型
        self.ui.tableView_43.setSelectionModel(self.selectionModel43)  # 设置选择模型

        # tb=QTableView  # 表头换行
        self.ui.tableView_43.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_43.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_43.setSelectionMode(oneOrMore)  # 可多选
        #
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_43.setSelectionBehavior(itemOrRow)  # 单元格选择

        # self.ui.tableView_12.verticalHeader().setDefaultSectionSize(28)  # 缺省行高
        self.ui.tableView_43.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_43.setAlternatingRowColors(True)  # 交替行颜色

        # self.ui.tableView_12.setStyleSheet("QTableView{border:1px solid #014F84}")

        # 设置表头边框样式
        self.ui.tableView_43.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"
            "QTableView::item{selection-background-color:#3399FF}")
        self.ui.tableView_43.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        # self.ui.tableView_12.setStyleSheet("selection-background-color:lightBlue")   # 单单元格选中变色  且光标离开保留颜色

        self.ui.tableView_43.setColumnWidth(0, 80)
        self.ui.tableView_43.setColumnWidth(1, 800)
        # self.ui.tableView_43.setColumnWidth(2, 100)  # 结算结果的real  后期 隐藏
        self.ui.tableView_43.horizontalHeader().hideSection(2)
        self.ui.tableView_43.setColumnWidth(3, 100)
        self.ui.tableView_43.setColumnWidth(4, 100)
        self.ui.tableView_43.setColumnWidth(5, 40)
        self.ui.tableView_43.setColumnWidth(6, 40)
        self.ui.tableView_43.setColumnWidth(7, 200)
        for i in range(self.itemModel43.rowCount()):
            item = QStandardItem(self.__NoCalTitle)  # 最后一列
            item.setFlags(self.__NoCalFlags)
            item.setCheckable(True)  # 非锁定
            item.setCheckState(Qt.Unchecked)  # 非勾选
            self.itemModel43.setItem(i, 5, item)  # 设置最后一列的item

    # 历史操作 初始化
    def __init_listView_undo(self):
        self.itemModel_listView_undo = QStringListModel(self)
        # lll = ["1", "2", "3"]
        # self.itemModel_listView_undo.setStringList(lll)
        self.ui.listView.setModel(self.itemModel_listView_undo)
        trig = (QAbstractItemView.DoubleClicked |QAbstractItemView.SelectedClicked)
        self.ui.listView.setEditTriggers(trig)

    # https://www.jb51.net/article/163874.htm 带全选 不带全选
    # 钢筋工程量表 发送至清单表 带全选下拉复选菜单
    def __init_comboBoxIron(self):
        items = [x for x in self.dictIron3.keys()]
        self.items = ["全选"] + items  # items list
        self.box_list = []  # selected items
        self.text = QLineEdit()  # use to selected items
        self.state = 0  # use to record state
        # print(type(items), items)
        # print(type(self.items), self.items)
        q = QListWidget()
        for i in range(len(self.items)):
            self.box_list.append(QCheckBox())
            self.box_list[i].setText(self.items[i])
            item = QListWidgetItem(q)
            q.setItemWidget(item, self.box_list[i])
            if i == 0:
                self.box_list[i].stateChanged.connect(self.all_selected)
            else:
                self.box_list[i].stateChanged.connect(self.show_selected)
        q.setStyleSheet("font-size: 12px; font-weight: bold; height: 25px; margin-left: 1px")
        # width 滑动条的宽度  height combobox的高度
        self.ui.comboBoxIron.setStyleSheet("width: 20px; height: 25px; font-size: 12px; font-weight: bold")
        self.text.setReadOnly(True)  # 下拉框只读
        self.ui.comboBoxIron.setLineEdit(self.text)
        self.ui.comboBoxIron.setModel(q.model())
        self.ui.comboBoxIron.setView(q)
        self.ui.comboBoxIron = QComboBox
        str12 = self.items[1] + '；' + self.items[2]   # 第一选项和第二选项 拼接 可选
        self.text.setText(str12)  # 设置下拉框的默认值 可选

    # 辅助 __init_comboBoxIron
    def all_selected(self):
        """
        decide whether to check all
        :return:
        """
        # change state
        if self.state == 0:   # 全选
            self.state = 1
            for i in range(1, len(self.items)):
                self.box_list[i].setChecked(True)
        else:   # 全消除
            self.state = 0
            for i in range(1, len(self.items)):
                self.box_list[i].setChecked(False)
        self.show_selected()

    # 辅助 __init_comboBoxIron
    def get_selected(self) -> list:
        """
        get selected items
        :return:
        """
        ret = []
        for i in range(1, len(self.items)):
            if self.box_list[i].isChecked():
                ret.append(self.box_list[i].text())
        return ret

    # 辅助 __init_comboBoxIron
    def show_selected(self):
        """
        show selected items
        :return:
        """
        self.text.clear()
        ret = '；'.join(self.get_selected())
        self.text.setText(ret)

    # 信号与函数连接 工程量表
    def connectAll(self):
        self.selectionModel11.currentChanged.connect(self.selectionModel11_currentChanged)
        self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged)
        self.selectionModel12.currentChanged.connect(self.selectionModel12_currentChanged)
        self.itemModel12.itemChanged.connect(self.itemModel12_itemChanged)

        self.selectionModel2.currentChanged.connect(self.selectionModel2_currentChanged)
        self.itemModel2.itemChanged.connect(self.itemModel2_itemChanged)

        self.selectionModel31.currentChanged.connect(self.selectionModel31_currentChanged)
        self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged)
        self.selectionModel32.currentChanged.connect(self.selectionModel32_currentChanged)
        self.itemModel32.itemChanged.connect(self.itemModel32_itemChanged)
        self.selectionModel33.currentChanged.connect(self.selectionModel33_currentChanged)
        self.itemModel33.itemChanged.connect(self.itemModel33_itemChanged)

        self.selectionModel41.currentChanged.connect(self.selectionModel41_currentChanged)
        self.itemModel41.itemChanged.connect(self.itemModel41_itemChanged)
        self.selectionModel42.currentChanged.connect(self.selectionModel42_currentChanged)
        self.itemModel42.itemChanged.connect(self.itemModel42_itemChanged)
        self.selectionModel43.currentChanged.connect(self.selectionModel43_currentChanged)
        self.itemModel43.itemChanged.connect(self.itemModel43_itemChanged)

        self.ui.tableView_31.doubleClicked.connect(self.tableView_31_doubleClicked)  # 钢筋文件表双击事件
        self.ui.tableView_41.doubleClicked.connect(self.tableView_41_doubleClicked)  # 绘图文件表双击事件
        # self.ui.btnOK.clicked.connect(QCoreApplication.instance().quit)

    def disconnectAll(self):
        self.selectionModel11.currentChanged.disconnect(self.selectionModel11_currentChanged)
        self.itemModel11.itemChanged.disconnect(self.itemModel11_itemChanged)
        self.selectionModel12.currentChanged.disconnect(self.selectionModel12_currentChanged)
        self.itemModel12.itemChanged.disconnect(self.itemModel12_itemChanged)
        self.selectionModel2.currentChanged.disconnect(self.selectionModel2_currentChanged)
        self.itemModel2.itemChanged.disconnect(self.itemModel2_itemChanged)

        self.selectionModel31.currentChanged.disconnect(self.selectionModel31_currentChanged)
        self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)
        self.selectionModel32.currentChanged.disconnect(self.selectionModel32_currentChanged)
        self.itemModel32.itemChanged.disconnect(self.itemModel32_itemChanged)
        self.selectionModel33.currentChanged.disconnect(self.selectionModel33_currentChanged)
        self.itemModel33.itemChanged.disconnect(self.itemModel33_itemChanged)

        self.selectionModel41.currentChanged.disconnect(self.selectionModel41_currentChanged)
        self.itemModel41.itemChanged.disconnect(self.itemModel41_itemChanged)
        self.selectionModel42.currentChanged.disconnect(self.selectionModel42_currentChanged)
        self.itemModel42.itemChanged.disconnect(self.itemModel42_itemChanged)
        self.selectionModel43.currentChanged.disconnect(self.selectionModel43_currentChanged)
        self.itemModel43.itemChanged.disconnect(self.itemModel43_itemChanged)

        self.ui.tableView_31.doubleClicked.disconnect(self.tableView_31_doubleClicked)  # 钢筋文件表双击事件
        self.ui.tableView_41.doubleClicked.disconnect(self.tableView_41_doubleClicked)  # 绘图文件表双击事件
        # self.ui.spinBoxSize.valueChanged.disconnect(self.spinBoxSize_valueChanged)

    # TODO  ==============自定义功能函数============
    # # 向主界面发送数据
    # def emitAllTableToMain(self):
    #     # 把模型变为list
    #     Table0List = self.itemModel_listView_undo.stringList()
    #     Table11List = self.getModelToList(self.itemModel11)
    #     Table2List = self.getModelToList(self.itemModel2, 6)
    #     Table31List = self.getModelToList(self.itemModel31)
    #     Table32List = self.getModelToList(self.itemModel32)
    #     Table41List = self.getModelToList(self.itemModel41)
    #     Table42List = self.getModelToList(self.itemModel42)
    #     self.quantitiesDict["历史操作"] = Table0List
    #     self.quantitiesDict["工程量总表"] = Table11List
    #     self.quantitiesDict["分类表"] = Table2List
    #     self.quantitiesDict["钢筋工程量文件表"] = Table31List
    #     self.quantitiesDict["钢筋工程量表"] = Table32List
    #     self.quantitiesDict["绘图工程量文件表"] = Table41List
    #     self.quantitiesDict["绘图工程量表"] = Table42List
        # return dictAllTable
        # if Table3List:
        #     self.quantitiesTable3List.emit(Table3List)  # 发射信号

    # 钢筋表 土建表 合并 整理成二维list 辅助 公用
    def ironGlodonSendMaintable(self, model, source):
        rows = model.rowCount()
        mainrows = self.itemModel11.rowCount()
        # listmain = [["", "", "", "", "", "", "", [], ""]]  # 准备写入总表的 二维列表形式数据
        dictmain = {}  # 准备写入总表的 字典形式数据
        for row in range(rows):
            # 提取一些参数
            rfid = model.item(row, 1).text()  # 文件序号
            rfName = model.item(row, 3).text()  # 报表名称
            features = model.item(row, 4).text()  # 特征描述
            if not features:  # 如果没有子目名称就 跳过
                continue
            unit = model.item(row, 5).text()  # 计量单位
            Quan = model.item(row, 6).text()  # 工程量
            if Quan:  # 有工程量时
                if unit == "kg":
                    Quanf = round(float(Quan)/1000, 4)  # 用于参与累加
                    Quan = str(Quanf)
                    unit = "t"
                elif unit == "个":
                    Quanf = int(Quan)
                    Quan = str(Quanf)
                else:
                    Quanf = round(float(Quan), 4)
                    Quan = str(Quanf)
            else:
                # Quanf = 0
                continue
            notes = model.item(row, 8).text()  # 备注

            # 字典形式写入
            flist = [rfName, Quan + '【' + rfid + ' 第' + str(row + 1) + '行】',
                     Quan, Quan, Quan, '', '', notes]  # 组合一个明细表
            mergeKeyStr = features + '_' + unit
            result = dictmain.get(mergeKeyStr)
            if result == None:  # 没有key
                quanlist = ["", source, features, features, '', unit, Quanf, [], ""]
                quanlist[7].append(flist)
                dictmain[mergeKeyStr] = quanlist
            else:  # 找到已匹配的key
                quanlist = dictmain[mergeKeyStr]  # 先取出此清单
                quanlist[7].append(flist)
                quanlist[6] += Quanf
        return dictmain

    # 整理后的字典（钢筋 土建分表） 发至总表
    def subtable_to_main(self, dictdata):
        if len(dictdata) < 1:
            QMessageBox.information(self, "失败！:", "原始数据有误，请检查", QMessageBox.Ok)
            return
            # 断开链接
        self.itemModel11.itemChanged.disconnect(self.itemModel11_itemChanged)
        # 排序输出
        sortstr = [x for x in dictdata.keys()]
        sortstr.sort()
        # 循环
        num = 1
        for name in sortstr:  # 循环每一条清单+单位组合名称
            itemlist = []
            for col in range(len(dictdata[name])):
                if col == 0:  # 第一个序号
                    item = QStandardItem(str(num))
                elif col == 6:  # 第一个序号
                    item = QStandardItem(str(round(dictdata[name][col], 4)))
                else:
                    item = QStandardItem(str(dictdata[name][col]))
                itemlist.append(item)
            num += 1
            self.itemModel11.appendRow(itemlist)

        # rows = len(gcllist)
        # cols = len(gcllist[0])
        # zrow = self.itemModel11.rowCount()  # 主表的最大行
        # for row in range(rows):
        #     itemlist = []
        #     for col in range(cols):
        #         temp = gcllist[row][col]
        #         item = QStandardItem(str(temp))
        #         itemlist.append(item)
        #     self.itemModel11.appendRow(itemlist)
        # ~~~~~~最后做的几项收尾工作~~~~~~~~~
        self.changeColor()  # 《来源》列 变色
        self.m3ChangeRed(self.itemModel11)  # m3 变色
        # 刷新一下 总计
        self.quantitiesSum()  # 所有行的 汇总总计
        # 恢复链接
        self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged)

    # 总表 所有行的 汇总总计
    def quantitiesSum(self):
        rows = self.itemModel11.rowCount()
        cols = self.itemModel11.columnCount()
        sumfloat = 0.00
        for row in range(rows):
            item = self.itemModel11.item(row, 6)
            try:
                item = float(item.text())
                sumfloat += item
            except:
                continue
        sumfloat = str(round(sumfloat, 3))
        self.ui.label_sum.setText(sumfloat)

    # 钢筋表 所有行的 汇总总计
    def quantitiesSum32(self):
        rows = self.itemModel32.rowCount()
        cols = self.itemModel32.columnCount()
        sumfloat = 0.00
        for row in range(rows):
            item = self.itemModel32.item(row, 6)
            try:
                item = float(item.text())
                sumfloat += item
            except:
                continue
        sumfloat = str(round(sumfloat, 3))
        self.ui.label_sum32.setText(sumfloat)

    # 土建表 所有行的 汇总总计
    def quantitiesSum42(self):
        rows = self.itemModel42.rowCount()
        cols = self.itemModel42.columnCount()
        sumfloat = 0.00
        for row in range(rows):
            item = self.itemModel42.item(row, 6)
            try:
                item = float(item.text())
                sumfloat += item
            except:
                continue
        sumfloat = str(round(sumfloat, 4))
        self.ui.label_sum42.setText(sumfloat)


    # 返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制
    # 表名 1 面积表 2 措施表 31 分部分项表 32 明细表 4 分类表
    def viewCurrentModel(self):
        # print("开始判断模型")
        tablename = None
        selectmodelRows11 = len(self.selectionModel11.selectedRows())
        selectmodelRows12 = len(self.selectionModel12.selectedRows())
        selectmodelRows2 = len(self.selectionModel2.selectedRows())

        selectmodelRows31 = len(self.selectionModel31.selectedRows())
        selectmodelRows32 = len(self.selectionModel32.selectedRows())
        selectmodelRows33 = len(self.selectionModel33.selectedRows())

        selectmodelRows41 = len(self.selectionModel41.selectedRows())
        selectmodelRows42 = len(self.selectionModel42.selectedRows())
        selectmodelRows43 = len(self.selectionModel43.selectedRows())
        # 11 12 先判断 明细表再主表，  应为 主表点击 明细表选中状态取消
        if selectmodelRows12:
            tablename = 12
            tableobj = self.ui.tableView_12
            model = self.itemModel12
            selectModel = self.selectionModel12
        elif selectmodelRows11:
            tablename = 11
            tableobj = self.ui.tableView_11
            model = self.itemModel11
            selectModel = self.selectionModel11
        elif selectmodelRows2:
            tablename = 2
            tableobj = self.ui.tableView_2
            model = self.itemModel2
            selectModel = self.selectionModel2

        elif selectmodelRows31:
            tablename = 31
            tableobj = self.ui.tableView_31
            model = self.itemModel31
            selectModel = self.selectionModel31
        elif selectmodelRows33:
            tablename = 33
            tableobj = self.ui.tableView_33
            model = self.itemModel33
            selectModel = self.selectionModel33
        elif selectmodelRows32:
            tablename = 32
            tableobj = self.ui.tableView_32
            model = self.itemModel32
            selectModel = self.selectionModel32

        elif selectmodelRows41:
            tablename = 41
            tableobj = self.ui.tableView_41
            model = self.itemModel41
            selectModel = self.selectionModel41
        elif selectmodelRows43:
            tablename = 43
            tableobj = self.ui.tableView_43
            model = self.itemModel43
            selectModel = self.selectionModel43
        elif selectmodelRows42:
            tablename = 42
            tableobj = self.ui.tableView_42
            model = self.itemModel42
            selectModel = self.selectionModel42
        else:
            # print("表1 未获得焦点")
            return
        rowsIndexList = [i.row() for i in selectModel.selectedRows()]
        rowsIndexList.sort(reverse=True)
        return tablename, tableobj, model, selectModel, rowsIndexList

    # 设置房号名称
    def setBuildName(self, namestr):
        self.ui.label_2.setText(namestr)

    # 把字典数据 变为  模型
    def setDictToModel(self, dict):
        for key, value in dict.items():
            if not value:  # 空表就跳过
                return
            self.disconnectAll()
            if key == "历史操作":
                self.itemModel_listView_undo.setStringList(value)
            elif key == "工程量总表":
                self.listToModel(value, self.itemModel11)
                self.tabindex = "工程量总表"
                self.m3ChangeRed(self.itemModel11)
                self.quantitiesSum()  # 所有行的 汇总总计
                self.changeColor()
            elif key == "分类表":
                self.listToModel(value, self.itemModel2, 6)
                self.tabindex = "分类表"
                self.m3ChangeRed(self.itemModel2)
            elif key == "钢筋工程量文件表":
                self.listToModel(value, self.itemModel31)
            elif key == "钢筋工程量表":
                self.listToModel(value, self.itemModel32)
                self.tabindex = "钢筋工程量表"
                self.m3ChangeRed(self.itemModel32)
                self.quantitiesSum32()
            elif key == "绘图工程量文件表":
                self.listToModel(value, self.itemModel41)
            elif key == "绘图工程量表":
                self.listToModel(value, self.itemModel42)
                self.tabindex = "绘图工程量表"
                self.m3ChangeRed(self.itemModel42)
                self.quantitiesSum42()
            self.connectAll()

    def listToModel(self, templist, model, nocol=None):
        rows = len(templist)
        cols = len(templist[0])
        model.setRowCount(rows)  # 1 重置行数
        # print('开始循环读取 写入')
        for x in range(rows):  # 循环每一行 x 代表一行 列表形式的数据
            for y in range(cols):  # 循环每一个列  y代表每个数据
                str1 = templist[x][y]
                if y == nocol:
                    item = QStandardItem(self.__NoCalTitle)
                    item.setFlags(self.__NoCalFlags)
                    item.setCheckable(True)  # 非锁定
                    if str1 == "1":
                        item.setCheckState(Qt.Checked)  # 勾选
                    else:
                        item.setCheckState(Qt.Unchecked)  # 勾选
                    model.setItem(x, y, item)  # 赋值
                else:
                    item = QStandardItem(str1)  # 需要给一个str格式
                    model.setItem(x, y, item)  # 赋值

    # 把模型 变为 列表 ，函数返回一个列表
    def getModelToList(self, model, Nocol=None):
        if not model:
            return None
        rows = model.rowCount()
        cols = model.columnCount()
        tablelist = []
        for row in range(rows):
            for col in range(cols):
                if col == 0:
                    tablelist.append([])
                item = model.item(row, col)
                if col == Nocol:
                    if item.checkState() == Qt.Checked:
                        item = "1"
                    else:
                        item = ""
                    tablelist[row].append(item)
                else:
                    tablelist[row].append(item.text())
        return tablelist

    # 选项卡被点击时触发 清单表 分析表
    def tabClicked_1(self, index):
        # print('tabClicked_1')
        result = self.ui.tabWidget_1.tabText(index)
        self.tabindex = result
        # 取消掉选择行
        self.selectionModel11.clearSelection()
        self.selectionModel12.clearSelection()
        self.selectionModel2.clearSelection()
        self.selectionModel31.clearSelection()
        self.selectionModel32.clearSelection()
        self.selectionModel33.clearSelection()
        self.selectionModel41.clearSelection()
        self.selectionModel42.clearSelection()
        self.selectionModel43.clearSelection()
        # 粘贴行 菜单不可用
        # self.ui.QuanBtnCopy.setEnabled(False)
        self.ui.QuanBtnPaste.setEnabled(False)
        # self.ui.QuanBtnDelRow.setEnabled(False)
        # self.ui.QuanBtnInsertRow.setEnabled(False)

    # 选项卡被点击时触发 清单表 分析表
    def tabWidget_1_currentChanged(self, index):
        # print("tabWidget_1_currentChanged", index)
        self.ui.QuanBtnImportClassification.setEnabled(False)  # 分类表 导入
        tabname = self.ui.tabWidget_1.tabText(index)
        if tabname == "绘图工程量表":
            # print("绘图工程量表 被点击！")
            path = r'data\quan_data\默认设置.pkl'
            result = os.path.isfile(path)
            if not result:
                self.ui.QuanBtnImportGlodon_set.setEnabled(False)
                # self.ui.QuanBtnGlodonImport.setEnabled(False)
                QMessageBox.critical(self, "错误", "匹配表缺失，请整体解压程序包！", QMessageBox.Cancel)
                return
            else:
                self.ui.QuanBtnImportGlodon_set.setEnabled(True)
                # self.ui.QuanBtnGlodonImport.setEnabled(True)
            pickle_file = open(path, 'rb')
            self.matchingdict = pickle.load(pickle_file)  # 取得匹配表数据 self.matchingList 字典形式｛默认表：[[],[],[]]｝
            pickle_file.close()
            tabnums = len(self.matchingdict)
            if not tabnums:
                QMessageBox.critical(self, "错误", "匹配表缺失，请整体解压程序包！", QMessageBox.Cancel)
                return
            self.ui.comboBox.clear()
            nameslist = self.matchingdict.keys()
            for name in nameslist:
                if name == "其他":
                    continue
                self.ui.comboBox.addItem(name)

    #  写入历史操作的功能
    def undoredo_listview_write(self):
        lastRow = self.itemModel_listView_undo.rowCount()
        self.itemModel_listView_undo.insertRow(lastRow)  # 在尾部插入一空行
        index = self.itemModel_listView_undo.index(lastRow, 0)  # 获取最后一行的ModelIndex
        self.itemModel_listView_undo.setData(index, f"{self.beforestrundo[1]} to {self.afterstrundo}")  # 设置显示文字
        self.ui.listView.setCurrentIndex(index)  # 设置当前选中的行

    # 不计标志变色 辅助 参数： 模型，不计标志的列，表达式的列号,最后两个列号是需要归0
    def noCalcuColorChange(self, model, noCol, noColstr, zeroCol=[2, 4]):
        rows = model.rowCount()
        for row in range(rows):
            boolstr = model.item(row, noCol)  # 不计标志 单元格对象
            strtemp = model.item(row, noColstr).text()  # 要变色的单元格文字
            if boolstr.checkState() == Qt.Checked:  # 勾选了不计标志
                # 真实的计算结果real 部位 楼层 归0
                for x in zeroCol:
                    item = QStandardItem('')
                    model.setItem(row, x, item)
                # item = QStandardItem('')
                # model.setItem(row, resultRealcol, item)
                # item = QStandardItem('')
                # model.setItem(row, floor, item)
                # item = QStandardItem('')
                # model.setItem(row, position, item)
                # 表达式 变红
                item = QStandardItem(strtemp)
                font = item.font()
                font.setBold(True)
                item.setForeground(QColor(200, 0, 0))
                # item.setBackground(QColor(200, 0, 0))
                item.setFont(font)
                model.setItem(row, noColstr, item)
            else:  # 未勾选了 不计标志
                item = QStandardItem(strtemp)
                font = item.font()
                font.setBold(False)
                item.setForeground(QColor(0, 0, 0))
                # item.setBackground(QColor(255, 255, 255))
                item.setFont(font)
                model.setItem(row, noColstr, item)

    # 表达式辅助
    def expression_machining(self, str):
        # print("expression_machining", str)
        # ~~~~~~~~~~~~~~~~~~【注释】 中文圆括号 处理~~~~~~~~~~~~~~~~~~~~~
        strtemp = re.sub(r'\u3010[^\u3010]*\u3011', '', str)  # 【注释】 删除
        strtemp = re.sub('（', '(', strtemp)  # 中文圆括号替换英文 删除
        strtemp = re.sub('）', ')', strtemp)  # 中文圆括号替换英文 删除
        # 此时 注释 中文括号 已经处理完毕 只剩下<楼层>引用
        # 尝试转换公式 处理楼层在 主函数处理
        try:  # 传入的非表达式  无法转换为eval则出错
            strtempstr = eval(strtemp)  # 表达式 转结果
            strtempfloat = round(strtempstr, 3)  # 保留三位小数
            return "float", strtempfloat
        except:
            return "待处理", strtemp  # 可能有错误 可能有楼层引用

    def m3ChangeRed(self, model, row=None):
        rows = model.rowCount()
        # 判断 计量单位所在的列号
        if self.tabindex == "工程量总表":
            col = 5
        elif self.tabindex == "钢筋工程量表":
            col = 5
        elif self.tabindex == "绘图工程量表":
            col = 5
        elif self.tabindex == "分类表":
            col = 3
        else:
            return
        if row != None:  # 说明是 主表数据变化传进来的
            # print("单行 m3m2 变色 开始")
            curstr = model.item(row, col).text()
            item = QStandardItem(curstr)
            font = item.font()
            if curstr == 'm3':  # 首列字符 粗体
                font.setBold(True)
                colorstr = QColor(200, 0, 0)
            elif curstr == 'm2':  # 首列字符 粗体
                font.setBold(True)
                colorstr = QColor(0, 200, 200)
            else:
                font.setBold(False)
                colorstr = QColor(0, 0, 0)  # 黑色
            model.item(row, col).setForeground(colorstr)  # 设置字体颜色
            model.item(row, col).setFont(font)  # 设置粗体
            # print("单行 m3m2 变色 完成")
        else:  # 说明是初始化 判断全部m3
            for x in range(rows):
                curstr = model.item(x, col).text()
                item = QStandardItem(curstr)
                font = item.font()
                if curstr == 'm3':  # 首列字符 粗体
                    font.setBold(True)
                    colorstr = QColor(200, 0, 0)
                elif curstr == 'm2':  # 首列字符 粗体
                    font.setBold(True)
                    colorstr = QColor(0, 200, 200)
                else:
                    font.setBold(False)
                    colorstr = QColor(0, 0, 0)  # 黑色
                model.item(x, col).setForeground(colorstr)  # 设置字体颜色
                model.item(x, col).setFont(font)  # 设置粗体

    # 《来源》列 变色
    def changeColor(self):
        rows = self.itemModel11.rowCount()
        cols = self.itemModel11.columnCount()
        for row in range(rows):
            temp = self.itemModel11.item(row, 1).text()
            if temp.startswith('分类表'):
                colorstr = QColor(255,205,66)
            elif temp.startswith('绘图工程量表'):
                colorstr = QColor(51,196,129)
            elif temp.startswith('钢筋工程量表'):
                colorstr = QColor(179,123,215)
            elif temp.startswith('合并工程量'):
                # colorstr = QColor(151, 255, 255)
                colorstr = QColor(51,136,255)
            else:
                colorstr = self.itemModel11.item(row, 0).background()
            self.itemModel11.item(row, 1).setBackground(colorstr)

    # 工程量清单明细表内 “楼层”小计
    def floorPositionSum(self, rows, model):
        floorstrRow = [None, 0]  # 行号，小计
        for row in range(rows):
            temp1 = model.item(row, 0).text().strip()  # 取出楼号
            realNum  = model.item(row, 2).text()  # 取出计算结果
            if realNum:  #如果有数
                try:
                    realNum = round(float(realNum), 3)  # 变成浮点数
                except:
                    realNum = "str"
            else:
                realNum = 0
            # 开始汇总计算
            if floorstrRow[0] == None and temp1:  # 第一次找到房号名
                floorstrRow[0] = row
                if realNum != "str":
                    floorstrRow[1] += realNum
                model.setItem(floorstrRow[0], 4, QStandardItem(str(round(floorstrRow[1], 3))))
            elif floorstrRow[0] != None and temp1:  # 第二次找到房号名
                model.setItem(floorstrRow[0], 4, QStandardItem(str(round(floorstrRow[1], 3))))  # 写上次楼号
                floorstrRow[0] = row
                if realNum != "str":
                    floorstrRow[1] = realNum
                model.setItem(floorstrRow[0], 4, QStandardItem(str(round(floorstrRow[1], 3))))  # 写这次楼号
            elif floorstrRow[0] != None and not temp1:  # 有楼层号 找到明细行
                if realNum != "str":
                    floorstrRow[1] += realNum
                model.setItem(floorstrRow[0], 4, QStandardItem(str(round(floorstrRow[1], 3))))  # 写这次楼号
            elif floorstrRow[0] == None and not temp1:  # 无楼层号 找到明细行
                pass

    # TODO  ==============自动链接的 槽函数============
    def selectionModel11_currentChanged(self, index):
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        self.ui.QuanBtnImportClassification.setEnabled(False)  # 分类表 导入
        currRow = index.row()  # 当前行
        currCol = index.column()  # 当前列
        self.beforestrundo = (11, self.itemModel11.item(currRow, currCol).text())

        self.itemModel12.itemChanged.disconnect(self.itemModel12_itemChanged)  # 明细表的数据有变化
        self.ui.tableView_12.setEnabled(True)  # 明细表可用
        temp7List = self.itemModel11.item(currRow, 7).text()  # 获取当前行第7列数据 明细表
        temp3 = self.itemModel11.item(currRow, 3).text().strip()  # 获取 清单项目名称
        temp6 = self.itemModel11.item(currRow, 6).text()  # 获取 清单 当前行工程量
        self.ui.label_6.setText(f'{currRow + 1} ')
        self.ui.label_7.setText(f"行；项名称：{temp3}")
        self.ui.label_3.setText(temp6)

        if temp7List:  # 如果明细表数据不为空,则载入明细到明细表
            templist = ast.literal_eval(temp7List)  # str形式列表转化为真正的二维list
            rows = len(templist)  # 二维表共计多少行
            cols = len(templist[0])
            self.itemModel12.setRowCount(rows)  # 1 重置行数
            # print('开始循环读取 写入')
            for x in range(rows):  # 循环每一行 x 代表一行 列表形式的数据
                for y in range(cols):  # 循环每一个列  y代表每个数据
                    str1 = templist[x][y]
                    if y == 5:
                        item = QStandardItem(self.__NoCalTitle)
                        item.setFlags(self.__NoCalFlags)
                        item.setCheckable(True)  # 非锁定
                        if str1 == "1":
                            item.setCheckState(Qt.Checked)  # 勾选
                        else:
                            item.setCheckState(Qt.Unchecked)  # 勾选
                        self.itemModel12.setItem(x, y, item)  # 赋值
                    else:
                        item = QStandardItem(str1)  # 需要给一个str格式
                        self.itemModel12.setItem(x, y, item)  # 赋值
            self.noCalcuColorChange(self.itemModel12, 5, 1)  # “不计标志”变色
        else:  # 第七列数据为空  则初始化一个明细表
            rows = 6
            self.itemModel12.setRowCount(rows)
            self.initItemModelBlank(self.itemModel12)
            for i in range(rows):
                item = QStandardItem(self.__NoCalTitle)  # 最后一列
                item.setFlags(self.__NoCalFlags)
                item.setCheckable(True)  # 非锁定
                item.setCheckState(Qt.Unchecked)  # 非勾选
                self.itemModel12.setItem(i, 5, item)  # 设置最后一列的item
        # 选中附表的第一行为当前行
        self.selectionModel12.clearSelection()  # 点击主表 取消明细表的选中状态，以便复制行
        self.ui.tableView_12.verticalScrollBar().setSliderPosition(0)
        self.itemModel12.itemChanged.connect(self.itemModel12_itemChanged)  # 明细表的数据有变化

    def selectionModel12_currentChanged(self, index):
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        currRow = index.row()  # 当前行
        currCol = index.column()  # 当前列
        self.beforestrundo = (12, self.itemModel12.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # print(self.beforestrundo)
        # print("明细表被点击！")
        # mainrow = self.selectionModel11.currentIndex().row()
        # test = self.itemModel11.item(mainrow, 4).text()
        # self.ui.label_35.setText(f'{mainrow + 1} ')
        # self.ui.label_34.setText(f"项名称：{test}")

    def selectionModel2_currentChanged(self, index):
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        self.ui.QuanBtnImportClassification.setEnabled(True)  # 分类表 导入
        currRow = index.row()  # 当前行
        currCol = index.column()  # 当前列
        self.beforestrundo = (2, self.itemModel2.item(currRow, currCol).text())  # 用于历史操作的原始字符

    def selectionModel31_currentChanged(self, index):
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        self.selectionModel32.clearSelection()
        self.selectionModel33.clearSelection()
        # self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        # currRow = index.row()  # 当前行
        # currCol = index.column()  # 当前列
        # self.beforestrundo = (3, self.itemModel31.item(currRow, currCol).text())  # 用于历史操作的原始字符

    def selectionModel32_currentChanged(self, index):
        self.selectionModel31.clearSelection()
        self.ui.QuanBtnIronImport.setEnabled(True)
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        # self.ui.QuanBtnImportClassification.setEnabled(False)  # 分类表 导入
        currRow = index.row()  # 当前行
        currCol = index.column()  # 当前列
        self.beforestrundo = (32, self.itemModel32.item(currRow, currCol).text())  #历史文字
        self.itemModel33.itemChanged.disconnect(self.itemModel33_itemChanged)  # 明细表的数据有变化
        self.ui.tableView_33.setEnabled(True)  # 明细表可用
        temp7List = self.itemModel32.item(currRow, 7).text()  # 获取当前行第7列数据 明细表
        temp3 = self.itemModel32.item(currRow, 3).text().strip()  # 获取 项目名称
        temp6 = self.itemModel32.item(currRow, 6).text()  # 获取 当前行工程量
        self.ui.label_29.setText(f'{currRow + 1} ')
        self.ui.label_30.setText(f"行；项名称：{temp3}")
        self.ui.label_32.setText(temp6)
        if temp7List:  # 如果明细表数据不为空,则载入明细到明细表
            templist = ast.literal_eval(temp7List)  # str形式列表转化为真正的二维list
            rows = len(templist)  # 二维表共计多少行
            cols = len(templist[0])
            self.itemModel33.setRowCount(rows)  # 1 重置行数
            # print('开始循环读取 写入')
            for x in range(rows):  # 循环每一行 x 代表一行 列表形式的数据
                for y in range(cols):  # 循环每一个列  y代表每个数据
                    str1 = templist[x][y]
                    if y == 5:
                        item = QStandardItem(self.__NoCalTitle)
                        item.setFlags(self.__NoCalFlags)
                        item.setCheckable(True)  # 非锁定
                        if str1 == "1":
                            item.setCheckState(Qt.Checked)  # 勾选
                        else:
                            item.setCheckState(Qt.Unchecked)  # 勾选
                        self.itemModel33.setItem(x, y, item)  # 赋值
                    else:
                        item = QStandardItem(str1)  # 需要给一个str格式
                        self.itemModel33.setItem(x, y, item)  # 赋值
            self.noCalcuColorChange(self.itemModel33, 5, 1)  # “不计标志”变色
        else:  # 第七列数据为空  则初始化一个明细表
            rows = 6
            self.itemModel33.setRowCount(rows)
            self.initItemModelBlank(self.itemModel33)
            for i in range(rows):
                item = QStandardItem(self.__NoCalTitle)  # 最后一列
                item.setFlags(self.__NoCalFlags)
                item.setCheckable(True)  # 非锁定
                item.setCheckState(Qt.Unchecked)  # 非勾选
                self.itemModel33.setItem(i, 5, item)  # 设置最后一列的item
        # 选中附表的第一行为当前行
        self.selectionModel33.clearSelection()  # 点击主表 取消明细表的选中状态，以便复制行
        self.ui.tableView_33.verticalScrollBar().setSliderPosition(0)
        self.itemModel33.itemChanged.connect(self.itemModel33_itemChanged)  # 明细表的数据有变化

    def selectionModel33_currentChanged(self, index):
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        currRow = index.row()  # 当前行
        currCol = index.column()  # 当前列
        self.beforestrundo = (33, self.itemModel33.item(currRow, currCol).text())  # 用于历史操作的原始字符

    def selectionModel41_currentChanged(self, index):
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        self.selectionModel42.clearSelection()
        self.selectionModel43.clearSelection()

    def selectionModel42_currentChanged(self, index):
        self.selectionModel41.clearSelection()
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        self.ui.QuanBtnGlodonImport.setEnabled(True)  #
        # self.ui.QuanBtnImportClassification.setEnabled(False)  # 分类表 导入
        currRow = index.row()  # 当前行
        currCol = index.column()  # 当前列
        self.beforestrundo = (42, self.itemModel42.item(currRow, currCol).text())  #历史文字

        self.itemModel43.itemChanged.disconnect(self.itemModel43_itemChanged)  # 明细表的数据有变化
        self.ui.tableView_43.setEnabled(True)  # 明细表可用
        temp7List = self.itemModel42.item(currRow, 7).text()  # 获取当前行第7列数据 明细表
        temp3 = self.itemModel42.item(currRow, 3).text().strip()  # 获取 项目名称
        temp6 = self.itemModel42.item(currRow, 6).text()  # 获取 当前行工程量
        self.ui.label_18.setText(f'{currRow + 1} ')
        self.ui.label_19.setText(f"行；项名称：{temp3}")
        self.ui.label_21.setText(temp6)

        if temp7List:  # 如果明细表数据不为空,则载入明细到明细表
            templist = ast.literal_eval(temp7List)  # str形式列表转化为真正的二维list
            rows = len(templist)  # 二维表共计多少行
            cols = len(templist[0])
            self.itemModel43.setRowCount(rows)  # 1 重置行数
            # print('开始循环读取 写入')
            for x in range(rows):  # 循环每一行 x 代表一行 列表形式的数据
                for y in range(cols):  # 循环每一个列  y代表每个数据
                    str1 = templist[x][y]
                    if y == 5:
                        item = QStandardItem(self.__NoCalTitle)
                        item.setFlags(self.__NoCalFlags)
                        item.setCheckable(True)  # 非锁定
                        if str1 == "1":
                            item.setCheckState(Qt.Checked)  # 勾选
                        else:
                            item.setCheckState(Qt.Unchecked)  # 勾选
                        self.itemModel43.setItem(x, y, item)  # 赋值
                    else:
                        item = QStandardItem(str1)  # 需要给一个str格式
                        self.itemModel43.setItem(x, y, item)  # 赋值
            self.noCalcuColorChange(self.itemModel43, 5, 1)  # “不计标志”变色
        else:  # 第七列数据为空  则初始化一个明细表
            rows = 6
            self.itemModel43.setRowCount(rows)
            self.initItemModelBlank(self.itemModel43)
            for i in range(rows):
                item = QStandardItem(self.__NoCalTitle)  # 最后一列
                item.setFlags(self.__NoCalFlags)
                item.setCheckable(True)  # 非锁定
                item.setCheckState(Qt.Unchecked)  # 非勾选
                self.itemModel43.setItem(i, 5, item)  # 设置最后一列的item
        # 选中附表的第一行为当前行
        self.selectionModel43.clearSelection()  # 点击主表 取消明细表的选中状态，以便复制行
        self.ui.tableView_43.verticalScrollBar().setSliderPosition(0)
        self.itemModel43.itemChanged.connect(self.itemModel43_itemChanged)  # 明细表的数据有变化

    def selectionModel43_currentChanged(self, index):
        self.ui.QuanBtnDelRow.setEnabled(True)  # 删除按钮可用
        currRow = index.row()  # 当前行
        currCol = index.column()  # 当前列
        self.beforestrundo = (43, self.itemModel43.item(currRow, currCol).text())  # 用于历史操作的原始字符

    # 引用编码重复时红色  辅助
    def sameCodeName(self):
        rows = self.itemModel11.rowCount()
        listkeyword = []
        for row in range(rows):
            keystr = self.itemModel11.item(row, 2).text()
            item = QStandardItem(keystr)
            font = item.font()
            if keystr in listkeyword:  # 如果有重复
                # 变红色
                font.setBold(True)
                colorstr = QColor(200, 0, 0)
            else:   # 如果没有重复
                # 黑色
                font.setBold(False)
                colorstr = QColor(0, 0, 0)  # 黑色
            self.itemModel11.setItem(row, 2, QStandardItem(keystr))
            self.itemModel11.item(row, 2).setForeground(colorstr)  # 设置字体颜色
            listkeyword.append(keystr)

    # 工程量总表 有数据变化时
    def itemModel11_itemChanged(self, index):
        self.itemModel11.itemChanged.disconnect(self.itemModel11_itemChanged)  # 主表有变化时触发
        # print("工程量总表数据更改了！")
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel11.rowCount()
        columns = self.itemModel11.columnCount()
        # print(self.beforestrundo)
        # 写入历史操作
        if currCol in [0, 1, 2, 3, 4, 5, 6, 8]:
            if self.beforestrundo[0] == 11:
                self.afterstrundo = self.itemModel11.item(currRow, currCol).text()  # 用于历史操作的原始字符
                # print(self.afterstrundo)
                self.undoredo_listview_write()  # 写入历史操作
                self.beforestrundo = (11, self.itemModel11.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # 《来源》列 变色
        if currCol == 1:
            self.changeColor()
        # 项目名称更改了
        if currCol == 3:
            test = self.itemModel11.item(currRow, 3).text()
            self.ui.label_6.setText(f'{currRow + 1} ')
            self.ui.label_7.setText(f"行；项名称：{test}")
        # m3红色 m2青色
        if currCol == 5:
            # print("计量单位有数据变化")
            self.m3ChangeRed(self.itemModel11, currRow)
        # 工程量 单价如果输入的不是数值类型则清除 并提示, 如果是数字则按控件上的小数位显示
        if currCol == 6:
            itemtext = self.itemModel11.item(currRow, currCol).text().strip()
            try:
                fltemp = float(itemtext)
                item = round(fltemp, 3)
                item = QStandardItem(str(item))
                self.itemModel11.setItem(currRow, currCol, item)
            except:
                itemo = QStandardItem("")
                self.itemModel11.setItem(currRow, currCol, itemo)  # 工程量 单价 为空
            finally:
                pass
        self.sameCodeName()  # 引用编码重复时红色  辅助
        self.quantitiesSum()  # 所有行的 汇总总计
        self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged)  # 主表有变化时触发

    def itemModel12_itemChanged(self, index):
        self.itemModel12.itemChanged.disconnect(self.itemModel12_itemChanged)
        # self.itemModel11.itemChanged.disconnect(self.itemModel11_itemChanged)  # 主表有变化时触发
        currRow = index.row()
        currCol = index.column()
        # print(self.beforestrundo)
        # ~~~~~~~~~~~~~~~ 写入历史操作 ~~~~~~~~~~~~~~~
        if currCol in [0, 1, 7]:
            if self.beforestrundo[0] == 12:
                self.afterstrundo = self.itemModel12.item(currRow, currCol).text()  # 用于历史操作的原始字符
                # print(self.afterstrundo)
                self.undoredo_listview_write()
            self.beforestrundo = (12, self.itemModel12.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # 辅助 明细表 数据有变化 处理所有行行数据
        self.itemModel12_itemChanged_All()
        # print("每一行数据写入总表 完成")
        # 刷新一下 总计
        self.quantitiesSum()  # 所有行的 汇总总计
        # 恢复响应
        self.itemModel12.itemChanged.connect(self.itemModel12_itemChanged)
        # self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged) # 主表有变化时触发

    def itemModel12_itemChanged_All(self):
        rows = self.itemModel12.rowCount()
        columns = self.itemModel12.columnCount()
        table11_row = self.selectionModel11.currentIndex().row()
        # print("明细表数据更改了！主表在：", table11_row, "行")
        # ~~~~~~~~~~~~~~~“计算结果real 计算结果 面积 计算错误 ●”~~~~先清0~~~~~~~~~~~~~~~~~~~~~~~~
        # print("“计算错误 ●”~~~~先清0")
        for row in range(rows):
            self.itemModel12.setItem(row, 2, QStandardItem(""))
            self.itemModel12.setItem(row, 3, QStandardItem(""))
            self.itemModel12.setItem(row, 4, QStandardItem(""))
            self.itemModel12.setItem(row, 6, QStandardItem(""))
        # ~~~~~~~~~~~~~~~“表达式处理”~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # print("表达式处理")
        for row in range(rows):
            # 取出表达式的原始字符
            str_temp = self.itemModel12.item(row, 1).text().strip()  # 取出表达式
            if str_temp == "":  # 表达式为空时   '计算结果real'3, '计算结果'4, 都为空
                self.itemModel12.setItem(row, 2, QStandardItem(""))
                self.itemModel12.setItem(row, 3, QStandardItem(""))
            else:   # 表达式的有内容进一步处理
                # print("# 注释 中文括号 处理，楼层引用暂不处理")
                result = self.expression_machining(str_temp)
                if result[0] == "float":
                    # "如果可以转换为 数值"
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel12.setItem(row, 2, item)  # 能转化为结果 则直接赋值 给 计算结果列
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel12.setItem(row, 3, item)  # 结算结果real
                elif result[0] == "待处理":  # 把待处理的字符放在 temp 格
                    str_temp = result[1]
                    item = QStandardItem(str_temp)  # 需要给一个str格式
                    self.itemModel12.setItem(row, 2, item)  # 待处理字符放在 计算结算real列
                    item = QStandardItem("")
                    self.itemModel12.setItem(row, 3, item)  # 清空
        # 处理可能存在的楼层引用前，先 楼层汇总一次
        # print("楼层汇总第一次")
        self.noCalcuColorChange(self.itemModel12, 5, 1)  # “不计标志”变色 真实计算结果归0 部位 楼层归0
        self.floorPositionSum(rows, self.itemModel12)
        # 处理 <楼层> 引用
        for row in range(rows):
            # 取出表达式的原始字符  计算结果real
            str_temp3 = self.itemModel12.item(row, 2).text()
            if not str_temp3:  # 空则执行下一行
                continue
            while True:  # 直到没有引用的符号
                strkey = re.search('\<[^\<]*?\>', str_temp3)  # 先用正则搜索
                if not strkey:  # 没有匹配正则 则退出while
                    break  # 退出while
                strkey = strkey.group()[1:-1]  # 脱去两侧的括号
                keystr = "开始找"
                # print('strkey找到，开始找引用楼层', strkey)
                for rowr in range(0, rows):  #
                    floorrowstr = self.itemModel12.item(rowr, 0).text().strip()  # 获取楼层字符
                    if floorrowstr == strkey:
                        str_temp3 = re.sub('\<[^\<]*?\>', self.itemModel12.item(rowr, 4).text(), str_temp3,
                                          count=1)  # 把第6列楼层汇总 替换成[楼层]字符，替换第一处
                        keystr = "找到了"
                        break  # 找到就跳出 楼层字符的循环 for
                if keystr == "开始找":  # 说明没找到 有错误
                    self.itemModel12.setItem(row, 2, QStandardItem(""))  # 赋值 给 计算结果temp列
                    self.itemModel12.setItem(row, 3, QStandardItem(""))
                    self.itemModel12.setItem(row, 6, QStandardItem('●'))
                    break  # 退出while
            # 试着处理计算式 可能有错误的引用
            try:  # 传入的非表达式  无法转换为eval则出错
                strtempstr = eval(str_temp3)  # 表达式 转结果
                strtempfloat = round(strtempstr, 4)  # 保留三位小数
                strtempfin = str(strtempfloat)  # 转成字符
                item = QStandardItem(strtempfin)  # 需要给一个str格式
                self.itemModel12.setItem(row, 2, item)  # 能转化为结果 则直接赋值 给 计算结果列
                item = QStandardItem(strtempfin)
                self.itemModel12.setItem(row, 3, item)
            except:  # 去除【】 [] 以后 还不能表达式转结果  8列 错误列标记
                self.itemModel12.setItem(row, 2, QStandardItem(""))  # 赋值 给 计算结果temp列
                self.itemModel12.setItem(row, 3, QStandardItem(""))
                self.itemModel12.setItem(row, 6, QStandardItem('●'))
        # <楼层> 引用处理后 再 楼层汇总一次
        # print("楼层汇总第二次")
        self.noCalcuColorChange(self.itemModel12, 5, 1)  # “不计标志”变色
        self.floorPositionSum(rows, self.itemModel12)
        self.noCalcuColorChange(self.itemModel12, 5, 1)  # “不计标志”变色
        # ~~~~~~~~~~~~~~~~~每一行数据写入总表~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        listdata = []  # 用于存放数据的空列表
        sumlist = 0  # 本条清单子目的工程量小计
        for x in range(rows):
            for y in range(columns):
                tempstr1 = self.itemModel12.item(x, y)  # 暂时不考虑 前后有空格的数据
                if y == 0:
                    listdata.append([])  # 如果是每一行的开始，则新增一个空列表
                # 不计标志 如果勾选填1 要不然填空
                if y == 5:
                    if (tempstr1.checkState() == Qt.Checked):  # 勾选了不计标志
                        item = '1'
                        listdata[x].append(item)
                    else:
                        item = ''
                        listdata[x].append(item)
                else:
                    listdata[x].append(tempstr1.text())
            strtemp = self.itemModel12.item(x, 4).text()  # 取楼号小计工程量
            if strtemp:
                sumlist += float(strtemp)
        items = str(listdata)
        item = QStandardItem(items)
        self.itemModel11.setItem(table11_row, 7, item)  # 明细写入隐藏列
        if sumlist:
            sumlist = round(sumlist, 3)
            item = QStandardItem(str(sumlist))
            self.itemModel11.setItem(table11_row, 6, item)
            self.ui.label_3.setText(str(sumlist))
        else:
            item = QStandardItem("")
            self.itemModel11.setItem(table11_row, 6, item)
            self.ui.label_3.setText(str(0))

    def itemModel2_itemChanged(self, index):
        self.itemModel2.itemChanged.disconnect(self.itemModel2_itemChanged)  # 主表有变化时触发
        # print("分类表数据更改了！")
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel2.rowCount()
        columns = self.itemModel2.columnCount()
        self.ui.QuanBtnImportClassificationToQuan.setEnabled(True)  # ← 发送至清单
        # print(self.beforestrundo)
        # 写入历史操作
        if currCol in [0, 1, 2, 3, 4, 5, 8]:
            if self.beforestrundo[0] == 2:
                self.afterstrundo = self.itemModel2.item(currRow, currCol).text()  # 用于历史操作的原始字符
                # print(self.afterstrundo)
                self.undoredo_listview_write()  # 写入历史操作
                self.beforestrundo = (2, self.itemModel2.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # m3红色 m2青色
        if currCol == 3:
            # print("计量单位有数据变化")
            self.m3ChangeRed(self.itemModel2, currRow)
        # 工程量 单价如果输入的不是数值类型则清除 并提示, 如果是数字则按控件上的小数位显示
        if currCol == 5:
            itemtext = self.itemModel2.item(currRow, currCol).text().strip()
            try:
                fltemp = float(itemtext)
                item = round(fltemp, 3)
                item = QStandardItem(str(item))
                self.itemModel2.setItem(currRow, currCol, item)
            except:
                itemo = QStandardItem("")
                self.itemModel2.setItem(currRow, currCol, itemo)  # 工程量 单价 为空
            finally:
                pass
        # 不计标志
        if currCol == 4 or currCol == 6:
            # self.noCalcuColorChange(self.itemModel2, 6, 4, [5])  # “不计标志”变色
            # print("计量单位有数据变化")
            item = self.itemModel2.item(currRow, 6)
            itemstrss = self.itemModel2.item(currRow, 4).text()
            itemstr = QStandardItem(itemstrss)
            font = itemstr.font()
            if item.checkState() == Qt.Checked:
                font.setBold(True)
                itemstr.setForeground(QColor(200, 0, 0))
                self.itemModel2.setItem(currRow, 5, QStandardItem(""))
            else:
                font.setBold(False)
                itemstr.setForeground(QColor(0, 0, 0))

                strtemp = re.sub(r'\u3010[^\u3010]*\u3011', '', itemstrss)  # 【注释】 删除
                strtemp = re.sub('（', '(', strtemp)  # 中文圆括号替换英文 删除
                strtemp = re.sub('）', ')', strtemp)  # 中文圆括号替换英文 删除
                # 此时 注释 中文括号 已经处理完毕 只剩下<楼层>引用
                # 尝试转换公式 处理楼层在 主函数处理
                try:  # 传入的非表达式  无法转换为eval则出错
                    strtempstr = eval(strtemp)  # 表达式 转结果
                    strtempfloat = round(strtempstr, 3)  # 保留三位小数
                    self.itemModel2.setItem(currRow, 5, QStandardItem(str(strtempfloat)))
                    self.itemModel2.setItem(currRow, 7, QStandardItem(""))
                except:
                    self.itemModel2.setItem(currRow, 5, QStandardItem(""))
                    self.itemModel2.setItem(currRow, 7, QStandardItem("●"))
            itemstr.setFont(font)
            self.itemModel2.setItem(currRow, 4, itemstr)
        self.itemModel2.itemChanged.connect(self.itemModel2_itemChanged)  # 主表有变化时触发

    def itemModel31_itemChanged(self, index):  # 钢筋工程量文件表
        self.ui.QuanBtnIronImport.setEnabled(True)
        # table31list = self.getModelToList(self.itemModel31)
        # self.quantitiesDict["钢筋工程量文件表"] = table31list

    def itemModel32_itemChanged(self, index):
        self.ui.QuanBtnIronOK.setEnabled(True)
        self.ui.QuanBtnIronImport.setEnabled(True)
        self.itemModel32.itemChanged.disconnect(self.itemModel32_itemChanged)  # 主表有变化时触发
        # print("钢筋量表数据更改了！")
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel32.rowCount()
        columns = self.itemModel32.columnCount()
        # print(self.beforestrundo)
        # 写入历史操作
        if currCol in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]:
            if self.beforestrundo[0] == 32:
                self.afterstrundo = self.itemModel32.item(currRow, currCol).text()  # 用于历史操作的原始字符
                # print(self.afterstrundo)
                self.undoredo_listview_write()  # 写入历史操作
                self.beforestrundo = (32, self.itemModel32.item(currRow, currCol).text())  # 用于历史操作的原始字符

        # table32list = self.getModelToList(self.itemModel32)
        # self.quantitiesDict["钢筋工程量表"] = table32list
        self.quantitiesSum32()  # 钢筋表所有行的 汇总总计
        self.m3ChangeRed(self.itemModel32)
        self.itemModel32.itemChanged.connect(self.itemModel32_itemChanged)  # 主表有变化时触发

    def itemModel33_itemChanged(self, index):
        # print("itemModel33_itemChanged", index)
        self.itemModel33.itemChanged.disconnect(self.itemModel33_itemChanged)
        # self.itemModel11.itemChanged.disconnect(self.itemModel11_itemChanged)  # 主表有变化时触发
        currRow = index.row()
        currCol = index.column()
        # print(self.beforestrundo)
        # ~~~~~~~~~~~~~~~ 写入历史操作 ~~~~~~~~~~~~~~~
        if currCol in [0, 1, 7]:
            if self.beforestrundo[0] == 33:
                self.afterstrundo = self.itemModel33.item(currRow, currCol).text()  # 用于历史操作的原始字符
                print(self.afterstrundo)
                self.undoredo_listview_write()
            self.beforestrundo = (33, self.itemModel33.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # 辅助 明细表 数据有变化 处理所有行行数据
        self.itemModel33_itemChanged_All()
        # print("每一行数据写入总表 完成")
        # 刷新一下 总计
        self.quantitiesSum32()  # 钢筋表所有行的 汇总总计
        # 恢复响应
        self.itemModel33.itemChanged.connect(self.itemModel33_itemChanged)
        # self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged) # 主表有变化时触发

    def itemModel33_itemChanged_All(self):
        rows = self.itemModel33.rowCount()
        columns = self.itemModel33.columnCount()
        table32_row = self.selectionModel32.currentIndex().row()
        # print("明细表数据更改了！主表在：", table32_row, "行")
        # ~~~~~~~~~~~~~~~“计算结果real 计算结果 面积 计算错误 ●”~~~~先清0~~~~~~~~~~~~~~~~~~~~~~~~
        # print("“计算错误 ●”~~~~先清0")
        for row in range(rows):
            self.itemModel33.setItem(row, 2, QStandardItem(""))
            self.itemModel33.setItem(row, 3, QStandardItem(""))
            self.itemModel33.setItem(row, 4, QStandardItem(""))
            self.itemModel33.setItem(row, 6, QStandardItem(""))
        # ~~~~~~~~~~~~~~~“表达式处理”~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # print("表达式处理")
        for row in range(rows):
            # 取出表达式的原始字符
            str_temp = self.itemModel33.item(row, 1).text().strip()  # 取出表达式
            if str_temp == "":  # 表达式为空时   '计算结果real'3, '计算结果'4, 都为空
                self.itemModel33.setItem(row, 2, QStandardItem(""))
                self.itemModel33.setItem(row, 3, QStandardItem(""))
            else:  # 表达式的有内容进一步处理
                # print("# 注释 中文括号 处理，楼层引用暂不处理")
                result = self.expression_machining(str_temp)
                if result[0] == "float":
                    # "如果可以转换为 数值"
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel33.setItem(row, 2, item)  # 能转化为结果 则直接赋值 给 计算结果列
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel33.setItem(row, 3, item)  # 结算结果real
                elif result[0] == "待处理":  # 把待处理的字符放在 temp 格
                    str_temp = result[1]
                    item = QStandardItem(str_temp)  # 需要给一个str格式
                    self.itemModel33.setItem(row, 2, item)  # 待处理字符放在 计算结算real列
                    item = QStandardItem("")
                    self.itemModel33.setItem(row, 3, item)  # 清空
        # 处理可能存在的楼层引用前，先 楼层汇总一次
        # print("楼层汇总第一次")
        self.noCalcuColorChange(self.itemModel33, 5, 1)  # “不计标志”变色 真实计算结果归0 部位 楼层归0
        self.floorPositionSum(rows, self.itemModel33)
        # 处理 <楼层> 引用
        for row in range(rows):
            # 取出表达式的原始字符  计算结果real
            str_temp3 = self.itemModel33.item(row, 2).text()
            if not str_temp3:  # 空则执行下一行
                continue
            while True:  # 直到没有引用的符号
                strkey = re.search('\<[^\<]*?\>', str_temp3)  # 先用正则搜索
                if not strkey:  # 没有匹配正则 则退出while
                    break  # 退出while
                strkey = strkey.group()[1:-1]  # 脱去两侧的括号
                keystr = "开始找"
                # print('strkey找到，开始找引用楼层', strkey)
                for rowr in range(0, rows):  #
                    floorrowstr = self.itemModel33.item(rowr, 0).text().strip()  # 获取楼层字符
                    if floorrowstr == strkey:
                        str_temp3 = re.sub('\<[^\<]*?\>', self.itemModel33.item(rowr, 4).text(), str_temp3,
                                           count=1)  # 把第6列楼层汇总 替换成[楼层]字符，替换第一处
                        keystr = "找到了"
                        break  # 找到就跳出 楼层字符的循环 for
                if keystr == "开始找":  # 说明没找到 有错误
                    self.itemModel33.setItem(row, 2, QStandardItem(""))  # 赋值 给 计算结果temp列
                    self.itemModel33.setItem(row, 3, QStandardItem(""))
                    self.itemModel33.setItem(row, 6, QStandardItem('●'))
                    break  # 退出while
            # 试着处理计算式 可能有错误的引用
            try:  # 传入的非表达式  无法转换为eval则出错
                strtempstr = eval(str_temp3)  # 表达式 转结果
                strtempfloat = round(strtempstr, 4)  # 保留三位小数
                strtempfin = str(strtempfloat)  # 转成字符
                item = QStandardItem(strtempfin)  # 需要给一个str格式
                self.itemModel33.setItem(row, 2, item)  # 能转化为结果 则直接赋值 给 计算结果列
                item = QStandardItem(strtempfin)
                self.itemModel33.setItem(row, 3, item)
            except:  # 去除【】 [] 以后 还不能表达式转结果  8列 错误列标记
                self.itemModel33.setItem(row, 2, QStandardItem(""))  # 赋值 给 计算结果temp列
                self.itemModel33.setItem(row, 3, QStandardItem(""))
                self.itemModel33.setItem(row, 6, QStandardItem('●'))
        # <楼层> 引用处理后 再 楼层汇总一次
        # print("楼层汇总第二次")
        self.noCalcuColorChange(self.itemModel33, 5, 1)  # “不计标志”变色
        self.floorPositionSum(rows, self.itemModel33)
        self.noCalcuColorChange(self.itemModel33, 5, 1)  # “不计标志”变色
        # ~~~~~~~~~~~~~~~~~每一行数据写入总表~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        listdata = []  # 用于存放数据的空列表
        sumlist = 0  # 本条清单子目的工程量小计
        for x in range(rows):
            for y in range(columns):
                tempstr1 = self.itemModel33.item(x, y)  # 暂时不考虑 前后有空格的数据
                if y == 0:
                    listdata.append([])  # 如果是每一行的开始，则新增一个空列表
                # 不计标志 如果勾选填1 要不然填空
                if y == 5:
                    if (tempstr1.checkState() == Qt.Checked):  # 勾选了不计标志
                        item = '1'
                        listdata[x].append(item)
                    else:
                        item = ''
                        listdata[x].append(item)
                else:
                    listdata[x].append(tempstr1.text())
            strtemp = self.itemModel33.item(x, 4).text()  # 取楼号小计工程量
            if strtemp:
                sumlist += float(strtemp)
        items = str(listdata)
        item = QStandardItem(items)
        self.itemModel32.setItem(table32_row, 7, item)  # 明细写入隐藏列
        if sumlist:
            sumlist = round(sumlist, 3)
            item = QStandardItem(str(sumlist))
            self.itemModel32.setItem(table32_row, 6, item)
            self.ui.label_32.setText(str(sumlist))
        else:
            item = QStandardItem("")
            self.itemModel32.setItem(table32_row, 6, item)
            self.ui.label_32.setText(str(0))

    def itemModel41_itemChanged(self, index):  # 钢筋工程量文件表
        self.ui.QuanBtnGlodonImport.setEnabled(True)  # 可用
        # table41list = self.getModelToList(self.itemModel41)
        # self.quantitiesDict["绘图工程量文件表"] = table41list

    def itemModel42_itemChanged(self, index):
        self.ui.QuanBtnGlodonOK.setEnabled(True)  # 土建量表 发送至清单总表
        self.ui.QuanBtnGlodonImport.setEnabled(True)  #
        self.itemModel42.itemChanged.disconnect(self.itemModel42_itemChanged)  # 主表有变化时触发
        # print("钢筋量表数据更改了！")
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel42.rowCount()
        columns = self.itemModel42.columnCount()
        # print(self.beforestrundo)
        # 写入历史操作
        if currCol in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]:
            if self.beforestrundo[0] == 42:
                self.afterstrundo = self.itemModel42.item(currRow, currCol).text()  # 用于历史操作的原始字符
                # print(self.afterstrundo)
                self.undoredo_listview_write()  # 写入历史操作
                self.beforestrundo = (42, self.itemModel42.item(currRow, currCol).text())  # 用于历史操作的原始字符

        # table42list = self.getModelToList(self.itemModel42)
        # self.quantitiesDict["绘图工程量表"] = table42list
        self.quantitiesSum42()  # 钢筋表所有行的 汇总总计
        self.m3ChangeRed(self.itemModel42)
        self.itemModel42.itemChanged.connect(self.itemModel42_itemChanged)  # 主表有变化时触发

    def itemModel43_itemChanged(self, index):
        self.itemModel43.itemChanged.disconnect(self.itemModel43_itemChanged)
        # self.itemModel11.itemChanged.disconnect(self.itemModel11_itemChanged)  # 主表有变化时触发
        currRow = index.row()
        currCol = index.column()
        # print(self.beforestrundo)
        # ~~~~~~~~~~~~~~~ 写入历史操作 ~~~~~~~~~~~~~~~
        if currCol in [0, 1, 7]:
            if self.beforestrundo[0] == 43:
                self.afterstrundo = self.itemModel43.item(currRow, currCol).text()  # 用于历史操作的原始字符
                # print(self.afterstrundo)
                self.undoredo_listview_write()
            self.beforestrundo = (43, self.itemModel43.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # 辅助 明细表 数据有变化 处理所有行行数据
        self.itemModel43_itemChanged_All()

        # print("每一行数据写入总表 完成")
        # 刷新一下 总计
        self.quantitiesSum42()  # 钢筋表所有行的 汇总总计
        # 恢复响应
        self.itemModel43.itemChanged.connect(self.itemModel43_itemChanged)
        # self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged) # 主表有变化时触发

    def itemModel43_itemChanged_All(self):
        rows = self.itemModel43.rowCount()
        columns = self.itemModel43.columnCount()
        table42_row = self.selectionModel42.currentIndex().row()
        # print("明细表数据更改了！主表在：", table42_row, "行")
        # ~~~~~~~~~~~~~~~“计算结果real 计算结果 面积 计算错误 ●”~~~~先清0~~~~~~~~~~~~~~~~~~~~~~~~
        # print("“计算错误 ●”~~~~先清0")
        for row in range(rows):
            self.itemModel43.setItem(row, 2, QStandardItem(""))
            self.itemModel43.setItem(row, 3, QStandardItem(""))
            self.itemModel43.setItem(row, 4, QStandardItem(""))
            self.itemModel43.setItem(row, 6, QStandardItem(""))
        # ~~~~~~~~~~~~~~~“表达式处理”~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # print("表达式处理")
        for row in range(rows):
            # 取出表达式的原始字符
            str_temp = self.itemModel43.item(row, 1).text().strip()  # 取出表达式
            if str_temp == "":  # 表达式为空时   '计算结果real'3, '计算结果'4, 都为空
                self.itemModel43.setItem(row, 2, QStandardItem(""))
                self.itemModel43.setItem(row, 3, QStandardItem(""))
            else:  # 表达式的有内容进一步处理
                # print("# 注释 中文括号 处理，楼层引用暂不处理")
                result = self.expression_machining(str_temp)
                if result[0] == "float":
                    # "如果可以转换为 数值"
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel43.setItem(row, 2, item)  # 能转化为结果 则直接赋值 给 计算结果列
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel43.setItem(row, 3, item)  # 结算结果real
                elif result[0] == "待处理":  # 把待处理的字符放在 temp 格
                    str_temp = result[1]
                    item = QStandardItem(str_temp)  # 需要给一个str格式
                    self.itemModel43.setItem(row, 2, item)  # 待处理字符放在 计算结算real列
                    item = QStandardItem("")
                    self.itemModel43.setItem(row, 3, item)  # 清空
        # 处理可能存在的楼层引用前，先 楼层汇总一次
        # print("楼层汇总第一次")
        self.noCalcuColorChange(self.itemModel43, 5, 1)  # “不计标志”变色 真实计算结果归0 部位 楼层归0
        self.floorPositionSum(rows, self.itemModel43)
        # 处理 <楼层> 引用
        for row in range(rows):
            # 取出表达式的原始字符  计算结果real
            str_temp3 = self.itemModel43.item(row, 2).text()
            if not str_temp3:  # 空则执行下一行
                continue
            while True:  # 直到没有引用的符号
                strkey = re.search('\<[^\<]*?\>', str_temp3)  # 先用正则搜索
                if not strkey:  # 没有匹配正则 则退出while
                    break  # 退出while
                strkey = strkey.group()[1:-1]  # 脱去两侧的括号
                keystr = "开始找"
                # print('strkey找到，开始找引用楼层', strkey)
                for rowr in range(0, rows):  #
                    floorrowstr = self.itemModel43.item(rowr, 0).text().strip()  # 获取楼层字符
                    if floorrowstr == strkey:
                        str_temp3 = re.sub('\<[^\<]*?\>', self.itemModel43.item(rowr, 4).text(), str_temp3,
                                           count=1)  # 把第6列楼层汇总 替换成[楼层]字符，替换第一处
                        keystr = "找到了"
                        break  # 找到就跳出 楼层字符的循环 for
                if keystr == "开始找":  # 说明没找到 有错误
                    self.itemModel43.setItem(row, 2, QStandardItem(""))  # 赋值 给 计算结果temp列
                    self.itemModel43.setItem(row, 3, QStandardItem(""))
                    self.itemModel43.setItem(row, 6, QStandardItem('●'))
                    break  # 退出while
            # 试着处理计算式 可能有错误的引用
            try:  # 传入的非表达式  无法转换为eval则出错
                strtempstr = eval(str_temp3)  # 表达式 转结果
                strtempfloat = round(strtempstr, 4)  # 保留三位小数
                strtempfin = str(strtempfloat)  # 转成字符
                item = QStandardItem(strtempfin)  # 需要给一个str格式
                self.itemModel43.setItem(row, 2, item)  # 能转化为结果 则直接赋值 给 计算结果列
                item = QStandardItem(strtempfin)
                self.itemModel43.setItem(row, 3, item)
            except:  # 去除【】 [] 以后 还不能表达式转结果  8列 错误列标记
                self.itemModel43.setItem(row, 2, QStandardItem(""))  # 赋值 给 计算结果temp列
                self.itemModel43.setItem(row, 3, QStandardItem(""))
                self.itemModel43.setItem(row, 6, QStandardItem('●'))
        # <楼层> 引用处理后 再 楼层汇总一次
        # print("楼层汇总第二次")
        self.noCalcuColorChange(self.itemModel43, 5, 1)  # “不计标志”变色
        self.floorPositionSum(rows, self.itemModel43)
        self.noCalcuColorChange(self.itemModel43, 5, 1)  # “不计标志”变色
        # ~~~~~~~~~~~~~~~~~每一行数据写入总表~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        listdata = []  # 用于存放数据的空列表
        sumlist = 0  # 本条清单子目的工程量小计
        for x in range(rows):
            for y in range(columns):
                tempstr1 = self.itemModel43.item(x, y)  # 暂时不考虑 前后有空格的数据
                if y == 0:
                    listdata.append([])  # 如果是每一行的开始，则新增一个空列表
                # 不计标志 如果勾选填1 要不然填空
                if y == 5:
                    if (tempstr1.checkState() == Qt.Checked):  # 勾选了不计标志
                        item = '1'
                        listdata[x].append(item)
                    else:
                        item = ''
                        listdata[x].append(item)
                else:
                    listdata[x].append(tempstr1.text())
            strtemp = self.itemModel43.item(x, 4).text()  # 取楼号小计工程量
            if strtemp:
                sumlist += float(strtemp)
        items = str(listdata)
        item = QStandardItem(items)
        self.itemModel42.setItem(table42_row, 7, item)  # 明细写入隐藏列
        if sumlist:
            sumlist = round(sumlist, 4)
            item = QStandardItem(str(sumlist))
            self.itemModel42.setItem(table42_row, 6, item)
            self.ui.label_21.setText(str(sumlist))
        else:
            item = QStandardItem("")
            self.itemModel42.setItem(table42_row, 6, item)
            self.ui.label_21.setText(str(0))

    # 钢筋文件表 双击事件
    def tableView_31_doubleClicked(self, index):
        # print("土建文件表 双击事件")
        row = index.row()
        pathstr = self.itemModel31.item(row, 1).text()
        path = None
        if pathstr:
            filepath = pathstr
        if filepath:
            # 判断文件是否存在
            # result = os.path.isfile(filepath)
            result = os.path.exists(filepath)
            if not result:
                QMessageBox.critical(self, "错误", "文件不存在", QMessageBox.Cancel)
                return
            # 打开
            xl = Dispatch("Excel.Application")
            xl.Visible = True  # otherwise excel is hidden
            # newest excel does not accept forward slash in path
            wb = xl.Workbooks.Open(filepath)
            # wb.Close()
            # xl.Quit()

    # 土建文件表 双击事件
    def tableView_41_doubleClicked(self, index):
        # print("土建文件表 双击事件")
        row = index.row()
        pathstr = self.itemModel41.item(row, 1).text()
        path = None
        if pathstr:
            filepath = pathstr
        if filepath:
            # 判断文件是否存在
            # result = os.path.isfile(filepath)
            result = os.path.exists(filepath)
            if not result:
                QMessageBox.critical(self, "错误", "文件不存在", QMessageBox.Cancel)
                return
            # 打开
            xl = Dispatch("Excel.Application")
            xl.Visible = True  # otherwise excel is hidden
            # newest excel does not accept forward slash in path
            wb = xl.Workbooks.Open(filepath)
            # wb.Close()
            # xl.Quit()

    # TODO  ==============控件触发============

##  ==========由connectSlotsByName() 自动连接的槽函数==================        
    # 历史操作 清空 按钮
    @pyqtSlot()
    def on_QuanBtnClearListview_clicked(self):
        self.__init_listView_undo()

    # 全屏
    @pyqtSlot()
    def on_QuanBtnFull_clicked(self):
        self.setWindowState(Qt.WindowMaximized)  # 窗口最大化显示 全屏

    # 刷新主程序的字典文件  需要发送到主表
    @pyqtSlot()
    def on_QuanBtnRefresh_clicked(self):
        # 把模型变为list
        Table0List = self.itemModel_listView_undo.stringList()
        Table11List = self.getModelToList(self.itemModel11)
        Table2List = self.getModelToList(self.itemModel2, 6)
        Table31List = self.getModelToList(self.itemModel31)
        Table32List = self.getModelToList(self.itemModel32)
        Table41List = self.getModelToList(self.itemModel41)
        Table42List = self.getModelToList(self.itemModel42)
        self.quantitiesDict["历史操作"] = Table0List
        self.quantitiesDict["工程量总表"] = Table11List
        self.quantitiesDict["分类表"] = Table2List
        self.quantitiesDict["钢筋工程量文件表"] = Table31List
        self.quantitiesDict["钢筋工程量表"] = Table32List
        self.quantitiesDict["绘图工程量文件表"] = Table41List
        self.quantitiesDict["绘图工程量表"] = Table42List
        self.quantitiesTableLists.emit(self.quantitiesDict)

    # 搜索
    @pyqtSlot()
    def on_QuanBtnFind_clicked(self):
        # print("搜索按钮")
        # 搜索前 先切换页签 和标志
        self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 页面
        self.tabindex = "工程量总表"  # 清单表
        text, okPressed = QInputDialog.getText(self, "工程量总表—搜索", "搜索范围包括：来源 编码 名称 特征 单位 备注以及 对应的明细表\n请输入关键词:（分隔 ， 中英文混用皆可）", QLineEdit.Normal, "")
        rows = self.itemModel11.rowCount()
        if okPressed and text != '':
            # print("OK")
            list1 = text.split(sep='，')
            list1 = ','.join(list1)
            keyList = list1.split(sep=',')
            for row in range(rows):
                itemstr0 = self.itemModel11.item(row, 0).text()
                itemstr1 = self.itemModel11.item(row, 1).text()
                itemstr2 = self.itemModel11.item(row, 2).text()
                itemstr3 = self.itemModel11.item(row, 3).text()
                itemstr4 = self.itemModel11.item(row, 4).text()
                itemstr5 = self.itemModel11.item(row, 5).text()
                itemstr7 = self.itemModel11.item(row, 7).text()
                itemstr8 = self.itemModel11.item(row, 8).text()
                str1_8 = itemstr0+itemstr1+itemstr2+itemstr3+itemstr4+itemstr5+itemstr7+itemstr8
                if str1_8 == "":
                    self.ui.tableView_11.hideRow(row)
                    continue
                else:
                    for keystr in keyList:
                        if keystr in str1_8:
                            self.ui.tableView_11.showRow(row)
                            break
                        else:
                            self.ui.tableView_11.hideRow(row)
        # 如果搜索的是空格，或者 按下了cancel按钮则
        elif not okPressed or text == "":
            # print("Cancel")
            for row in range(rows):
                self.ui.tableView_11.showRow(row)

    # 复制行
    @pyqtSlot()
    def on_QuanBtnCopy_clicked(self):
        # 返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            QMessageBox.critical(self, "错误", "表格行未被选中，无法复制行", QMessageBox.Cancel)
            return
        # print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        self.templist = []
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        rows = len(selectModel.selectedRows())  # 选中了多少行
        strmess = [r+1 for r in rowsIndexList]  # 每个索引+1
        strmess.sort(reverse=False)
        # self.ui.statusBar.showMessage(f"复制成功了{strmess} 行", 5000)
        # print("写入用于复制行的 临时变量 self.templist")
        for row in range(rows):  # 循环了选中的行对象
            for col in range(cols):
                if col == 0:
                    # print("第一列开始，新建一个空列表")
                    self.templist.append([])
                tempItem = model.item(strmess[row]-1, col)
                if (tablename == 12 and col == 5) or (tablename == 2 and col == 6) or \
                        (tablename == 33 and col == 5) or (tablename == 43 and col == 5):
                    # print("有 不计标志 的 单元格")
                    if tempItem.checkState() == Qt.Checked:
                        lineStr = "1"
                    else:
                        lineStr = "0"
                    self.templist[row].append(lineStr)
                else:
                    # print("普通单元格写入！templist")
                    self.templist[row].append(tempItem.text())  # rowslist[row]依次取出行号
        # print(self.templist)
        # 粘贴行 菜单可用
        self.ui.QuanBtnPaste.setEnabled(True)

    # 粘贴行
    @pyqtSlot()
    def on_QuanBtnPaste_clicked(self):
        # print("粘贴行")
        if not self.templist:  # 如果复制列表为空 则不执行
            QMessageBox.critical(self, "错误", "没有复制原数据！", QMessageBox.Cancel)
        # print("临时变量为： ", self.templist)
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            QMessageBox.critical(self, "错误", "表格行未被选中，无法粘贴行", QMessageBox.Cancel)
            return
        # print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        # print("当前取回的表为： ", tablename)
        # print("判断一共多少列")
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        colsdata = len(self.templist[0])
        if cols != colsdata:
            QMessageBox.critical(self, "错误", "复制源与粘贴处 不为同一张表", QMessageBox.Cancel)
            return
        # print("# 开始插入行")
        rows = len(self.templist)
        if rows < 1:
            return
        # 断开所有 链接
        self.disconnectAll()
        for row in range(rows):
            itemlist = []  # QStandardItem 对象列表
            for i in range(cols):
                if (tablename == 12 and i == 5) or (tablename == 2 and i == 6) or \
                        (tablename == 33 and i == 5) or (tablename == 43 and i == 5):
                    # print("#  不计标志 复选框")
                    item = QStandardItem("")
                    item.setFlags(self.__NoCalFlags)
                    item.setCheckable(True)  # 非锁定
                    if self.templist[row][i] == "1":
                        item.setCheckState(Qt.Checked)  # 勾选
                    else:
                        item.setCheckState(Qt.Unchecked)  # 非勾选
                else:
                    temp = self.templist[row][i]
                    item = QStandardItem(temp)
                itemlist.append(item)  # 一行空数据
            model.insertRow(rowsIndexList[-1] + row, itemlist)  # 在最下面行的下面插入一行
        # self.ui.statusBar.showMessage("‘粘贴行’ 成功！", 5000)
        # 变色 粗体 m3变色
        if tablename == 11:
            self.sameCodeName()  # 总表  引用编码重复时红色  辅助
            self.m3ChangeRed(model)
            self.changeColor()  # 总表  《来源》列 变色
            self.quantitiesSum()  # 总表  所有行的 汇总总计
        elif tablename == 12:
            self.itemModel12_itemChanged_All()
            self.quantitiesSum()  # 总表  所有行的 汇总总计
        elif tablename == 2:  # 分类表
            self.m3ChangeRed(model)
        elif tablename == 32:  # 钢筋 汇总量
            self.m3ChangeRed(model)
            self.quantitiesSum32()  # 钢筋表32  所有行的 汇总总计
        elif tablename == 33:  # 钢筋
            self.itemModel33_itemChanged_All()
            self.quantitiesSum32()  # 钢筋表32  所有行的 汇总总计
        elif tablename == 42:  # 土建 汇总量
            self.m3ChangeRed(model)
            self.quantitiesSum42()
        elif tablename == 43:  # 土建
            self.itemModel43_itemChanged_All()
            self.quantitiesSum42()
        # print("# 释放 临时变量")
        self.templist = []
        # 粘贴行 菜单不可用
        self.ui.QuanBtnPaste.setEnabled(False)
        # 恢复所有 链接
        self.connectAll()

    # 删除行
    @pyqtSlot()
    def on_QuanBtnDelRow_clicked(self):
        # 返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            QMessageBox.critical(self, "错误", "没有选择的行 无法删除", QMessageBox.Cancel)
            return
        # 取出返回值
        tablename, tableobj, model, selectModel, rowsIndexList = result
        # print("共计多少行 ", model.rowCount())
        # print("选中了行索引列表 排序后 ", rowsIndexList)
        # 如果只有最后一行也不可删除
        if tablename == 31 or tablename == 41:
            pass
        elif model.rowCount() <= 1:
            QMessageBox.critical(self, "提示", "就一行了，还是别删吧！", QMessageBox.Cancel)
            return
        elif model.rowCount() == len(selectModel.selectedRows()):
            QMessageBox.critical(self, "提示", "全删就没了，留一行吧！", QMessageBox.Cancel)
            return
        # 删除前 警告确认
        res = QMessageBox.warning(self, "警告", "请确认是否需要删除选中的行！", QMessageBox.Yes | QMessageBox.No)
        if QMessageBox.Yes == res:
            pass
            # print("点了确认！")
        elif QMessageBox.No == res:
            # print("点了否！")
            return
        # 断开所有 链接
        self.disconnectAll()
        # 行删除执行
        for row in rowsIndexList:
            model.removeRow(row)
        # 变色 粗体 m3变色
        if tablename == 11:
            self.sameCodeName()  # 总表  引用编码重复时红色  辅助
            self.m3ChangeRed(model)
            self.changeColor()  # 总表  《来源》列 变色
            self.quantitiesSum()  # 总表  所有行的 汇总总计
            self.ui.tableView_12.setEnabled(False)  # 不可用
        elif tablename == 12:
            self.itemModel12_itemChanged_All()
            self.quantitiesSum()  # 总表  所有行的 汇总总计
        elif tablename == 2:  # 分类表
            self.m3ChangeRed(model)
        elif tablename == 32:  # 钢筋 汇总量
            self.m3ChangeRed(model)
            self.quantitiesSum32()  # 钢筋表32  所有行的 汇总总计
            self.ui.tableView_33.setEnabled(False)  # 不可用
        elif tablename == 33:  # 钢筋
            self.itemModel33_itemChanged_All()
            self.quantitiesSum32()  # 钢筋表32  所有行的 汇总总计
        elif tablename == 42:  # 土建 汇总量
            self.m3ChangeRed(model)
            self.quantitiesSum42()
            self.ui.tableView_43.setEnabled(False)  # 不可用
        elif tablename == 43:  # 土建
            self.itemModel43_itemChanged_All()
            self.quantitiesSum42()
        # 恢复所有 链接
        self.connectAll()

    # 插入行
    @pyqtSlot()
    def on_QuanBtnInsertRow_clicked(self):
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            QMessageBox.critical(self, "错误", "没有选择的行 往哪插入呢？", QMessageBox.Cancel)
            return
        # print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        # print("当前取回的表为： ", tablename)
        if tablename == 31 or tablename == 41:
            QMessageBox.critical(self, "错误", "文件表不可以插入行，请用“导入”添加文件！", QMessageBox.Cancel)
            return
        # print("判断一共多少列")
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        # 开始插入行
        rows = len(selectModel.selectedRows())
        # 断开所有 链接
        # self.disconnectAll()
        itemlist = []  # QStandardItem 对象列表
        for i in range(cols):
            item = QStandardItem("")
            #  不计标志 复选框
            if (tablename == 12 and i == 5) or (tablename == 2 and i == 6) or \
                        (tablename == 33 and i == 5) or (tablename == 43 and i == 5):
                item.setFlags(self.__NoCalFlags)
                item.setCheckable(True)  # 非锁定
                item.setCheckState(Qt.Unchecked)  # 非勾选
            itemlist.append(item)  # 一行空数据
        model.insertRow(rowsIndexList[-1], itemlist)  # 在最下面行的下面插入一行

    # 增加行
    @pyqtSlot()
    def on_QuanBtnAddRow_clicked(self, num=1):
        # print("返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制")
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            QMessageBox.critical(self, "错误", "没有选择的行 往哪增加呢？", QMessageBox.Cancel)
            return
        # print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        # print("当前取回的表为： ", tablename)
        if tablename == 31 or tablename == 41:
            QMessageBox.critical(self, "错误", "文件表不可以添加行，请用“导入”添加文件！", QMessageBox.Cancel)
            return
        # print("判断一共多少列")
        if tablename == 31 or tablename == 41:
            return  # 钢筋文件 绘图文件禁止手工添加
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        # 开始在最后增加一行
        # 断开所有 链接
        # self.disconnectAll()
        for x in range(num):
            itemlist = []  # QStandardItem 对象列表
            for i in range(cols):
                item = QStandardItem("")
                #  不计标志 复选框
                if (tablename == 12 and i == 5) or (tablename == 2 and i == 6) or \
                        (tablename == 33 and i == 5) or (tablename == 43 and i == 5):
                    item.setFlags(self.__NoCalFlags)
                    item.setCheckable(True)  # 非锁定
                    item.setCheckState(Qt.Unchecked)  # 非勾选
                itemlist.append(item)  # 一行空数据
            model.appendRow(itemlist)  # 在最下面行的下面插入一行

    # 选择 广联达量表 文件
    @pyqtSlot()
    def on_QuanBtnShapeSelect_clicked(self):
        # print("选择 绘图输入工程量汇总表")
        if self.filename == None:
            curPath = QDir.currentPath()  # 获取系统当前目录
        else:
            curPath = self.filename
        dlgTitle = "选择“绘图输入工程量汇总表”"  # 对话框标题
        filt = "绘图输入工程量汇总表文件(*.xls *.xlsx);;所有文件(*.*)"  # 文件过滤器
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        if not filename:
            return
            # QMessageBox.critical(self, "错误", "无有效报表文件选择！", QMessageBox.Cancel)
            # self.ui.lineEditIron.setText('')
            # self.ui.QuanBtnIronOK.setEnabled(False)  # 钢筋量表确认导入
        if filename:
            self.filename = filename
            self.ui.QuanBtnGlodonImport.setEnabled(True)  # 不可用
            # self.ui.lineEditIron.setText(filename)
            # self.ui.QuanBtnIronOK.setEnabled(True)  # 钢筋量表确认导入
            answer = QMessageBox.information(self, "温馨提示", "导入《绘图输入工程量汇总表》！", QMessageBox.Ok | QMessageBox.Cancel)
            if answer == QMessageBox.Ok:
                # print("OK")
                # self.__filename = filename
                # self.ui.label_path.setText(f'导入文件路径：{filename}')
                # FieldNameCol = {'楼层名称': 0, '构件大类': 1, '构件小类': 2, '构件名称': 3, '钢筋等级': 4,
                #                 '钢筋直径': 5, '接头类型': 6, '总重(kg)': 7, '其中箍筋(kg)': 8, '接头个数': 9}
                dataDict = {}
                with xlrd.open_workbook(filename) as f:
                    sheetsName = f.sheet_names()  # 获取所有sheet名字
                    # if "楼层构件类型统计汇总表" not in sheetsName:
                    #     QMessageBox.critical(self, "错误", "未找到匹配的钢筋量表！\n请选择软件导出的原始表格，勿做任何修改！",
                    #                          QMessageBox.Cancel)
                    #     return
                    # print(sheetsName)
                    shaperows = 0  # 总行数
                    sheetcount = 0  # 总表数
                    for sheetName in sheetsName:
                        if not "工程量汇总表" in sheetName:  # 如果表不是原始表 则跳过
                            continue
                        sheetObj = f.sheet_by_name(sheetName)  # 得到表格对象
                        datasheet = sheetObj._cell_values  # 得到二维表数据
                        shaperows += len(datasheet)
                        sheetcount += 1
                        sheetName = sheetName[sheetName.rfind("-") + 1:].strip()  # 去掉每张表名的前缀
                        dataDict[sheetName] = datasheet
                if len(dataDict) < 1:  # 如果没有符合要求的表 则退出
                    QMessageBox.critical(self, "错误", "未找到匹配的量表！\n请选择软件导出的原始表格，勿做任何修改！",
                                         QMessageBox.Cancel)
                    return
                self.itemModel41.itemChanged.disconnect(self.itemModel41_itemChanged)
                # 文件序号
                id = self.itemModel41.rowCount()
                if id == 0:
                    id = "1"
                else:
                    id = str(int(self.itemModel41.item(id - 1, 0).text()) + 1)
                listitem = []
                listitem.append(QStandardItem(id))
                listitem.append(QStandardItem(str(filename)))
                listitem.append(QStandardItem(str(f'{sheetcount} 张表，{shaperows} 行')))
                listitem.append(QStandardItem(str(dataDict)))
                self.itemModel41.appendRow(listitem)
                # self.quantitiesDict["钢筋工程量表"] = datasheet[1:]  # 第一行为字段名
                QMessageBox.information(self, "温馨提示",
                                        f"成功导入数据！共计：{sheetcount} 张表，共计{shaperows} 行", QMessageBox.Ok)
                self.itemModel41.itemChanged.connect(self.itemModel41_itemChanged)
            else:
                QMessageBox.critical(self, "提示", "想好再导，不要乱点！", QMessageBox.Cancel)
                # self.ui.lineEditIron.setText('')
                # self.ui.QuanBtnIronOK.setEnabled(False)  # 钢筋量表发送

    # 修改模板  打开 广联达量表 匹配表 子窗口
    @pyqtSlot()
    def on_QuanBtnImportGlodon_set_clicked(self):
        # print("修改模板")
        dlgTableMatching = QmyDialogImportGlodon()  # 局部变量，构建时不能传递self
        # dlgTableMatching.setIniSize(self.itemModel.rowCount(),
        #                         self.itemModel.columnCount())
        ret = dlgTableMatching.exec()  # 模态方式运行对话框
        if ret == QDialog.Accepted:
            # print('确定按钮被按下')
            # gcllist = dlgTableMatching.getTableSize()  # 得到提取后的二维列表 matching
            resultdict = dlgTableMatching.getmatching()  # 得到提取后的二维列表 matching
            self.tabWidget_1_currentChanged(3)  # 刷新下字典文件 下拉列表

    # 提取 广联达量表 土建
    @pyqtSlot()
    def on_QuanBtnGlodonImport_clicked(self):
        # print("QuanBtnGlodonImport  提取")
        # 是否区分地上地下
        YorNUpDown = self.matchingdict['其他']['区分地上地下']
        # 不提量关键字
        Nokeys = self.matchingdict['其他']['不提量关键字']  # 列表类型
        # 获取匹配表
        str1 = self.ui.comboBox.currentText()   # 获取当前选项的内容
        matchingList = self.matchingdict[str1]  # 获取匹配表

        rows4 = self.itemModel41.rowCount()  # 文件表的总行数
        if rows4 < 1:
            QMessageBox.critical(self, "错误", "请先导入报表文件，可以追加多个文件！", QMessageBox.Cancel)
            return
        totalList = []  # 存放提取后的二维表数据
        for r in range(rows4):  # 依次遍历 文件表
            id = self.itemModel41.item(r, 0).text()  # 文件序号
            datadictstr = self.itemModel41.item(r, 3).text()  # str格式dict
            if not datadictstr:
                QMessageBox.critical(self, "错误", "请先导入报表文件，可以追加多个文件！", QMessageBox.Cancel)
                return
            wbdict = ast.literal_eval(datadictstr)
            for sht, data2 in wbdict.items():
                # print("表名 ： ", sht)
                for x in range(len(matchingList)):
                    if x % 2:  # 偶数行跳过
                        continue
                    if sht == matchingList[x][0]:  # 表名和匹配表内的工程量表名相同
                        matchingListevery = matchingList[x:x + 2]  # 取出匹配的 两行列表  工程量&简称
                        # 处理每一张匹配的二维表
                        list2 = self.extractEverySheet(sht, data2, matchingListevery, YorNUpDown, Nokeys, id)
                        if list2:
                            totalList.extend(list2)  # 每个表提取的数据datalist 扩展到totalList
            # pprint(totalList)
        if len(totalList) < 1:
            QMessageBox.critical(self, "错误", "导入的报表不符合要求，请用原始导出表，不要做任何修改！", QMessageBox.Cancel)
            return
        # 开始把提取的量表导入 清单表
        if len(totalList) > 0:
            # print('开始把提取的量表导入 清单表')
            # 断开链接
            self.itemModel42.itemChanged.disconnect(self.itemModel42_itemChanged)
            zrow = self.itemModel42.rowCount()  # 主表的最大行
            # self.on_actAppRow_triggered(len(gcllist))  # 先增加多行
            for row in range(len(totalList) - 1):  # 去除最后一行
                itemlist = []
                for col in range(len(totalList[0])):
                    temp = totalList[row][col]
                    item = QStandardItem(str(temp))
                    itemlist.append(item)
                self.itemModel42.appendRow(itemlist)
            # ~~~~~~最后做的几项收尾工作~~~~~~~~~
            # self.changeColor()  # 《来源》列 变色
            # self.m3ChangeRed(self.itemModel42)  # m3 变色
            # self.quantitiesSum()  # 所有行的 汇总总计
            # 恢复链接
            self.itemModel42.itemChanged.connect(self.itemModel42_itemChanged)
            # self.ui.statusBar.showMessage("提取广联达工程量成功！", 5000)
            QMessageBox.information(self, "成功导入:", f"{len(totalList)} 行,广联达工程量！\n 开始计算！", QMessageBox.Ok)
        else:
            QMessageBox.information(self, "失败！:", "请使用原始软件导出的 绘图工程量表，勿做任何改动！", QMessageBox.Ok)
        # 刷新一下 总计
        self.quantitiesSum42()  # 钢筋表所有行的 汇总总计
        self.m3ChangeRed(self.itemModel42)

        self.ui.QuanBtnGlodonOK.setEnabled(True)  # 土建量表 发送至清单总表
        self.ui.QuanBtnGlodonImport.setEnabled(False)  # 不可用

    # 处理每一张找到的工程量表  辅助  get 方法
    # 参数依次为： 表对象，匹配字，地上地下，不提量
    def extractEverySheet(self, shtName, itemsList, matchingListevery, YorNUpDown, Nokeys, id):  # matchingListevery 两行多列表
        # print(Nokeys)
        # pinyinbianma = Pinyin()   # 中文转拼音实例化
        vaNo = 0  # 清单行的序号
        itemsRows = len(itemsList)
        itemsCols = len(itemsList[0])
        # ~~~~~~~~~找到关键字的行号和列号~~~~~~~~~~~~~~0
        keydict = {x: [] for x in ["楼层", "材质", "混凝土强度等级", "砼标号", "名称", "工程量名称", ]}  # keydict["工程量名称"]=[开始列,结束列]
        for col in range(len(itemsList[0])):  # 在第一行循环
            item = itemsList[0][col]
            if item in keydict.keys():
                if keydict.get(item):  # 说明是第二次匹配到
                    keydict[item][1] = col  # [小列号，大列号]
                else:
                    keydict[item] = [col, col]
        keydict["首层"] = -1  # 没有首层 则匹配-1行
        # print(shtName, keydict)
        if not keydict.get('楼层') or len(keydict["楼层"]) < 1:  # TODO 没打开楼层 则退出
            # if len(keydict["楼层"]) < 1:  # TODO 没打开首层 则退出
            return
        col = keydict["楼层"][0]
        # print(col)
        for row in range(2, itemsRows):
            onefloorName = itemsList[row][col]
            strkey = re.search(r"\([^\(]*\)", onefloorName)  # # 多轴网文件 (楼层)   用()内的楼层名
            if strkey:  # 没有匹配正则 则退出while
                onefloorName = strkey.group()[1:-1]  # 脱去两侧的括号

            if onefloorName in ["首层", "第1层", "第一层"]:  # TODO 首层的楼层名字
                keydict["首层"] = row
                break
        # ~~~~~~~~~找到定位关键字的行号和列号~~~~~~~~~~~~~~1
        # print(shtName, keydict)
        # print('matchingListevery', matchingListevery)
        # ~~~~~~~~~工程量名称 简称 匹配~~~~~~~~~~~~~~0
        namedict = {}  # namedict {'体积(m3)': '主体', '模板面积(m2)': '模板'}
        matchingCols = len(matchingListevery[0])  # matchingListevery 两行 列表
        for col in range(1, matchingCols):  # 依次浏览每一个需要提量的原始的工程量名
            item = matchingListevery[0][col]
            if item:
                namedict[item] = matchingListevery[1][col]
        # ~~~~~~~~~工程量名称 简称 匹配~~~~~~~~~~~~~~1
        # print('namedict', namedict)
        # ~~~~~~~~~工程量名称 列号 ~~~~~~~~~~~~~~0
        for i in namedict.keys():
            for col in range(keydict["工程量名称"][0], keydict["工程量名称"][1] + 1):  # 在第二行循环
                item = itemsList[1][col]
                if item in namedict.keys():
                    keydict[item] = [col, col]
                    continue  # 改break 提速
        # ~~~~~~~~~工程量名称 列号 ~~~~~~~~~~~~~~1
        # print("shtName, keydict", shtName, keydict)
        # ~~~~~~~~~循环量表的每一行~~~~~~~~~~~~~~0
        rowsList = []
        for old, new in namedict.items():  # ole 原名，new 简称
            # print(f"~~~~~~~~~~开始循环匹配《{shtName}》表的工程量：{old},简称：{new}")
            if len(keydict["楼层"]) <= 1:  # TODO 楼层标记未打开,所在列未识别到 跳过 后期做消息框提醒
                continue
            floorname = ''  # 楼层名字初始化
            str3 = ''  # 特征 初始化
            str4 = old[old.rfind("(") + 1:-1]  # 计量单位
            str5 = ''  # 明细表内的计算表达式 初始化
            str6 = 0  # 明细表计算结果,楼层工程量累加 初始化
            sumnum = 0  # 清单表内的工程量初始化
            detailedlist = []  # 明细表 初始化
            valuelist = []  # 清单表初始化

            for row in range(2, itemsRows - 1):  # 跳过表头和 最后的 “ 合计 ” 行
                # print(f"~~~~~~~~~~~~~~表《{shtName}》-{old}-的量，第{row + 1}行开始运行！~~~~~~~~~~~~~~")
                if keydict.get(old):  # 对应的工程量简称 没有导出 则跳过
                    num = itemsList[row][keydict[old][0]]  # TODO 工程量  str5 做累加处理
                else:
                    break
                strxiaojiheji = itemsList[row][keydict["工程量名称"][0] - 1]
                if strxiaojiheji == "小计" or strxiaojiheji == "合计":  # TODO 是小计 则跳过
                    continue
                # 构件名称未打开,则不过滤
                if len(keydict["名称"]) > 1:
                    gjname = itemsList[row][keydict["名称"][0]]  # 如果构件名称内含有不计量关键字 则跳过
                    flag = 'ok'
                    for strgj in Nokeys:  # 在“不计标志” 关键字内 则跳过
                        # print(strgj)
                        if strgj in gjname:
                            flag = 'Next'
                            break
                    if flag == 'Next':
                        continue
                else:
                    return  # TODO 2022/01/08
                # ~~~~~~~~~判断首层所在 区分地上地下~~~~~~~~~~~~~~0
                if len(keydict["楼层"]) < 1:  # 楼层名称未打开 则跳过  疑似无用
                    continue
                floorname = itemsList[row][keydict["楼层"][0]]
                strkey = re.search(r"\([^\(]*\)", floorname)  # # 多轴网文件 (楼层)   用()内的楼层名
                if strkey:  # 没有匹配正则 则退出while
                    floorname = strkey.group()[1:-1]  # 脱去两侧的括号

                if YorNUpDown == "不区分":
                    YorNUpDownName = ""
                else:  # 要 "区分" ,判断是否找的到
                    # if keydict["首层"] == -1:  # 说明要区分 又没找到首层 判断 是否为地下
                    undergroundKeys = ['基', '地下', '负', '第-']  # TODO 判断楼层是否为地下的关键字
                    for underKey in undergroundKeys:
                        if underKey in floorname:
                            YorNUpDownName = '地下_'
                            break
                        else:
                            YorNUpDownName = '地上_'
                str2 = YorNUpDownName
                # ~~~~~~~~~~~~~~特征描述~~~~~~~~~~~~~~~~0
                if str4 == 'm3':
                    if len(keydict["混凝土强度等级"]):  # TODO 特征描述 混凝土强度等级 砖种类
                        str3 = itemsList[row][keydict["混凝土强度等级"][0]]
                    elif len(keydict["砼标号"]):
                        str3 = itemsList[row][keydict["砼标号"][0]]
                    elif len(keydict["材质"]):
                        str3 = itemsList[row][keydict["材质"][0]]
                    else:
                        str3 = ''
                # 建筑做法等提量 按“名称” 来分类的构件
                elif shtName in ['楼地面', '墙面', '天棚', '踢脚', '墙裙', '吊顶', '独立柱装修',
                                 '单梁装修', '屋面', '台阶', '散水']:
                    str3 = itemsList[row][keydict["名称"][0]]
                    str3 = re.sub(r'\[[^\[]*\]', '', str3)  # 【注释】 删除
                if str3:
                    str3 = "_" + str3
                if str5 == '':  # 第一次执行这里
                    str5 = num
                    str6 = float(str5)
                    sumnum = float(str5)
                else:  # TODO 楼层、特征、项目名称（地上地下+简称）和上次不同时,计算表达式 归 "" str3 特征
                    if floorname != detailedlist[-1][0] and (str2 + new + str3) == valuelist[-1][4]:  # 特征 项目名称 同清单 换楼层
                        detailedlist.append([])  # 明细增加一行
                        # print('明细行增加了一行')
                        str5 = num  # str5 = '' 楼层名和上一次不同 初始化 表达式
                        str6 = float(num)
                        sumnum += float(num)
                    elif (str2 + new  + str3) != valuelist[-1][4]:  # 特征 项目名称不同 换清单
                        detailedlist = []  # 明细初始化
                        # print('明细初始化')
                        valuelist.append([])  # 清单表增加一行
                        vaNo += 1
                        # print('清单表增加了一行')
                        str5 = num  # str5 = '' 楼层名和上一次不同 初始化 表达式
                        str6 = float(num)
                        sumnum = float(num)
                    else:  # 名称、特征、楼层都和上一行一致
                        str5 = str5 + '+' + num
                        str6 += eval(num)
                        sumnum += eval(num)
                # print('2str5', type(str5), str5)
                # print('2str6', str6)
                # print('detailedlist, valuelist', detailedlist, valuelist)
                # str6 = 111
                if len(detailedlist) < 1:  # 如果有明细表 则在最后一行 添加
                    detailedlist.append([])
                detailedlist[-1] = [floorname, str5, str(round(str6, 4)), str(round(str6, 4)),
                                    str(round(str6, 4)), '', '', '']  # 按明细结构 的二维表
                # sumnum  str6
                # 编码
                # str1_1 = pinyinbianma.get_initials(str2 + new, '').lower()  # 首拼音 转小写
                # str1_2 = pinyinbianma.get_pinyin(shtName + '_' + str3, '')  # 全拼音
                # str1 = str1_1 + '_' + str1_2  # 名称首拼音+特征描述全拼音 的编码
                if len(valuelist) < 1:
                    valuelist.append([])
                    vaNo += 1

                # valuelist[-1] = [vaNo, '', str2 + new, str3, str4, str(round(sumnum,6)), '', detailedlist, '', '',
                #                  '广联达绘图工程量_' + shtName]  # 按清单表结构 新建
                valuelist[-1] = [vaNo, '文件序号_' + id, '', shtName, str2 + new + str3, str4,
                                 str(round(sumnum, 4)), detailedlist, '']  # 按清单表结构 新建
                # print('detailedlist, valuelist', detailedlist, valuelist)
            # print(valuelist)
            rowsList.extend(valuelist)  # 每个类型的工程量 增加至总表 二维表
        # ~~~~~~~~~循环量表的每一行~~~~~~~~~~~~~~1
        # print(rowsList)
        # 最后一行添加空行
        rowsList.extend([['', '', '', '', '', '', '', [['', '', '', '', '', '', '', '']], '']])

        # print(rowsList)
        return rowsList

    # 绘图工程量表 发送至 工程量总表
    @pyqtSlot()
    def on_QuanBtnGlodonOK_clicked(self):
        # print("绘图工程量表 发送至 工程量总表")
        name = "绘图工程量表"
        result = self.ironGlodonSendMaintable(self.itemModel42, name)
        # 整理后的字典（钢筋 土建分表） 发至总表
        self.subtable_to_main(result)
        # 写入前 先切换页签 和标志
        self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 页面
        self.tabindex = "工程量总表"  # 清单表

    # 导入 分类表
    @pyqtSlot()
    def on_QuanBtnImportClassification_clicked(self):
        # print("导入 分类表")
        dlgTableFExcel = QmyDialogFExcel()  # 局部变量，构建时不能传递self
        ret = dlgTableFExcel.exec()  # 模态方式运行对话框
        if (ret == QDialog.Accepted):
            # 断开连接
            self.itemModel2.itemChanged.disconnect(self.itemModel2_itemChanged)
            detailedData, datarow = dlgTableFExcel.getData()  # 返回匹配好的 二维列表 需要转置
            if not detailedData or not datarow:
                return
            zrow = self.itemModel2.rowCount()  # 主表的最大行
            self.on_QuanBtnAddRow_clicked(datarow)  # 先增加多行
            # self.itemModel2.setRowCount(zrow+datarow)
            cols = self.itemModel2.columnCount()
            for col in range(len(detailedData)):
                if not len(detailedData[col]):  # 如果空值 跳过
                    continue
                for row in range(len(detailedData[col])):
                    if detailedData[col][row] == '':  # 跳过空列
                        continue
                    if col == 6:  # 不为数值型字符 则赋值为0
                        try:
                            temp = detailedData[col][row]
                            item = QStandardItem(str(float(temp)))
                            # item = QStandardItem(temp)
                        except:
                            temp = ""  # 数值型内输入了 其他字符 闪退的情况
                            item = QStandardItem(temp)
                        finally:
                            self.itemModel2.setItem(zrow + row, col, item)
                            continue
                    temp = detailedData[col][row]
                    item = QStandardItem(temp)
                    self.itemModel2.setItem(zrow + row, col, item)
            # self.ui.statusBar.showMessage(f"成功导入:{datarow}行分类表文件！", 5000)
            QMessageBox.information(self, "成功导入:", f"{datarow} 行分类表文件！\n 开始计算！", QMessageBox.Ok)
            # 刷新所有行的计算工程量
            rows = self.itemModel2.rowCount()
            for row in range(zrow, rows):
                strtemp = self.itemModel2.item(row, 4).text().strip()  # 获取字符
                if strtemp == "":  # 表达式为空时  合价为空
                    item = QStandardItem("")
                    self.itemModel2.setItem(row, 5, item)
                    item = QStandardItem("")
                    self.itemModel2.setItem(row, 7, item)
                else:  # 表达式的有内容进一步处理
                    # print("# 注释 中文括号 工程量表 处理，")
                    # ~~~~~~~~~~~~~~~~~~【注释】 中文圆括号 处理~~~~~~~~~~~~~~~~~~~~~
                    strtemp = re.sub(r'\u3010[^\u3010]*\u3011', '', strtemp)  # 【注释】 删除
                    strtemp = re.sub('（', '(', strtemp)  # 中文圆括号替换英文 删除
                    strtemp = re.sub('）', ')', strtemp)  # 中文圆括号替换英文 删除
                    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    try:  # 传入的非表达式  无法转换为eval则出错
                        strtempstr = eval(strtemp)  # 表达式 转结果
                        strtempfloat = round(strtempstr, 3)  # 保留三位小数
                        strtempfin = str(strtempfloat)  # 转成字符
                        item = QStandardItem(strtempfin)  # 需要给一个str格式
                        self.itemModel2.setItem(row, 5, item)  # 能转化为结果 则直接赋值 给 计算结果列
                        item = QStandardItem("")
                        self.itemModel2.setItem(row, 7, item)
                    except:
                        item = QStandardItem("")
                        self.itemModel2.setItem(row, 5, item)
                        item = QStandardItem('●')
                        self.itemModel2.setItem(row, 7, item)
            # print('写入完成开始m3')
            self.m3ChangeRed(self.itemModel2)  # m3 变色
            self.ui.QuanBtnImportClassificationToQuan.setEnabled(True)  # ← 发送至清单
            # 恢复连接
            self.itemModel2.itemChanged.connect(self.itemModel2_itemChanged)

    # 分类表 发送至工程量清单
    @pyqtSlot()
    def on_QuanBtnImportClassificationToQuan_clicked(self):
        # print("分类表发送至工程量清单")
        # 断开链接
        self.itemModel11.itemChanged.disconnect(self.itemModel11_itemChanged)
        # self.itemModel2.itemChanged.disconnect(self.itemModel2_itemChanged)
        # self.itemModel.itemChanged.disconnect(self.itemChanged)   # 用于undo redo
        # 开始转入清单表
        rowim1 = self.itemModel11.rowCount()  # 清单表的总行数
        rowim3 = self.itemModel2.rowCount()  # 分类表的总行数
        dictim3 = {}  # 空字典用于存放 最终写入主表的数据  工程量名称为key  清单条目为value(一维复合表),条目中的第7列 为明细
        # print('开始循环每一行分类表')
        flag = None
        for row in range(rowim3):  # 循环分表每一行
            str1 = self.itemModel2.item(row, 1).text().strip()  # 部位/楼层
            if flag == None and str1 == "":  # 跳过前面几行空行
                continue
            elif str1 != "":
                floorStr = str1
                flag = 'OK'
            else:
                pass
            str2 = self.itemModel2.item(row, 2).text().strip()  # 项目名称
            # 以下2行 暂时注释
            # str2 = re.sub('（', '(', str2)  # 中文圆括号替换英文 删除
            # str2 = re.sub('）', ')', str2)  # 中文圆括号替换英文 删除

            str4 = self.itemModel2.item(row, 4).text().strip()  # 表达式
            if str2 == "" and str4 == "":  # 如果工程量名称和表达式不存在，则下一row
                continue
            str3 = self.itemModel2.item(row, 3).text().strip()  # 计量单位
            tempkey = str2 + str3   # 工程量名称+计量单位+单价  str 形式  有一项不同都判断是不一样的清单子目
            # print('tempkey', tempkey)  # 辅助
            str5 = self.itemModel2.item(row, 5).text()  # 计算结果
            # print('1')
            if not str5:
                str5f = 0
                str5 = ""
            else:
                str5f = round(float(str5), 3)
            # print('2')
            item6 = self.itemModel2.item(row, 6)  # 不计标志
            if item6.checkState() == Qt.Checked:
                str6 = '1'
            else:
                str6 = ''
            str7 = self.itemModel2.item(row, 7).text()  # 错误
            str8 = self.itemModel2.item(row, 8).text()  # 备注
            # 以字典形式 写入
            detailedlist = [floorStr, str4, str5, str5, str5, str6, str7, str8]  # 按明细结构 的二维表
            result = dictim3.get(tempkey)
            if result == None:  # 工程名名称不存在 第一次新建value
                # dictim3[tempkey + "工程量"] = 0  # 先创建一个工程量key 用于工程量的汇总
                # detailedlist = [floorStr, str4, str5, str5, str5, str6, str7, str8]  # 按明细结构 的二维表
                valuelist = ['', '分类表', 'FLB_' + str2 + '_' + str3, str2, '', str3, str5f, [], '']  # 按清单表结构 新建
                valuelist[7].append(detailedlist)  # 把明细表 新增到value内
                dictim3[tempkey] = valuelist  # 把初始化的表格写入字典
            else:  # 已存在此key
                tempValue = dictim3[tempkey]  # 取出对应清单结构的行数据（列表）
                tempValue[7].append(detailedlist)  # 把明细表 新增到value内
                tempValue[6] += str5f

        # print('分类表循环完成',dictim3)
        addrows = len(dictim3.keys())
        if addrows < 1:
            # 恢复链接
            self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged)
            return

        # pprint(f'写入dictim3 数据 \n {dictim3}')
        num = 1
        sortstr = [x for x in dictim3.keys()]
        sortstr.sort()
        for name in sortstr:  # 循环每一条清单+单位组合名称
            itemlist = []
            for col in range(len(dictim3[name])):
                if col == 0:  # 第一个序号
                    item = QStandardItem(str(num))
                elif col == 6:
                    item = QStandardItem(str(round(dictim3[name][col], 4)))
                else:
                    item = QStandardItem(str(dictim3[name][col]))
                itemlist.append(item)
            num += 1
            self.itemModel11.appendRow(itemlist)

        # print('字典循环完成')
        # ~~~~~~最后做的几项收尾工作~~~~~~~~~
        self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 页面
        self.tabindex = "工程量总表"  # 清单表
        self.changeColor()   # 《来源》列 变色
        self.m3ChangeRed(self.itemModel11)  # m3 变色
        self.quantitiesSum()  # 所有行的 汇总总计
        self.ui.QuanBtnImportClassificationToQuan.setEnabled(False)   # 发送按钮不可用
        # 恢复链接
        self.itemModel11.itemChanged.connect(self.itemModel11_itemChanged)

    # 选择 钢筋表
    @pyqtSlot()
    def on_QuanBtnIronSelect_clicked(self):
        # print("选择 钢筋表")
        if self.filename == None:
            curPath = QDir.currentPath()  # 获取系统当前目录
        else:
            curPath = self.filename
        dlgTitle = "选择钢筋工程量表文件"  # 对话框标题
        filt = "钢筋工程量表文件(*.xls *.xlsx);;所有文件(*.*)"  # 文件过滤器
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        if not filename:
            pass
            # QMessageBox.critical(self, "错误", "无有效报表文件选择！", QMessageBox.Cancel)
            # self.ui.lineEditIron.setText('')
            # self.ui.QuanBtnIronOK.setEnabled(False)  # 钢筋量表确认导入
        if filename:
            # self.ui.lineEditIron.setText(filename)
            # self.ui.QuanBtnIronOK.setEnabled(True)  # 钢筋量表确认导入
            self.filename = filename
            answer = QMessageBox.information(self, "温馨提示", "导入钢筋量表！", QMessageBox.Ok | QMessageBox.Cancel)
            if answer == QMessageBox.Ok:
                # print("OK")
                # self.__filename = filename
                # self.ui.label_path.setText(f'导入文件路径：{filename}')
                FieldNameCol = {'楼层名称': 0, '构件大类': 1, '构件小类': 2, '构件名称': 3, '钢筋等级': 4,
                                '钢筋直径': 5, '接头类型': 6, '总重(kg)': 7, '其中箍筋(kg)': 8, '接头个数': 9}
                with xlrd.open_workbook(filename) as f:
                    sheetsName = f.sheet_names()  # 获取所有sheet名字
                    if "楼层构件类型统计汇总表" not in sheetsName:
                        QMessageBox.critical(self, "错误", "未找到匹配的钢筋量表！\n请选择软件导出的原始表格，勿做任何修改！",
                                             QMessageBox.Cancel)
                        return
                    # print(sheetsName)
                    for sheetName in sheetsName:
                        if "楼层构件类型统计汇总表" in sheetName:
                            sheetObj = f.sheet_by_name(sheetName)  # 得到表格对象
                            datasheet = sheetObj._cell_values  # 得到二维表数据
                            break
                rows = len(datasheet)
                cols = len(datasheet[0])
                # print(rows, cols)
                # 获得真正的字段顺序
                realdatasheet = []
                for key, val in FieldNameCol.items():
                    if datasheet[0][val] == key:
                        continue
                    else:
                        # print("数据格式有误")
                        QMessageBox.critical(self, "错误", "请检查表头文字是否齐全，勿做任何修改！",
                                             QMessageBox.Cancel)
                        return
                # 二维表导入 数据模型
                self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)
                id = self.itemModel31.rowCount()
                if id == 0:
                    id = "1"
                else:
                    id = str(int(self.itemModel31.item(id - 1, 0).text()) + 1)
                listitem = []
                listitem.append(QStandardItem(id))
                listitem.append(QStandardItem(str(filename)))
                listitem.append(QStandardItem(str(f'{rows} 行')))
                listitem.append(QStandardItem(str(datasheet)))
                self.itemModel31.appendRow(listitem)
                # self.quantitiesDict["钢筋工程量表"] = datasheet[1:]  # 第一行为字段名
                QMessageBox.information(self, "温馨提示", f"成功导入数据！共计：{rows} 行", QMessageBox.Ok)
                self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged)
                self.ui.QuanBtnIronImport.setEnabled(True)
            else:
                print("NOOK")
                # self.ui.lineEditIron.setText('')
                # self.ui.QuanBtnIronOK.setEnabled(False)  # 钢筋量表发送

    # 钢筋工程量表 提取钢筋
    @pyqtSlot()
    def on_QuanBtnIronImport_clicked(self):
        # print("钢筋工程量表 提取")
        rows3 = self.itemModel31.rowCount()  # 文件表的总行数
        if rows3 < 1:
            QMessageBox.critical(self, "错误", "请先导入报表文件，可以追加多个文件！", QMessageBox.Cancel)
            return

        # 取报表 关键字初始化
        field_name_list = self.text.text().split('；')  # 获取下拉复选框 选中的内容！ 转成列表形式
        # print(field_name_list)
        lowerKeysList = ["基", "地下", "负", "第-"]   # 楼层名称 地下 关键字
        meshKeysList = ["网", "刚防", "钢防"]    # 钢筋网片 关键字
        masonryIronKeysList = ["砌体", "拉结筋", "钢防"]    # 判断砌体加筋、板缝筋 关键字
        valuelists = []  # 存放写入总表的二维表
        blankrow = ['' for _ in range(9)]  # 定义一个空行
        for r in range(rows3):  # 依次遍历 文件表
            # 钢筋工程量表 初始化时 也需定义一样的字典
            self.dictIron3 = {'级别总重': {},  # 钢筋种类的总重量 如 一级钢A 500t 二级钢B 600t
                              '接头个数': {},  # 除绑扎以外的接头总数 如 电渣压力焊 100个
                              '接头直径个数': {},  # 一级钢A6 500t
                              '构件总重': {},
                              '构件楼层总重': {},
                              '楼层构件总重': {},
                              '级别直径总重': {},
                              '构件纵筋总重': {},
                              '江苏2014_钢筋重量': {},
                              '江苏2014_接头个数': {}
                              }
            datastr = self.itemModel31.item(r, 3).text()  #str格式二维list
            if not datastr:
                QMessageBox.critical(self, "错误", "请先导入报表文件，可以追加多个文件！", QMessageBox.Cancel)
                return
            datas = ast.literal_eval(datastr)
            # 空字典用于存放 最终写入主表的数据  工程量名称为key  清单条目为value(一维复合表),条目中的第7列 为明细
            # totalKg = 0
            # totalConnect = 0
            id = self.itemModel31.item(r, 0).text()
            rows = len(datas)
            for row in range(1, rows):  # 遍历钢筋表的每一行
                # print("第 ", row, " 行")
                # 以此提取每一列数据
                floorName = datas[row][0]  # 楼层名称
                partsBig = datas[row][1]  # 构件大类
                partsSmall = datas[row][2]  # 构件小类
                partsName = datas[row][3]  # 构件名称
                ironLevel = datas[row][4] + "级"  # 钢筋等级
                ironD = datas[row][5]  # 钢筋直径
                connectType = datas[row][6]  # 接头类型
                kg = datas[row][7]  # 总重kg
                hoopKg = datas[row][8]  # 箍筋kg
                connectNum = datas[row][9]  # 接头个数
                # 判断地上、下  多轴网 多施工段 需考虑
                floorName = re.sub(r"\([^\(]*\)", '', floorName)  # 多轴网文件 (轴网) 删除
                lowUppstr = "地上"  # 默认都为地上
                if not floorName:
                    continue
                # 如果楼层名称在 地下楼层关键字列表内 则判断为 地下
                for lowkey in lowerKeysList:
                    if lowkey in floorName:
                        lowUppstr = "地下"
                        break
                # 判断砌体加筋、板缝筋
                masonryIron = False
                for makey in masonryIronKeysList:
                    if partsSmall in makey:
                        masonryIron = True  # 是砌体加筋
                # 判断是否冷拔丝
                # 判断是否为网片
                meshflag = False  # 判断是否为网片
                drawingIron = False  # 判断是否冷拔丝
                if int(float(ironD)) < 6:  # 小于6圆都属于网片
                    meshflag = True
                    drawingIron = True
                else:  # 大于等于6圆 公司计算口径也可能认定为网片
                    for meshkey in meshKeysList:
                        if meshkey in partsName and int(float(ironD)) < 9:  # 如果名字带有关键字 且直径 < 9 圆 判定为网片
                            meshflag = True
                            break
                # 处理 重量 直径
                ironD = str(ironD)+"圆"
                ironDInt = float(ironD[:-1])  # 判断直径大小用
                kg = round(float(kg), 3)
                hoopKg = round(float(hoopKg), 3)
                connectNum = int(connectNum)
                # 各种样式的汇总数据 总量 接头 按级别 数量
                # ~~~~~~开始分类~~~~~~
                # 级别总重
                if meshflag == False:  # 非网片钢筋
                    qdm = lowUppstr + "_" + ironLevel  # qdm 清单名称 一级子目
                    result = self.dictIron3['级别总重'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['级别总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['级别总重'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['级别总重'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['级别总重'][qdm].get(ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['级别总重'][qdm][ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['级别总重'][qdm][ironD][0] += kg  # 拼接表达式 放弃 太长
                        self.dictIron3['级别总重'][qdm][ironD][1] += kg  # 结果real累加
                        self.dictIron3['级别总重'][qdm][ironD][2] += kg  # 计算结果累加
                        self.dictIron3['级别总重'][qdm][ironD][3] += kg  # 楼层累加
                else:  # 是网片钢筋
                    qdm = lowUppstr + "_钢筋网片"  # qdm 清单名称 一级子目
                    result = self.dictIron3['级别总重'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['级别总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['级别总重'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['级别总重'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['级别总重'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = A6”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['级别总重'][qdm][ironLevel + ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['级别总重'][qdm][ironLevel + ironD][0] += kg  #  表达式
                        self.dictIron3['级别总重'][qdm][ironLevel + ironD][1] += kg  # 结果real累加
                        self.dictIron3['级别总重'][qdm][ironLevel + ironD][2] += kg  # 计算结果累加
                        self.dictIron3['级别总重'][qdm][ironLevel + ironD][3] += kg  # 楼层累加
                # 级别直径总重
                if meshflag == False:  # 非网片钢筋
                    qdm = lowUppstr + "_" + ironLevel + "_" + ironD  # qdm 清单名称 一级子目
                    result = self.dictIron3['级别直径总重'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['级别直径总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['级别直径总重'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['级别直径总重'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['级别直径总重'][qdm].get(ironLevel + "_" + ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['级别直径总重'][qdm][ironLevel + "_" + ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['级别直径总重'][qdm][ironLevel + "_" + ironD][0] += kg  # 拼接表达式 放弃 太长
                        self.dictIron3['级别直径总重'][qdm][ironLevel + "_" + ironD][1] += kg  # 结果real累加
                        self.dictIron3['级别直径总重'][qdm][ironLevel + "_" + ironD][2] += kg  # 计算结果累加
                        self.dictIron3['级别直径总重'][qdm][ironLevel + "_" + ironD][3] += kg  # 楼层累加
                else:  # 是网片钢筋
                    qdm = lowUppstr + "_钢筋网片"  # qdm 清单名称 一级子目
                    result = self.dictIron3['级别直径总重'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['级别直径总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['级别直径总重'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['级别直径总重'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['级别直径总重'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = A6”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['级别直径总重'][qdm][ironLevel + ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['级别直径总重'][qdm][ironLevel + ironD][0] += kg  #  表达式
                        self.dictIron3['级别直径总重'][qdm][ironLevel + ironD][1] += kg  # 结果real累加
                        self.dictIron3['级别直径总重'][qdm][ironLevel + ironD][2] += kg  # 计算结果累加
                        self.dictIron3['级别直径总重'][qdm][ironLevel + ironD][3] += kg  # 楼层累加
                # 接头个数
                if connectNum != 0:  # 有接头数量才执行
                    qdm = lowUppstr + "_" + connectType  # qdm 清单名称 一级子目
                    result = self.dictIron3['接头个数'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['接头个数'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['接头个数'][qdm + "量"] = connectNum
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['接头个数'][qdm + "量"] += connectNum  # 结果real累加

                    res = self.dictIron3['接头个数'][qdm].get(ironLevel + "_" + ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['接头个数'][qdm][ironLevel + "_" + ironD] = [connectNum, connectNum, connectNum, connectNum, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['接头个数'][qdm][ironLevel + "_" + ironD][0] += connectNum  # 拼接表达式 放弃 太长
                        self.dictIron3['接头个数'][qdm][ironLevel + "_" + ironD][1] += connectNum  # 结果real累加
                        self.dictIron3['接头个数'][qdm][ironLevel + "_" + ironD][2] += connectNum  # 计算结果累加
                        self.dictIron3['接头个数'][qdm][ironLevel + "_" + ironD][3] += connectNum  # 楼层累加
                # 接头直径个数
                if connectNum != 0:  # 有接头数量才执行
                    qdm = lowUppstr + "_" + connectType + "_" + ironD  # qdm 清单名称 一级子目
                    result = self.dictIron3['接头直径个数'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['接头直径个数'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['接头直径个数'][qdm + "量"] = connectNum
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['接头直径个数'][qdm + "量"] += connectNum  # 结果real累加

                    res = self.dictIron3['接头直径个数'][qdm].get(ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['接头直径个数'][qdm][ironD] = [connectNum, connectNum, connectNum, connectNum, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['接头直径个数'][qdm][ironD][0] += connectNum  # 拼接表达式 放弃 太长
                        self.dictIron3['接头直径个数'][qdm][ironD][1] += connectNum  # 结果real累加
                        self.dictIron3['接头直径个数'][qdm][ironD][2] += connectNum  # 计算结果累加
                        self.dictIron3['接头直径个数'][qdm][ironD][3] += connectNum  # 楼层累加
                # 楼层构件总重
                qdm = lowUppstr + "_" + floorName + "_钢筋"  # qdm 清单名称 一级子目
                result = self.dictIron3['楼层构件总重'].get(qdm)  # 找一级清单子目
                if result == None:  # 说明key未创建
                    self.dictIron3['楼层构件总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                    self.dictIron3['楼层构件总重'][qdm + "量"] = kg
                else:  # 说明级别已存在  工程总量做累加
                    self.dictIron3['楼层构件总重'][qdm + "量"] += kg  # 结果real累加
                res = self.dictIron3['楼层构件总重'][qdm].get(partsBig)  # 找二级明细子目  “楼层吗 = key = 6圆”
                if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                    self.dictIron3['楼层构件总重'][qdm][partsBig] = [kg, kg, kg, kg, "", "", ""]
                else:  # 直径已存在
                    self.dictIron3['楼层构件总重'][qdm][partsBig][0] += kg  # 拼接表达式 放弃 太长
                    self.dictIron3['楼层构件总重'][qdm][partsBig][1] += kg  # 结果real累加
                    self.dictIron3['楼层构件总重'][qdm][partsBig][2] += kg  # 计算结果累加
                    self.dictIron3['楼层构件总重'][qdm][partsBig][3] += kg  # 楼层累加
                # 构件楼层总重
                qdm = lowUppstr + "_" + partsBig + "_钢筋"  # qdm 清单名称 一级子目
                result = self.dictIron3['构件楼层总重'].get(qdm)  # 找一级清单子目
                if result == None:  # 说明key未创建
                    self.dictIron3['构件楼层总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                    self.dictIron3['构件楼层总重'][qdm + "量"] = kg
                else:  # 说明级别已存在  工程总量做累加
                    self.dictIron3['构件楼层总重'][qdm + "量"] += kg  # 结果real累加
                res = self.dictIron3['构件楼层总重'][qdm].get(floorName)  # 找二级明细子目  “楼层吗 = key = 6圆”
                if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                    self.dictIron3['构件楼层总重'][qdm][floorName] = [kg, kg, kg, kg, "", "", ""]
                else:  # 直径已存在
                    self.dictIron3['构件楼层总重'][qdm][floorName][0] += kg  # 拼接表达式 放弃 太长
                    self.dictIron3['构件楼层总重'][qdm][floorName][1] += kg  # 结果real累加
                    self.dictIron3['构件楼层总重'][qdm][floorName][2] += kg  # 计算结果累加
                    self.dictIron3['构件楼层总重'][qdm][floorName][3] += kg  # 楼层累加
                # 构件纵筋总重
                if not hoopKg:  # 纵筋
                    qdm = lowUppstr + "_" + partsBig + "_纵筋"  # qdm 清单名称 一级子目
                    result = self.dictIron3['构件纵筋总重'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['构件纵筋总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['构件纵筋总重'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['构件纵筋总重'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['构件纵筋总重'][qdm].get(ironLevel + "_" + ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + "_" + ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + "_" + ironD][0] += kg  # 拼接表达式 放弃 太长
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + "_" + ironD][1] += kg  # 结果real累加
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + "_" + ironD][2] += kg  # 计算结果累加
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + "_" + ironD][3] += kg  # 楼层累加
                else:   # 非纵筋
                    qdm = lowUppstr + "_" + partsBig + "_非纵筋"  # qdm 清单名称 一级子目
                    result = self.dictIron3['构件纵筋总重'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['构件纵筋总重'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['构件纵筋总重'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['构件纵筋总重'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['构件纵筋总重'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = A6”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + ironD][0] += kg  #  表达式
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + ironD][1] += kg  # 结果real累加
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + ironD][2] += kg  # 计算结果累加
                        self.dictIron3['构件纵筋总重'][qdm][ironLevel + ironD][3] += kg  # 楼层累加
                # 江苏2014_钢筋重量  砌体加筋先提 再冷拔丝 再其他
                if masonryIron == True:  # 判断砌体加筋、板缝筋
                    qdm = lowUppstr + "_砌体筋"  # qdm 清单名称 一级子目
                    result = self.dictIron3['江苏2014_钢筋重量'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['江苏2014_钢筋重量'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['江苏2014_钢筋重量'][qdm].get(ironLevel + "_" + ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + "_" + ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + "_" + ironD][0] += kg  # 拼接表达式 放弃 太长
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + "_" + ironD][1] += kg  # 结果real累加
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + "_" + ironD][2] += kg  # 计算结果累加
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + "_" + ironD][3] += kg  # 楼层累加
                elif drawingIron == True:  # 判断是否冷拔丝
                    qdm = lowUppstr + "_冷拔丝"  # qdm 清单名称 一级子目
                    result = self.dictIron3['江苏2014_钢筋重量'].get(qdm)  # 找一级清单子目
                    if result == None:  # 说明key未创建
                        self.dictIron3['江苏2014_钢筋重量'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                        self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] = kg
                    else:  # 说明级别已存在  工程总量做累加
                        self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] += kg  # 结果real累加
                    res = self.dictIron3['江苏2014_钢筋重量'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = A6”
                    if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD] = [kg, kg, kg, kg, "", "", ""]
                    else:  # 直径已存在
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][0] += kg  # 表达式
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][1] += kg  # 结果real累加
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][2] += kg  # 计算结果累加
                        self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][3] += kg  # 楼层累加
                else:  # 除砌体筋与冷拔丝以外的钢筋
                    if ironDInt <= 12:  # 12以内
                        qdm = lowUppstr + ironLevel + "_Φ12以内"  # qdm 清单名称 一级子目
                        result = self.dictIron3['江苏2014_钢筋重量'].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            self.dictIron3['江苏2014_钢筋重量'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                            self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] = kg
                        else:  # 说明级别已存在  工程总量做累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] += kg  # 结果real累加
                        res = self.dictIron3['江苏2014_钢筋重量'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = A6”
                        if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD] = [kg, kg, kg, kg, "", "", ""]
                        else:  # 直径已存在
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][0] += kg  # 表达式
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][1] += kg  # 结果real累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][2] += kg  # 计算结果累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][3] += kg  # 楼层累加
                    elif 12 < ironDInt <= 25:  # 25以内
                        qdm = lowUppstr + ironLevel + "_Φ25以内"  # qdm 清单名称 一级子目
                        result = self.dictIron3['江苏2014_钢筋重量'].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            self.dictIron3['江苏2014_钢筋重量'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                            self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] = kg
                        else:  # 说明级别已存在  工程总量做累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] += kg  # 结果real累加
                        res = self.dictIron3['江苏2014_钢筋重量'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = A6”
                        if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD] = [kg, kg, kg, kg, "", "", ""]
                        else:  # 直径已存在
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][0] += kg  # 表达式
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][1] += kg  # 结果real累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][2] += kg  # 计算结果累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][3] += kg  # 楼层累加
                    else:
                        qdm = lowUppstr + ironLevel + "_Φ25以外"  # qdm 清单名称 一级子目
                        result = self.dictIron3['江苏2014_钢筋重量'].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            self.dictIron3['江苏2014_钢筋重量'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                            self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] = kg
                        else:  # 说明级别已存在  工程总量做累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm + "量"] += kg  # 结果real累加
                        res = self.dictIron3['江苏2014_钢筋重量'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = A6”
                        if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD] = [kg, kg, kg, kg, "", "", ""]
                        else:  # 直径已存在
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][0] += kg  # 表达式
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][1] += kg  # 结果real累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][2] += kg  # 计算结果累加
                            self.dictIron3['江苏2014_钢筋重量'][qdm][ironLevel + ironD][3] += kg  # 楼层累加
                # 江苏2014_接头个数 电渣压力焊部分直径 其余分
                if connectNum != 0:  # 有接头数量才执行
                    if connectType == "电渣压力焊":  # 不区分直径
                        qdm = lowUppstr + "_" + connectType  # qdm 清单名称 一级子目
                        result = self.dictIron3['江苏2014_接头个数'].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            self.dictIron3['江苏2014_接头个数'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                            self.dictIron3['江苏2014_接头个数'][qdm + "量"] = connectNum
                        else:  # 说明级别已存在  工程总量做累加
                            self.dictIron3['江苏2014_接头个数'][qdm + "量"] += connectNum  # 结果real累加
                        res = self.dictIron3['江苏2014_接头个数'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                        if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                            self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD] = [connectNum, connectNum, connectNum,
                                                                                connectNum, "", "", ""]
                        else:  # 直径已存在
                            self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][0] += connectNum  # 拼接表达式 放弃 太长
                            self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][1] += connectNum  # 结果real累加
                            self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][2] += connectNum  # 计算结果累加
                            self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][3] += connectNum  # 楼层累加
                    else:  # 其余接头分 25以内 25以外
                        if ironDInt <= 25:  # 25以内
                            qdm = lowUppstr + "_" + connectType + "Φ25以内"  # qdm 清单名称 一级子目
                            result = self.dictIron3['江苏2014_接头个数'].get(qdm)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                self.dictIron3['江苏2014_接头个数'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                                self.dictIron3['江苏2014_接头个数'][qdm + "量"] = connectNum
                            else:  # 说明级别已存在  工程总量做累加
                                self.dictIron3['江苏2014_接头个数'][qdm + "量"] += connectNum  # 结果real累加
                            res = self.dictIron3['江苏2014_接头个数'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                            if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD] = [connectNum, connectNum, connectNum,
                                                                                    connectNum, "", "", ""]
                            else:  # 直径已存在
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][0] += connectNum  # 拼接表达式 放弃 太长
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][1] += connectNum  # 结果real累加
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][2] += connectNum  # 计算结果累加
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][3] += connectNum  # 楼层累加
                        else:  # 25以外 不含25圆
                            qdm = lowUppstr + "_" + connectType + "Φ25以外"  # qdm 清单名称 一级子目
                            result = self.dictIron3['江苏2014_接头个数'].get(qdm)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                self.dictIron3['江苏2014_接头个数'][qdm] = {}  # '地下_A':{},'地下_A重量'：5
                                self.dictIron3['江苏2014_接头个数'][qdm + "量"] = connectNum
                            else:  # 说明级别已存在  工程总量做累加
                                self.dictIron3['江苏2014_接头个数'][qdm + "量"] += connectNum  # 结果real累加
                            res = self.dictIron3['江苏2014_接头个数'][qdm].get(ironLevel + ironD)  # 找二级明细子目  “楼层吗 = key = 6圆”
                            if res == None:  # 说明 直径 key未创建  ironD = 楼层号
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD] = [connectNum, connectNum,
                                                                                         connectNum, connectNum, "", "", ""]
                            else:  # 直径已存在
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][0] += connectNum  # 拼接表达式 放弃 太长
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][1] += connectNum  # 结果real累加
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][2] += connectNum  # 计算结果累加
                                self.dictIron3['江苏2014_接头个数'][qdm][ironLevel + ironD][3] += connectNum  # 楼层累加
            # self.dictIron3  完成
            #  输出考虑选择了哪些维度的报表
            # 向 工程量总表 发送数据
            # 字典整理成二维表
            # print("self.dictIron3", self.dictIron3)
            # print("field_name_list", field_name_list)
            # key = '级别总重'  ,value  = {清单名:{}，清单名+重量 = 工程量}  2个元素
            for key, value in self.dictIron3.items():  # 取出每一个维度的统计表
                if "个数" in key:
                    flag = "个"
                else:
                    flag = "kg"
                if len(value) < 1:  # 如果是空表 则跳过
                    continue
                if key not in field_name_list:   # 如果这个报表类型未选择 则跳过
                    continue
                NoNumber = 1  # 序号
                # TODO ka = 清单名，清单名+重量  2个元素  ；va = {行号：list} 或 float 值
                for ke, va in value.items():   # 取出 每一个分类的统计表
                    if isinstance(va, dict):  # 非字典就是 清单重量：131.12
                        temp = self.dictIron3[key][ke + "量"]
                        temp = round(temp, 3)
                        Quantities = str(temp)
                        detailslist = []  # 明细的二维表
                        # k = 楼层号； v 列表 = 一行明细数据
                        for k, v in va.items():
                            rowlist = []
                            rowlist.append(k)
                            for cell in v:
                                if cell != "":
                                    rowlist.append(str(round(cell, 3)))
                                else:
                                    rowlist.append(str(cell))
                            detailslist.append(rowlist)
                    else:  # va 取出值的一个元素 '地下_A级重量': 6.328
                        continue
                    valuelist = [str(NoNumber), '文件序号_' + id, '', key, ke, flag, Quantities, str(detailslist), '']  # 按清单表结构 一行
                    valuelists.append(valuelist)  # 都需要为字符型
                    NoNumber += 1
                # 加一个空行
                valuelists.append(blankrow)
        # print("valuelists", valuelists)
        # 二维表写入模型
        # 断开链接
        self.itemModel32.itemChanged.disconnect(self.itemModel32_itemChanged)
        rows = len(valuelists)
        cols = len(blankrow)
        for row in range(rows):
            itemlist = []
            for col in range(cols):  # 每一行数据 中的每一个单元格
                item = valuelists[row][col]
                item = QStandardItem(item)
                itemlist.append(item)
            self.itemModel32.appendRow(itemlist)
        # ~~~~~~最后做的几项收尾工作~~~~~~~~
        # self.changeColor()   # 《来源》列 变色
        # self.m3ChangeRed(self.itemModel11)  # m3 变色
        self.quantitiesSum32()  # 钢筋表所有行的 汇总总计
        # self.ui.QuanBtnImportClassificationToQuan.setEnabled(False)   # 发送按钮不可用
        # 恢复链接
        self.itemModel32.itemChanged.connect(self.itemModel32_itemChanged)
        self.ui.QuanBtnIronImport.setEnabled(False)
        self.ui.QuanBtnIronOK.setEnabled(True)
        # self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged)
        # print(self.dictIron3)

    # 钢筋表 发送至 工程量总表
    @pyqtSlot()
    def on_QuanBtnIronOK_clicked(self):
        # print("钢筋表 发送至 工程量总表")
        name = "钢筋工程量表"
        # 钢筋表 土建表 合并 整理成二维list
        result = self.ironGlodonSendMaintable(self.itemModel32, name)
        # 整理后的字典（钢筋 土建分表） 发至总表
        self.subtable_to_main(result)
        # pprint(result)
        # 写入前 先切换页签 和标志
        self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 页面
        self.tabindex = "工程量总表"  # 清单表

    # 测试1
    @pyqtSlot()
    def on_QuanBtnceshi_clicked(self):
        print("测试1")

    # 字体大小 调节按钮
    # @pyqtSlot()
    def spinBoxSize_valueChanged(self):
        # print("字体大小 调节按钮 更改了！")
        self.spinFontSize = self.ui.spinBoxSize.value()
        font = QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(self.spinFontSize)
        self.ui.tableView_11.setFont(font)
        self.ui.tableView_12.setFont(font)
        self.ui.tableView_2.setFont(font)
        self.ui.tableView_32.setFont(font)
        self.ui.tableView_33.setFont(font)
        self.ui.tableView_42.setFont(font)
        self.ui.tableView_43.setFont(font)

    # @pyqtSlot()
    # def on_btnOK_clicked(self):
    #     print("确定")

    # @pyqtSlot()
    # def on_btnExit_clicked(self):
    #     print("退出")

##  =============自定义槽函数===============================        

##  ============窗体测试程序 ================================
if __name__ == "__main__":  # 用于当前窗体测试
    app = QApplication(sys.argv)  # 创建GUI应用程序
    form = QmyDialogQuantities()  # 创建窗体
    form.show()
    sys.exit(app.exec_())
