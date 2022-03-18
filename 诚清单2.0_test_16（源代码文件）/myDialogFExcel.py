import sys
import xlrd

from PyQt5.QtWidgets import QApplication, QDialog, QAbstractItemView, QHeaderView, QDialog, \
    QFileDialog

##from PyQt5.QtWidgets import  QDialog

from PyQt5.QtCore import Qt, pyqtSlot, QItemSelectionModel, QDir

##from PyQt5.QtCore import  pyqtSlot,Qt

##from PyQt5.QtWidgets import  

from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor

from ui_QWDialogFExcel import Ui_Dialog
from myDelegates import QmyFloatSpinDelegate, QmyComboBoxDelegate  # 定义代理模块


class QmyDialogFExcel(QDialog):
    def __init__(self, rowCount=3, colCount=5, parent=None):
        super().__init__(parent)  # 调用父类构造函数，创建窗体
        self.ui = Ui_Dialog()  # 创建UI对象
        self.ui.setupUi(self)  # 构造UI界面

        self.edition = '工程量清单表中的《分类表》导入'  # 统一的版本号
        self.setWindowTitle(self.edition)  # 设置窗口名称

        self.__filename = None  # 工作薄 全路径 Str
        self.sheetname = None  # 工作表名字 Str
        self.rows = 0  # 导入的表共计多少行

        self.qualities = {'': '', '序号': 0, '楼层/部位': 1, '清单名称': 2, '计量单位': 3, '计算表达式': 4,
                          '工程量': 5, '不计标志': 6, '公式错误': 7, '备注': 8}  # value与主表中的 列索引对应
        self.sumcomboboxlist = ['', '楼层/部位', '清单名称', '计量单位', '计算表达式', '工程量', '备注']

        self.__initmatching2()  # 初始化 表头下拉菜单的
        # self.ui.detailedListRowmin.setEnabled(False)  # 不可用
        # self.ui.detailedListRowmax.setEnabled(False)  # 不可用
        self.ui.detailedListOK.setEnabled(False)  # 确定按钮 不可用

        # self.setWindowFlags(Qt.MSWindowsFixedSizeDialogHint)
        # self.setIniSize(rowCount,colCount)

    def __del__(self):  # 析构函数
        ##      super().__del__(self)
        print("QmyDialogSize 对象被删除了")

    # TODO  ==============初始化功能============
    # 非常重要 传入一个新建的数据模型 全部赋值为空白 否则无法获取text（）
    def initItemModelBlank(self, itemModel):
        rows = itemModel.rowCount()
        cols = itemModel.columnCount()
        for x in range(rows):
            for y in range(cols):
                item = QStandardItem("")
                itemModel.setItem(x, y, item)  # 初始化表 itemModel 每项目为空

    # 下拉列表表头的 表格控件 初始化
    def __initmatching2(self):
        self.ui.tableView_2.setEnabled(False)  # 表格控件不编辑
        self.itemModel2 = QStandardItemModel(1, 10, self)  # TODO 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel2)  # 初始化
        # self.qualities = {'': '', '序号': 0, '编码': 1, '部位/楼层': 2, '清单名称': 3, '计量单位': 4,'计算表达式': 5,  '工程量': 6,
        #                   '综合单价': 7,'合价': 8, '不计标志': 9,'错误': 10,   '备注': 11}  # value与主表中的 列索引对应


        # comboboxlist = [x for x in self.qualities.keys()]
        print('comboboxlist', self.sumcomboboxlist)
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)  # 添加代理组件 combobox
        self.UnitOfMeasurement.setItems(self.sumcomboboxlist, False)  # 下拉列表添加 条目 不可编辑

        self.itemModel2.dataChanged.connect(self.do_dataChanged)  # TODO 主表有变化时触发

        self.ui.tableView_2.setItemDelegateForRow(0, self.UnitOfMeasurement)  # 在第一行 添加代理组件

        self.ui.tableView_2.setModel(self.itemModel2)  # 设置数据模型

        self.ui.tableView_2.verticalHeader().setDefaultSectionSize(22)  # 缺省行高

        self.ui.tableView_2.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9}"
            "QComboBox{margin:2px};")  # 表头样式，代理组件combobox 内边距2
        # cols = self.itemModel2.columnCount()
        # for col in range(cols):
        #     item = QStandardItem("")
        #     font = item.font()
        #     font.setBold(True)
        #     # item.setForeground(QColor(200, 0, 0))
        #     item.setBackground(QColor(34, 165, 255))
        #     item.setFont(font)
        #     self.itemModel2.setItem(0, col, item)

    ##  ==============自定义功能函数============
    def setIniSize(self, rowCount, colCount):  ##设置表格大小
        pass
        # self.ui.spin_RwoCount.setValue(rowCount)
        # self.ui.spin_ColCount.setValue(colCount)

    def getData(self):  ##以元组数据同时返回行数和列数
        return self.dataList, self.rows  # 二维表，行数  TODO 是否需要输出 行数
        # return self.dataList   #二维表  TODO 是否需要输出
        # rows = self.ui.spin_RwoCount.value()
        # cols = self.ui.spin_ColCount.value()
        # return rows, cols

    ##  ==========由connectSlotsByName() 自动连接的槽函数==================
    # 表数据变化时触发
    def do_dataChanged(self, curr):
        if not curr:
            return
        row = curr.row()
        col = curr.column()
        if row != 0:
            return
        # print('表数据变化时触发1')
        self.itemModel2.dataChanged.disconnect(self.do_dataChanged)  # TODO 主表有变化时触发
        sumcomboboxlist = ['', '楼层/部位', '清单名称', '计量单位', '计算表达式', '工程量', '备注']
        comboboxlist = []
        # sumcomboboxlist = ['', '序号', '编码', '部位/楼层', '清单名称', '计量单位', '计算表达式', '工程量', '综合单价', '合价', '不计标志', '错误', '备注']
        # self.sumcomboboxlist = ['', '序号', '楼层/部位', '清单名称', '计量单位', '计算表达式', '工程量', '备注']
        # print('表数据变化时触发2')
        rows = self.itemModel2.rowCount()
        cols = self.itemModel2.columnCount()
        for c in range(cols):
            # print('表数据变化时触发3')
            try:
                temp = self.itemModel2.item(row, c).text()
                if temp:
                    comboboxlist.append(temp)
            except:
                continue
            # print('表数据变化时触发5')
        if len(comboboxlist) < 1:
            return
        # print('comboboxlist', comboboxlist)
        # print('sumcomboboxlist', sumcomboboxlist)
        for com in comboboxlist:
            if com:
                sumcomboboxlist.remove(com)
        # print('表数据变化时触发6')

        self.UnitOfMeasurement.setItems(sumcomboboxlist, False)  # 下拉列表添加 条目 不可编辑

        self.itemModel2.dataChanged.connect(self.do_dataChanged)  # TODO 主表有变化时触发

    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_pubton_clicked(self):  # 测试 按钮
        print('测试 按钮')
        # temp = self.itemModel2.item(2, 2).text()  # 没有值  会报错
        # print(type(temp),temp)

    # 确定 按钮
    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_detailedListOK_clicked(self):  # 确定 按钮
        print('确定 按钮')
        # TODO 处理有效字段  self.itemModel2
        self.dataList = [[], [], [], [], [], [], [], [], []]  # 0 ~ 8 列 序号~备注
        rows = self.itemModel2.rowCount()
        cols = self.itemModel2.columnCount()

        for col in range(cols):  # 循环第一行  获取真实的列号
            try:
                temp = self.itemModel2.item(0, col).text()  # 没有值  会报错
                print(type(temp))
            except:
                continue
            if temp:  # 如果下拉列表 有字符  则执行
                dataCol = int(self.qualities[temp])  # 取出真正的列号
                # print(f'真正的列号{dataCol}',type(dataCol))
                for row in range(1, rows):  # 跳过第一行（下拉列表行）
                    self.dataList[dataCol].append(self.itemModel2.item(row, col).text())

    @pyqtSlot(str)  ##“简单的ComboBox”的当前项变化  下拉框发生改变时 表格显示的变化
    def on_detailedListCbx_currentIndexChanged(self, curText):
        # ~~~~~~~~表格控件初始化~~~~~~~~~
        # 第一行清空
        cols = self.itemModel2.columnCount()
        for col in range(cols):
            item = QStandardItem("")
            self.itemModel2.setItem(0, col, item)
        self.ui.tableView_2.setEnabled(True)  # 表格控件可编辑
        self.ui.detailedListOK.setEnabled(True)  # 确定按钮 可用
        self.sheetname = curText  # 得到表名
        self.cellsMap()  # 把表写入数据模型  且显示在表格控件上

    # 选择清单文件
    @pyqtSlot()  # 选择清单文件,把工作薄内的每一张表都装入下拉列表 按钮
    def on_detailedListLoad_clicked(self):
        # print("重新提取")
        curPath = QDir.currentPath()  # 获取系统当前目录
        dlgTitle = "选择分类表文件"  # 对话框标题
        # filt = "清单文件(*.xls *.xlsx);;所有文件(*.*)"  # 文件过滤器
        filt = "分类表文件(*.xls *.xlsx)"  # 文件过滤器

        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        if filename:
            self.__initmatching2()

            self.__filename = filename
            self.ui.label_path.setText(f'导入文件路径：{filename}')
            with xlrd.open_workbook(filename) as f:
                sheetsName = f.sheet_names()  # 获取所有sheet名字
                self.ui.detailedListCbx.clear()  # 先清空下拉列表
                self.ui.detailedListCbx.addItems(sheetsName)
            self.ui.detailedListLoad.setEnabled(False)
                # print(sheetsName)

    ##  =============自定义槽函数===============================
    # 根据表名,#表写入数据模型
    def cellsMap(self):
        with xlrd.open_workbook(self.__filename) as f:
            sheet = f.sheet_by_name(self.sheetname)
            rows = sheet.nrows
            self.rows = rows
            cols = sheet.ncols
            # ~~~~~~~设置右下角 label~~~~~~
            self.ui.label_totalRows.setText(str(rows))
            self.ui.label_totalCols.setText(str(cols))
            # ~~~~~~~设置表头 tableview~~~~~~
            self.itemModel2.setColumnCount(cols)
            self.itemModel2.setRowCount(rows + 1)
            # self.ui.detailedListRowmax.setValue(rows + 1)
            for datax in range(rows):
                for datay in range(cols):
                    temp = sheet.cell_value(datax, datay)
                    if temp:
                        item = str(temp)
                    else:
                        item = 0
                    # temp = sheet.cell_value(datax, datay)
                    item = QStandardItem(item)
                    self.itemModel2.setItem(datax + 1, datay, item)  # 第一行是 下拉 列名行


##  ============窗体测试程序 ================================
if __name__ == "__main__":  # 用于当前窗体测试
    app = QApplication(sys.argv)  # 创建GUI应用程序
    form = QmyDialogFExcel()  # 创建窗体
    form.show()
    sys.exit(app.exec_())
