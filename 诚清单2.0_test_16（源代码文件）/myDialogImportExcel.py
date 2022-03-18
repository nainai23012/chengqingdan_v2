import sys
import xlrd
import re
import os
from PyQt5.QtWidgets import QApplication, QDialog, QAbstractItemView, QHeaderView, QDialog, \
    QFileDialog, QMessageBox

##from PyQt5.QtWidgets import  QDialog

from PyQt5.QtCore import Qt, pyqtSlot, QItemSelectionModel, QDir

##from PyQt5.QtCore import  pyqtSlot,Qt

##from PyQt5.QtWidgets import  

from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor

from ui_QWDialogImportExcel import Ui_Dialog
from myDelegates import QmyFloatSpinDelegate, QmyComboBoxDelegate  # 定义代理模块


class QmyDialogImportExcel(QDialog):
    def __init__(self, rowCount=3, colCount=5, parent=None):
        super().__init__(parent)  # 调用父类构造函数，创建窗体
        self.ui = Ui_Dialog()  # 创建UI对象
        self.ui.setupUi(self)  # 构造UI界面

        self.edition = '适用于Excel版本清单的导入'  # 统一的版本号
        self.setWindowTitle(self.edition)  # 设置窗口名称

        self.__filename = None  # 工作薄 全路径 Str
        self.sheetname = None  # 工作表名字 Str
        self.rows = 0  # 导入的表共计多少行

        # self.qualities = {'': '', '序号': 0, '楼层/部位': 1, '清单名称': 2, '计量单位': 3, '计算表达式': 4,
        #                   '工程量': 5, '不计标志': 6, '公式错误': 7, '备注': 8}  # value与主表中的 列索引对应
        self.qualities = {'': '', '序号': 0, '色标': 1, '地上地下': 2, '项目编码': 3, '项目名称': 4, '项目特征描述': 5,
                          '计量单位': 6, '工程量': 7, '∑工程量明细表': 8, '综合单价': 9, '项目合价': 10, '备注': 11,
                          '人工费': 12, '主材费': 13, '辅材费': 14, '机械费': 15, '管理费、利润': 16, '规费': 17,
                          '不含税综合单价': 18, '增值税税金': 19}  # value与主表中的 列索引对应
        self.sumcomboboxlist = ['', '序号', '项目编码', '项目名称', '项目特征描述', '计量单位', '工程量', '综合单价', '备注',
                                '人工费', '主材费', '辅材费', '机械费', '管理费、利润', '规费', '不含税综合单价', '增值税税金']

        self.__initmatching2()  # 初始化 表头下拉菜单的
        # self.ui.detailedListRowmin.setEnabled(False)  # 不可用
        # self.ui.detailedListRowmax.setEnabled(False)  # 不可用
        self.ui.detailedListOK.setEnabled(False)  # 确定按钮 不可用
        self.ui.detailedListAuto.setEnabled(False)  # 自动识别字段按钮 不可用

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
        # sumcomboboxlist = ['', '楼层/部位', '清单名称', '计量单位', '计算表达式', '工程量', '备注']
        sumcomboboxlist = ['', '序号', '项目编码', '项目名称', '项目特征描述', '计量单位', '工程量', '综合单价', '备注',
                                '人工费', '主材费', '辅材费', '机械费', '管理费、利润', '规费', '不含税综合单价', '增值税税金']
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
        self.itemModel2.dataChanged.connect(self.do_dataChanged)  # 主表有变化时触发

    # 自动识别字段
    @pyqtSlot(bool)
    def on_detailedListAuto_clicked(self):
        print('自动识别字段')
        rows = self.itemModel2.rowCount()
        cols = self.itemModel2.columnCount()
        if rows < 4:
            return
        filepath = r'.\data\main_data\FieldNameDictionary.xlsx'
        sheetname = r'FieldNameDictionary'
        result = os.path.exists(filepath)  # 判断是否存在
        if result != True:
            return  # 返回
        with xlrd.open_workbook(filepath) as f:
            sheet = f.sheet_by_name(sheetname)
            datas = sheet._cell_values
            datarows = sheet.nrows
        print(datarows)
        print(datas)
        for datarow in range(1,datarows):  # 原始数据列
            fieldname = datas[datarow][0]
            yeslist = datas[datarow][1]
            nolist = datas[datarow][2]
            flag = None
            for row in range(1, 10):  # 在前20行范围内搜索
                if flag != None:
                    break
                for col in range(cols):
                    item = self.itemModel2.item(row, col).text().strip()
                    if item == "":
                        continue
                    if item in yeslist and item not in nolist:
                        itemkey = QStandardItem(fieldname)
                        self.itemModel2.setItem(0, col, itemkey)
                        flag = "OK"
                        break

        # print(self.sumcomboboxlist)  # TODO 与\data\main_data\FieldNameDictionary.xlsx  数据要对应
        # rows = self.itemModel2.rowCount()
        # cols = self.itemModel2.columnCount()
        # for keystr in self.sumcomboboxlist:
        #     if keystr == "":
        #         continue
        #     flag = None
        #     for row in range(1, 20):  # 在前20行范围内搜索
        #         if flag != None:
        #             break
        #         for col in range(cols):
        #             item = self.itemModel2.item(row, col).text()
        #             if item == "":
        #                 continue
        #             if keystr == item:
        #                 itemkey = QStandardItem(keystr)
        #                 self.itemModel2.setItem(0, col, itemkey)
        #                 flag = "OK"
        #                 break

    # 确定 按钮
    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_detailedListOK_clicked(self):  # 确定 按钮
        print('确定 按钮')
        # TODO 处理有效字段  self.itemModel2
        self.dataList = [[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []]  # 0 ~ 8 列 序号~备注
        rows = self.itemModel2.rowCount()
        cols = self.itemModel2.columnCount()

        for col in range(cols):  # 循环第一行  获取真实的列号
            try:
                temp = self.itemModel2.item(0, col).text()  # 没有值  会报错
                # print(type(temp))
            except:
                continue
            if temp:  # 如果下拉列表 有字符  则执行
                dataCol = int(self.qualities[temp])  # 取出真正的列号
                # print(f'真正的列号{dataCol}',type(dataCol))
                for row in range(1, rows):  # 跳过第一行（下拉列表行）
                    self.dataList[dataCol].append(self.itemModel2.item(row, col).text())
        # TODO '处理竖向合并的项目名称'    如 “主体结构” “二次结构”等
        dcols = len(self.dataList)
        drows = len(self.dataList[4])  # 清单名称列的数量 = 总行数
        # for x in range(dcols):  # 循环找到最多行的一列，一样多
        #     if len(self.dataList[x]) > drows:
        #         drows = len(self.dataList[x])
        #         break
        # if drows < 1:  # 说明没选 表头 则退出
        #     QMessageBox.critical(self, "错误", "没有选择第一行的字段名", QMessageBox.Cancel)
        #     return
        if drows < 1:
            QMessageBox.critical(self, "提示", "您导入的清单表可能数据不全，请确认 字段名与所在列是否对应！", QMessageBox.Ok)
            return
        print('共计行数： ', drows)
        for row in range(drows):  # 没选名称和计量单位时会报错
            try:
                name = self.dataList[4][row]
                danwei = self.dataList[6][row]
                # TODO 名称为空  计量单位不为空 且计量单位等于上一行
                if not name and danwei and (danwei == self.dataList[6][row - 1]):
                    self.dataList[4][row] = self.dataList[4][row - 1]
                    # print('找到一个合并单元格',self.dataList[2][row-1])
            except:
                pass
        print("获取“地下工程”“地下部分” 字眼可以前后、字间空格")
        strdx, strds = None, None
        rowdx, rowds = None, None
        for row in range(1, rows):  # 跳过第一行（下拉列表行）
            for col in range(cols):
                temp = self.itemModel2.item(row, col).text()
                if temp == '':
                    continue
                temp = temp.strip()
                if not strdx:
                    # strdx = re.findall('^\s*?地\s*?下\s*?工\s*?程\s*?$|^\s*?地\s*?下\s*?部\s*?分\s*?$', temp)
                    strdx = re.findall('^\s*?地\s*?下\s*?工\s*?程\s*?$|^\s*?地\s*?下\s*?部\s*?分\s*?$', temp)
                    if strdx:
                        rowdx = row
                        continue
                if not strds:
                    strds = re.findall('^\s*?地\s*?上\s*?工\s*?程\s*?$|^\s*?地\s*?上\s*?部\s*?分\s*?$', temp)
                    if strds:
                        rowds = row
                        continue
        print('找到地下，地上 行号：', rowdx, rowds)
        # 有项目名称的写入 地下 地上 前缀
        if not rowdx and not rowds:  # 项目名称有数据 才写入 地上 地下标志
            return   #  没有找到 地上工程 地上部分  地下工程 地下分部 等字样 则跳过下面步骤
        if len(self.dataList[1]) < 1:
            # self.dataList[1] = [[] for a in range(drows)]  # 没有选序号 字段 则该字段用空值填入
            self.dataList[1] = ['' for a in range(drows)]  # 没有选序号 字段 则该字段用空值填入
        if rowdx:
            self.dataList[1][rowdx-1] = "1,红"  # 写入色标
        if rowds:
            self.dataList[1][rowds-1] = "1,红"  # 写入色标

        if len(self.dataList[2]) < 1:
            # self.dataList[2] = [[] for a in range(drows)]  # 没有选序号 字段 则该字段用空值填入
            self.dataList[2] = ['' for a in range(drows)]  # 没有选序号 字段 则该字段用空值填入
        if rowdx and not rowds:  # 只有地下 没有 地上
            for row in range(drows):
                if not self.dataList[4][row]:  # 项目名称为空时 不写入 地上地下
                    continue
                self.dataList[2][row] = '地下'
        elif not rowdx and rowds:  # 只有地上 没有地下
            for row in range(drows):
                if not self.dataList[4][row]:  # 项目名称为空时 不写入 地上地下
                    continue
                self.dataList[2][row] = '地上'
        else:  # 既有地上 又有地下标记
            for row in range(drows):
                if not self.dataList[4][row]:  # 项目名称为空时 不写入 地上地下
                    continue
                if rowdx < rowds:  # TODO 预防地下 地上颠倒
                    if rowds - 1 > row >= rowdx:  # 大于地下，小于地上
                        self.dataList[2][row] = '地下'
                    elif row >= rowds:  # 大于地上
                        self.dataList[2][row] = '地上'
                else:
                    if rowdx - 1 > row >= rowds:
                        self.dataList[2][row] = '地上'
                    elif row >= rowdx:  # 大于地上
                        self.dataList[2][row] = '地下'
        print('self.dataList', self.dataList)

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
        self.ui.detailedListAuto.setEnabled(True)  # 自动识别字段按钮 不可用
        self.sheetname = curText  # 得到表名
        self.cellsMap()  # 把表写入数据模型  且显示在表格控件上

    # 选择清单表文件：
    @pyqtSlot()  # 选择清单文件,把工作薄内的每一张表都装入下拉列表 按钮
    def on_detailedListLoad_clicked(self):
        print("选择清单表文件：")

        curPath = QDir.currentPath()  # 获取系统当前目录
        dlgTitle = "选择清单表文件"  # 对话框标题
        # filt = "清单文件(*.xls *.xlsx);;所有文件(*.*)"  # 文件过滤器
        filt = "清单表文件(*.xls *.xlsx)"  # 文件过滤器

        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        if not filename:
            return
        self.__initmatching2()
        self.__filename = filename
        self.ui.label_path.setText(f'导入文件路径：{filename}')
        try:
            with xlrd.open_workbook(filename) as f:
                sheetsName = f.sheet_names()  # 获取所有sheet名字
                self.ui.detailedListCbx.clear()  # 先清空下拉列表
                self.ui.detailedListCbx.addItems(sheetsName)
            self.ui.detailedListLoad.setEnabled(False)
            # print(sheetsName)
        except:
            QMessageBox.critical(self, "错误", "文件格式非法！", QMessageBox.Cancel)
            return

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
    form = QmyDialogImportExcel()  # 创建窗体
    form.show()
    sys.exit(app.exec_())
