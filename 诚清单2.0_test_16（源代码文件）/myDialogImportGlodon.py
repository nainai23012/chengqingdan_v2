'''
工程量匹配表的业务相关代码
'''
import sys
import xlrd
import re
from pickle import dump, load

from PyQt5.QtWidgets import QApplication, QDialog, QLabel, QAbstractItemView, QHeaderView, QFileDialog
##from PyQt5.QtWidgets import  QDialog
from PyQt5.QtCore import Qt, pyqtSlot, QItemSelectionModel, QDir
##from PyQt5.QtCore import  pyqtSlot,Qt
##from PyQt5.QtWidgets import
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor

# from ui_QWDialogMatching import Ui_QWDialogMatching
from ui_QWDialogImportGlodon import Ui_Dialog


# from icecream import ic  # TODO 输出测试


class QmyDialogImportGlodon(QDialog):
    def __init__(self, rowCount=3, colCount=5, parent=None):
        super().__init__(parent)  # 调用父类构造函数，创建窗体
        self.ui = Ui_Dialog()  # 创建UI对象
        self.ui.setupUi(self)  # 构造UI界面

        self.__initmatching()  # 初始化 匹配表
        # self.__buildUI()  # 动态创建组件，添加到工具栏和状态栏

        self.ui.MacomboBoxModel.setEditable(False)  # 下拉控件不可编辑
        self.ui.MaBtnDel.setEnabled(False)  # 按钮控件不可编辑
        # self.setWindowFlags(Qt.MSWindowsFixedSizeDialogHint)
        # self.setIniSize(rowCount, colCount)

    def __del__(self):  # 析构函数
        ##      super().__del__(self)
        print("QmyDialogMatching 对象被删除了")

    ##  ==============自定义功能函数============
    def savepickle(self, data):  # 存泡菜
        pickle_file = open(r'data\quan_data\默认设置.pkl', 'wb')

        dump(data, pickle_file)  # 将列表倒入文件
        pickle_file.close()  # 关闭pickle文件

    def loadpickle(self):  # 取泡菜
        pickle_file = open(r'data\quan_data\默认设置.pkl', 'rb')
        self.matchingList = load(pickle_file)  # 取得匹配表数据 self.matchingList 字典形式｛默认表：[[],[],[]]｝
        pickle_file.close()

    def __initmatching(self):
        # ~~~~~~载入匹配表数据~~~~~~~~~~~~~~
        # print('__initmatching')
        self.loadpickle()  # 取泡菜
        YorNUpDown = self.matchingList['其他']['区分地上地下']
        self.ui.MacomboBoxUpDown.setCurrentText(YorNUpDown)  # 初始化区分地上地下
        nokeys = self.matchingList['其他']['不提量关键字']  # 列表类型
        str1 = ""
        for no in nokeys:
            str1 = str1 + "," + no
        str1 = str1[1:]
        self.ui.lineEditNoCal.setText(str1)  # 初始化 不提量关键字

        self.itemModel = QStandardItemModel(5, 4, self)  # TODO 数据模型,初始化 不定义的话 会出错

        nameslist = self.matchingList.keys()
        for name in nameslist:
            if name == "其他":
                continue
            self.ui.MacomboBoxModel.addItem(name)  # 初始化下拉菜单  会触发下拉列表事件

        self.selectionModel = QItemSelectionModel(self.itemModel)  # Item选择模型
        self.selectionModel.currentChanged.connect(self.do_curChanged)  # 自定义槽函数 选择项发生改变时


        self.ui.MaTableWidget.setModel(self.itemModel)  # 设置数据模型
        self.ui.MaTableWidget.setSelectionModel(self.selectionModel)  # 设置选择模型

        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.MaTableWidget.setSelectionMode(oneOrMore)  # 可多选

        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.MaTableWidget.setSelectionBehavior(itemOrRow)  # 单元格选择

        # self.ui.MaTableWidget.verticalHeader().setDefaultSectionSize(22)  # 缺省行高
        self.ui.MaTableWidget.horizontalHeader().setDefaultSectionSize(120)  # 缺省列宽
        self.ui.MaTableWidget.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 换行
        self.ui.MaTableWidget.setAlternatingRowColors(True)  # 交替行颜色
        # 设置表头边框样式
        self.ui.MaTableWidget.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9}")

    # 返回匹配表字典文件
    def getmatching(self):
        listdata = []  # 嵌套列表形式
        keystr = self.ui.MacomboBoxModel.currentText()
        # print(keystr)
        rows = self.itemModel.rowCount()
        cols = self.itemModel.columnCount()

        for i in range(rows):  # 从0开始
            for j in range(cols):  # 从0开始
                item = self.itemModel.item(i, j).text()  # 获取每一项内的字符
                if j == 0:  # 如果是每一行的第一个数据 则添加一个空列表
                    listdata.append([])
                listdata[i].append(item)
        if not keystr in self.matchingList.keys():  # 如果模板名称不在字典中 增加下拉列表
            self.ui.MacomboBoxModel.insertItem(self.ui.MacomboBoxModel.count(), str(keystr))
        YorNUpDown = self.ui.MacomboBoxUpDown.currentText() # 是否区分地上地下
        Nokeys = list(self.ui.lineEditNoCal.text().split(","))  # 不提量关键字 list 类型
        self.matchingList['其他'] = {'区分地上地下': YorNUpDown, '不提量关键字': Nokeys}
        self.matchingList[keystr] = listdata  # 列表写入字典
        self.savepickle(self.matchingList)  # 存泡菜
        return self.matchingList

    ##以二维表数据返回提取后的工程量数据,用于主表调用时的返回
    def getTableSize(self):
        totalList = []  # 存放提取后的二维表数据
        # path = self.ui.lineEdit.text()
        # if not path:
        #     return
        wbdict = self.wbToData(path)  # 得到excel文件工程量的原始数据

        matchingList = self.modleToList(self.itemModel)  # 获取匹配表 二维表格式
        YorNUpDown = self.ui.MacomboBoxUpDown.currentText()  # 区分 / 不区分
        Nokeys = list(self.ui.lineEditNoCal.text().split(","))  # 不提量关键字 list 类型
        for sht, items in wbdict.items():
            for x in range(len(matchingList)):
                if x % 2:  # 偶数行跳过
                    continue
                if sht == matchingList[x][0]:  # 表名和匹配表内的工程量表名相同
                    matchingListevery = matchingList[x:x + 2]  # 取出匹配的 两行列表  工程量&简称
                    # 处理每一张匹配的二维表
                    items = self.extractEverySheet(sht, items, matchingListevery, YorNUpDown, Nokeys)
                    if items:
                        totalList.extend(items)  # 每个表提取的数据datalist 扩展到totalList

        return totalList

    #  ==========自定义的功能函数================== 0
    # 用于从excel文件提取数据 字典 {"表名1":([],[],[],...),"表名2":([],[],[],...)}  辅助
    def wbToData(self, path):
        try:
            wb = xlrd.open_workbook(path)
            sheetCount = wb.nsheets  # 工作薄内所有表的数量
            # TODO 每张表+二维表数据 写入 字典 加速程序运行
            wbdict = {}
            for i in range(sheetCount):  # 循环每一张工作表
                # print(sht.name)
                shortSheetName = wb.sheet_by_index(i).name[
                                 wb.sheet_by_index(i).name.rfind("-") + 1:].strip()
                                # 简称 绘图输入工程量汇总表(按构件) - 柱 ~~~~~~~~~ 柱
                item = wb.sheet_by_index(i)._cell_values
                # item=tuple(item)  # TODO 元组替换列表 提速改进
                wbdict[shortSheetName] = item  # 表名换成简写
                # print(wbdict)
        except:
            pass
        finally:
            pass
        return wbdict

    # 处理每一张找到的工程量表  辅助  get 方法
    # 参数依次为： 表对象，匹配字，地上地下，不提量
    def extractEverySheet(self, shtName, itemsList, matchingListevery, YorNUpDown, Nokeys):  # matchingListevery 两行多列表
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
            if itemsList[row][col] in ["首层", "第1层", "第一层", "1层", "一层"]:  # TODO 首层的楼层名字
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
                    pass
                # ~~~~~~~~~判断首层所在 区分地上地下~~~~~~~~~~~~~~0
                if len(keydict["楼层"]) < 1:  # 楼层名称未打开 则跳过  疑似无用
                    continue
                floorname = itemsList[row][keydict["楼层"][0]]
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

                    # else:  # 要区分地上地下，又找到了首层的行号，在每行中判断
                    #     pass  # 多轴网会有多个 首层  暂时先注释
                        # if row < keydict["首层"]:
                        #     YorNUpDownName = '地下_'
                        # else:
                        #     YorNUpDownName = '地上_'
                # ~~~~~~~~~判断首层所在 区分地上地下~~~~~~~~~~~~~~1
                # print("是否区分地上地下：", YorNUpDownName)

                # if not num:  # 等于0 跳过
                #     continue
                # print('num', type(num), num)
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
                                 '单梁装修', '屋面', '台阶', '散水', '建筑面积']:
                    str3 = itemsList[row][keydict["名称"][0]]
                    str3 = re.sub(r'\[[^\[]*\]', '', str3)  # 【注释】 删除

                # print('str3', str3)
                # ~~~~~~~~~~~~~~特征描述~~~~~~~~~~~~~~~~1
                # str1=str2+str3  # 根据名称和特征描述 生成唯一编码，后期清单引用工程量使用
                # print('str4', str4)
                # print('num', type(num), num)
                # print('1str5', type(str5), str5)
                # print('1str6', str6)
                # floorname = itemsList[row][keydict["楼层"][0]]
                if str5 == '':  # 第一次执行这里
                    str5 = num
                    str6 = float(str5)
                    sumnum = float(str5)
                else:  # TODO 楼层、特征、项目名称（地上地下+简称）和上次不同时,计算表达式 归 "" str3 特征
                    if floorname != detailedlist[-1][0] and (
                            str3 == valuelist[-1][4] and (str2 + new) == valuelist[-1][3]):  # 特征 项目名称 同清单 换楼层
                        detailedlist.append([])  # 明细增加一行
                        # print('明细行增加了一行')
                        str5 = num  # str5 = '' 楼层名和上一次不同 初始化 表达式
                        str6 = float(num)
                        sumnum += float(num)
                    elif str3 != valuelist[-1][4] or (str2 + new) != valuelist[-1][3]:  # 特征 项目名称不同 换清单
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
                valuelist[-1] = [vaNo, '广联达绘图工程量_' + shtName, '', str2 + new,
                                 str3, str4, str(round(sumnum, 4)), detailedlist, '']  # 按清单表结构 新建


                # print('detailedlist, valuelist', detailedlist, valuelist)
            # print(valuelist)
            rowsList.extend(valuelist)  # 每个类型的工程量 增加至总表 二维表
        # ~~~~~~~~~循环量表的每一行~~~~~~~~~~~~~~1
        # print(rowsList)
        # 最后一行添加空行
        rowsList.extend([['', '', '', '', '', '', '', [['', '', '', '', '', '', '', '']], '']])

        # print(rowsList)
        return rowsList

    # 传入 数据模型 返回一个二维表 colkey第几列要输出 辅助  匹配模板
    def modleToList(self, modledata):  # -> list
        rows = modledata.rowCount()
        cols = modledata.columnCount()
        mList = []
        for row in range(rows):
            for col in range(cols):
                # if colListKey!=None:
                # if not col in colListKey:
                #     continue
                if col == 0:
                    mList.append([])
                temp = modledata.item(row, col).text()
                mList[row].append(temp)
        return mList

    # 测试用按钮
    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_butTest_clicked(self):  # 确定 按钮
        print('测试 按钮')
        self.accept()

    #  ==========自定义的功能函数================== 1

    ##  ==========由connectSlotsByName() 自动连接的槽函数==================
    # 确定按钮 （提取按钮）
    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_MaBtnOK_clicked(self):  # 确定 按钮
        # print('确定 按钮')
        # if not self.ui.lineEdit.text():
        #     # print("文件路径为空")
        #     return
        self.on_MaBtnSave_clicked()  # 保存当前设置表
        self.accept()  # TODO 确定退出窗口 .Accepted 会发生一个信号给主窗口
        # self.MaBtnOK.clicked.connect(QWDialogMatching.accept) # 系统自动生成的连接

    # # 选择工程量文件
    # @pyqtSlot(bool)
    # def on_pushButton_clicked(self):
    #     print('选择工程量文件 按钮')
    #     curPath = QDir.currentPath()  # 获取系统当前目录
    #     dlgTitle = "选择一个文件"  # 对话框标题
    #     filt = "算量软件导出excel表(*.xls);;所有文件(*.*)"  # 文件过滤器
    #     filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
    #     if filename:
    #         # self.setWindowTitle(str(filename))  # 设置窗口名称为 全路径
    #         self.ui.lineEdit.setText(filename)

    # 锁 复选框 锁定下拉菜单和 删除 按钮
    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_MacheckBoxLock_clicked(self, checked):
        self.ui.MacomboBoxModel.setEditable(not checked)  # 表格控件可编辑
        self.ui.MaBtnDel.setEnabled(not checked)  # 删除按钮 可按下

    # 插入行
    @pyqtSlot(bool)
    def on_MaBtnAppRow_clicked(self):
        cols = self.itemModel.columnCount()  # 先要获得列数
        itemlist = []  # QStandardItem 对象列表
        for i in range(cols):
            item = QStandardItem('')
            itemlist.append(item)
        curIndex = self.selectionModel.currentIndex()  # 获取当前选中项的模型索引
        # if curIndex.row() == -1:  # 如果未选中行，则在最后一行插入
        #     curIndex = self.itemModel.rowCount()

        self.itemModel.insertRow(curIndex.row() + 1, itemlist)  # 在当前行的下面插入一行

        self.selectionModel.clearSelection()
        self.selectionModel.setCurrentIndex(curIndex, QItemSelectionModel.Select)

    # 插入列
    @pyqtSlot(bool)
    def on_MaBtnAppCol_clicked(self):
        row = self.itemModel.rowCount()  # 先要获得列数
        itemlist = []  # QStandardItem 对象列表
        for i in range(row):  # 不包括最后一列
            item = QStandardItem('')
            itemlist.append(item)
        curIndex = self.selectionModel.currentIndex().column()  # 获取当前选中项的模型索引
        # if curIndex == -1:  # 如果未选中行，则在最后一行插入
        #     curIndex = self.itemModel.rowCount()

        self.itemModel.insertColumn(curIndex + 1, itemlist)  # 在当前行的前面插入一行

        self.selectionModel.clearSelection()
        self.selectionModel.setCurrentIndex(self.selectionModel.currentIndex(), QItemSelectionModel.Select)

    # 删除行
    @pyqtSlot(bool)
    def on_MaBtnDelRow_clicked(self):
        curIndex = self.selectionModel.currentIndex()  # 获取当前选择单元格的模型索引
        self.itemModel.removeRow(curIndex.row())  # 删除当前行

    # 删除列
    @pyqtSlot(bool)
    def on_MaBtnDelCol_clicked(self):
        curIndex = self.selectionModel.currentIndex()  # 获取当前选择单元格的模型索引
        self.itemModel.removeColumn(curIndex.column())  # 删除当前行

    # 删除模板
    @pyqtSlot(bool)
    def on_MaBtnDel_clicked(self):
        strkey = self.ui.MacomboBoxModel.currentText()
        if strkey == '默认设置':
            # print('默认设置无法删除')
            return
        if strkey in self.matchingList.keys():
            self.ui.MacomboBoxModel.removeItem(
                self.ui.MacomboBoxModel.currentIndex())  # TODO 下拉列表移除 先查找当前条目的索引 将索引返回给移除方法
            del self.matchingList[strkey]  # 字典中删除
            self.savepickle(self.matchingList)  # 存泡菜

    # 保存模板
    @pyqtSlot(bool)
    def on_MaBtnSave_clicked(self):
        listdata = []  # 嵌套列表形式
        keystr = self.ui.MacomboBoxModel.currentText()
        # print(keystr)
        rows = self.itemModel.rowCount()
        cols = self.itemModel.columnCount()

        for i in range(rows):  # 从0开始
            for j in range(cols):  # 从0开始
                item = self.itemModel.item(i, j).text()  # 获取每一项内的字符
                if j == 0:  # 如果是每一行的第一个数据 则添加一个空列表
                    listdata.append([])
                listdata[i].append(item)
        if not keystr in self.matchingList.keys():  # 如果模板名称不在字典中 增加下拉列表
            self.ui.MacomboBoxModel.insertItem(self.ui.MacomboBoxModel.count(), str(keystr))
        YorNUpDown = self.ui.MacomboBoxUpDown.currentText() # 是否区分地上地下
        Nokeys = list(self.ui.lineEditNoCal.text().split(","))  # 不提量关键字 list 类型
        self.matchingList['其他'] = {'区分地上地下': YorNUpDown, '不提量关键字': Nokeys}
        self.matchingList[keystr] = listdata  # 列表写入字典
        self.savepickle(self.matchingList)  # 存泡菜

    # 模板选择下拉框   “简单的ComboBox”的当前项变化  下拉框发生改变时 表格显示的变化
    @pyqtSlot(str)
    def on_MacomboBoxModel_currentIndexChanged(self, curText):
        # ~~~~~~~~表格控件初始化~~~~~~~~~
        print(curText)
        self.ui.MaTableWidget.setEnabled(True)  # 表格控件可编辑
        self.ui.MaBtnSave.setEnabled(True)  # 表格控件可编辑
        # strlist = self.ui.matching_CboxList.currentText()

        self.__RowCount = len(self.matchingList[curText])  # 列数=6  #
        if self.__RowCount < 1:  # 空数据列表则返回
            self.ui.MaTableWidget.setEnabled(False)  # 表格控件不可编辑
            self.ui.MaBtnSave.setEnabled(False)  # 表格控件不可编辑
            return
        self.__ColCount = len(self.matchingList[curText][0])  # 列数  #

        if self.__ColCount > 1:  # 空数据列表则返回

            # self.itemModel = QStandardItemModel(self.__RowCount, self.__ColCount, self)  # 数据模型,10行6列 #
            self.itemModel.setRowCount(self.__RowCount)  # TODO  数据模型,重置行数
            self.itemModel.setColumnCount(self.__ColCount)  # TODO 数据模型,重置列数
            headerList = ['工程量表名', ] + ['工程量名%d\n简称' % x for x in range(1, self.__ColCount)]
            self.itemModel.setHorizontalHeaderLabels(headerList)  # TODO 设置表头标题

            # ~~~~写入数据模型~~~~~
            for i in range(self.__RowCount):
                for j in range(self.__ColCount):
                    item = self.matchingList[curText][i][j]

                    item = QStandardItem(item)
                    font = item.font()
                    if j == 0:  # TODO  首列字符 粗体
                        font.setBold(True)
                        font.setPointSize(12)
                        colorstr = QColor(59, 142, 234)
                    else:
                        font.setBold(False)
                        font.setPointSize(10)
                        colorstr = QColor(0, 0, 0)  # 黑色
                    item.setFont(font)
                    item.setForeground(colorstr)
                    self.itemModel.setItem(i, j, item)  # 设置模型的item   X,Y ,value  从0开始

    ##  =============自定义槽函数===============================
    def do_curChanged(self, current):  # 表格内选择发生变化时触发
        if (current != None):  # 当前模型索引有效
            # print(current)
            text = "%d行；%d列" % (current.row() + 1, current.column() + 1)
            # text2 = "%s" % (self.itemModel.itemFromIndex(current).text())
            row = current.row()
            if row % 2:
                row = row - 1
            # print(row)
            text2 = f"当前工程量大类：{self.itemModel.item(row, 0).text()}"
            self.ui.Malabel_1.setText(str(text))
            self.ui.Malabel_2.setText(str(text2))

##  ============窗体测试程序 ================================
if __name__ == "__main__":  # 用于当前窗体测试
    app = QApplication(sys.argv)  # 创建GUI应用程序
    form = QmyDialogImportGlodon()  # 创建窗体
    form.show()
    sys.exit(app.exec_())
