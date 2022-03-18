# -*- coding: utf-8 -*-
import sys
import os
import time
import json
import pickle  # 泡菜模块
import socket
import re  # 导入正则表达式模块
import ast  # 列表包裹在引号内，提取出来变成列表  字符转列表 字符转字典等
import getpass  # 获取当前用户名
import platform  # 获取操作系统版本相关
import wmi  # 获取硬盘 cpu 主板 mac bios等硬件信息模块
import pymysql  # mysql服务器
# import csv
import xlrd, xlwt
import xlwings as xw
# import pandas
from pprint import pprint
# # 数据可视化分析的图表模块
# from pyecharts import options as opts
# from pyecharts.charts import Pie
# from pyecharts.faker import Faker
from dataVisualization import DataCal  # 生成数据可视化的html文件
# 导入自定义模块
import some_infor
# 导入PYQT5 相关模块
from PyQt5.QtWidgets import qApp,QApplication, QMainWindow, QUndoStack, QUndoCommand, \
    QMessageBox, QSpinBox, QLabel, QTableView, QCheckBox, QAbstractItemView, QHeaderView, \
    QColorDialog, QDialog, QInputDialog, QLineEdit, QWidget, QSizePolicy, QAction, QComboBox, \
    QFileDialog, QHBoxLayout, QFrame, QPushButton
from PyQt5.QtCore import pyqtSlot, pyqtSignal, Qt, QItemSelectionModel, QStringListModel, \
    QTimer, QTime, QDateTime, QThread, QDir, QUrl, QFileInfo
from PyQt5.QtGui import QFont, QColor, QPalette, QStandardItemModel, QStandardItem, QIcon, QPixmap
from PyQt5.QtWebEngineWidgets import QWebEngineView
##from PyQt5.QtSql import
##from PyQt5.QtMultimedia import
##from PyQt5.QtMultimediaWidgets import
# 载入对话框
from ui_MainWindow import Ui_MainWindow
from myDelegates import QmyComboBoxDelegate
from myDialogInformation import QmyDialoginformation
from myDialogQuantities import QmyDialogQuantities
from myDialogImportExcel import QmyDialogImportExcel
from myDialogLogin import QmyDialogLogin

# 用于继承撤销的类 用于undo redo
class CommandItemEdit(QUndoCommand):
    def __init__(self, connectSignals, disconnectSignals, model, item, textBeforeEdit, description="Item edited"):
        QUndoCommand.__init__(self, description)
        self.model = model
        self.item = item
        self.textBeforeEdit = textBeforeEdit
        self.textAfterEdit = item.text()
        self.connectSignals = connectSignals
        self.disconnectSignals = disconnectSignals

    def redo(self):
        # try:
        self.disconnectSignals()
        self.item.setText(self.textAfterEdit)
        self.connectSignals()
        # except Exception as e:
        #     print("发生了redo CommandItemEdit错误 ： ", e)

    def undo(self):
        # try:
        self.disconnectSignals()
        self.item.setText(self.textBeforeEdit)
        self.connectSignals()
        # except Exception as e:
        #     print("发生了undo CommandItemEdit错误 ： ", e)


# 用户获取用户信息，以及同步mysql
class Worker(QThread):
    advertise = pyqtSignal(list)  # 打开程序时 从服务器接收数据 元组形式 广而告之；advertise

    def __init__(self, parent=None):
        super(Worker, self).__init__(parent)
        self.working = True
        self.num = 0

    def __del__(self):
        self.working = False
        self.wait()

    def run(self):
        while self.working == True:
            print("开始执行开机任务支线！")
            diskstr = some_infor.get_diskNum()  # 硬盘信息的查询  返回硬盘号
            # print('硬盘号:', diskstr)
            self.user_infor_send(diskstr)  # 电脑硬件信息发送
            self.user_login_send(diskstr)  # 登录信息发送

            # path1 = r'D:\wxy\zhanghu.docx'
            # result = self.get_file_size(path1)
            # print(result)
            # result = self.get_file_create_time(path1)
            # print(result)
            # result = self.get_file_modify_time(path1)
            # print(result)
            # visit_time= self.get_file_visit_time(path1)
            # print(visit_time)
            # now_time = self.get_now_time()
            # print(now_time)
            # print(self.get_hostname())
            # print(self.get_user_name())
            # print("操作系统为：", platform.platform())
            # # print(self.get_Extranet_IP()[0])
            # # print(self.get_Extranet_IP()[1])
            #
            # c = wmi.WMI().Win32_DiskDrive()  # 多块硬盘 就是返回list类型, 非list就是一块硬盘
            # # c = c.SerialNumber  # 序列号
            # print(type(c), c)
            # somestr = "决定书反抗精神的"
            # # 线程休眠2秒
            # # self.sleep(2)
            # # 发出信号
            # self.advertise.emit(somestr)

            some_infor.advertise_infor_receive()  # 更新广告
            self.advertise_emit_main()  # 本地取广告 并定时发送给主程序 TODO 无限循环 放在最后

            self.working = False
    # TODO  ============文件处理 相关 ================================保存 打开 新建 另存 备份等
    # 获取文件大小
    def get_file_size(self, filepath):
        """
        获取文件大小，结果保留两位小数，单位MB
        """
        f = os.path.getsize(filepath)
        f = f / float(1024 * 1024)
        return round(f, 2)


    # 获取文件创建时间
    def get_file_create_time(self, filepath):
        """
        获取文件创建时间
        """
        # print("获取文件创建时间")
        tf = os.path.getctime(filepath)
        t = time.localtime(tf)
        # 时间戳转换方法
        return time.strftime('%Y-%m-%d %H:%M:%S', t)

    # 获取文件最新修改时间
    def get_file_modify_time(self, filepath):
        """
        获取文件修改时间
        """
        tf = os.path.getmtime(filepath)
        t = time.localtime(tf)
        # 时间戳转换方法
        return time.strftime('%Y-%m-%d %H:%M:%S', t)

    # 获取文件访问时间
    def get_file_visit_time(self, filepath):
        """
        获取文件访问时间
        """
        tf = os.path.getatime(filepath)
        t = time.localtime(tf)
        # 时间戳转换方法
        return time.strftime('%Y-%m-%d %H:%M:%S', t)


    # TODO  ============MySQL 数据库相关 ================================ 连接
    def user_infor_send(self, diskstr):  # 电脑硬件信息发送
        # print("电脑硬件信息发送")
        if not diskstr:
            return
        try:
            diskstr = str(diskstr)
            conn = pymysql.connect(host='rm-bp114m07t2e13i30i9o.mysql.rds.aliyuncs.com',
                                   port=3306, user='use_cqd000', password='Cqd123456', db='chengqingdan2021',
                                   charset='utf8')
            # status = conn.server_status
            cursor = conn.cursor()
            sql = '''
                select * from user_infor where disknum=%s;
            '''
            effect_row = cursor.execute(sql, [diskstr])
            # print("查到了数据 ", effect_row, "条！")
            cuu = cursor.fetchall()
            if effect_row:  # 如果查的到数据 说明有
                pass
            else:  # 写入数据
                sql = '''
                    INSERT INTO user_infor(disknum) VALUES (%s);
                '''
                row = cursor.execute(sql, [diskstr])
                conn.commit()  # 提交
                # print("写入", row)
        except pymysql.MySQLError as err:
            # 回滚事务
            # conn.rollback()
            print(f"出现错误：{err}")
        else:  # 此处不可改为 "finally 语句" 否则会报错！
            cursor.close()  # 4.关闭游标
            conn.close()  # 5.关闭连接

    def user_login_send(self, diskstr):  # 登录信息发送
        # print("登录信息发送")
        ip_ext_str = some_infor.get_Extranet_IP()[0]
        ip_int_str = some_infor.get_Intranet_IP()
        logintime = some_infor.get_now_time()
        hostname = some_infor.get_hostname()
        systemname = some_infor.get_system_username()
        system_ver = some_infor.get_system_Version()
        try:
            diskstr = str(diskstr)
            conn = pymysql.connect(host='rm-bp114m07t2e13i30i9o.mysql.rds.aliyuncs.com',
                                   port=3306, user='use_cqd000', password='Cqd123456', db='chengqingdan2021',
                                   charset='utf8')
            # status = conn.server_status
            cursor = conn.cursor()
            sql = '''
                INSERT INTO user_login(ip_ext, ip_int, logintime, hostname, systemname, system_ver, disknum) 
                values(%s, %s, %s, %s, %s, %s, %s)
            '''
            effect_row = cursor.execute(sql, (ip_ext_str,ip_int_str,logintime,hostname,systemname,system_ver,diskstr))
            # print("查到了数据 ", effect_row, "条！")
            cuu = cursor.fetchall()
            conn.commit()  # 提交
        except pymysql.MySQLError as err:
            # 回滚事务
            # conn.rollback()
            print(f"出现错误：{err}")
        else:  # 此处不可改为 "finally 语句" 否则会报错！
            cursor.close()  # 4.关闭游标
            conn.close()  # 5.关闭连接

    # 本地取广告 并发送给主程序
    def advertise_emit_main(self):
        # print('# 本地取广告 并发送给主程序')
        try:  #  encoding='utf-8'
            with open(r".\data\main_data\advertiseDict.db", "r", encoding=None) as f:
                advertiseDict = json.load(f)  # 字典 列表
            if advertiseDict:
                cuu = advertiseDict['advertise']  # 取出元组
            if cuu:
                rows = len(cuu)   # 一共有多少条数据
                while True:  # 无限循环
                    for row in range(rows):
                        linetuple = cuu[row]
                        continuedtime = linetuple[1]
                        self.advertise.emit(linetuple)
                        time.sleep(continuedtime)
        except Exception as e:
            print("取adv失败！", e)


class QmyMainWindow(QMainWindow):
    modelEmitFindWindow = pyqtSignal(object)  # 打开搜索对话框时把model      发送给find窗口

    def __init__(self, parent=None):
        super().__init__(parent)  # 调用父类构造函数，创建窗体
        self.ui = Ui_MainWindow()  # 创建UI对象
        self.ui.setupUi(self)  # 构造UI界面

        self.edition = '诚清单(T2.1.16)'  # 统一的版本号
        self.setWindowTitle(self.edition)  # 设置窗口名称
        self.setWindowState(Qt.WindowMaximized)  # 窗口最大化显示 全屏
        # self.setWindowFlags(Qt.WindowMinimizeButtonHint |  # 使能最小化按钮
        #                     Qt.WindowMaximizeButtonHint |  # 使能最大化按钮
        #                     Qt.WindowCloseButtonHint |  # 使能关闭按钮
        #                     Qt.WindowStaysOnTopHint)  # 窗体总在最前端
        # 用于多线程 用户信息
        self.thread = Worker()
        self.thread.advertise.connect(self.advertise_info)
        self.thread.start()

        # 选项卡控件设置
        # self.ui.tabWidget_1.tabBarClicked.connect(self.tabClicked_1)  # 选项卡被点击时触发 清单表 分析表
        self.ui.tabWidget_1.currentChanged.connect(self.tabClicked_1)  # 选项卡被点击时触发 清单表 分析表
        # self.ui.tabWidget_2.tabBarClicked.connect(self.tabClicked_2)  # 选项卡被点击时触发 面积 措施 其他 分布 分类表
        self.ui.tabWidget_2.currentChanged.connect(self.tabClicked_2)  # 选项卡被点击时触发 面积 措施 其他 分布 分类表

        self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 清单表页面
        self.ui.DAtab.setCurrentIndex(0)  # 分析系统 在 总览表
        self.ui.tabWidget_2.setCurrentIndex(3)  # 切换到tabindex 3 分部分项清单页面

        # 公共变量
        self.tabindex = "分部分项清单"  # 选项卡 中文名
        self.templist = []  # 复制粘贴行用的临时列表变量
        self.spinFontSize = 10  # 字体大小
        self.decimalPlaces = 3  # 小数位数
        self.beforestrundo = [0, 0]  # 改之前的原字符  用于记录历史操作的
        self.afterstrundo = ""  # 改字符之后的新字符  用于记录历史操作的
        self.filename = None  # 文件保存路径
        self.__NoCalFlags = (Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)  # 不计标志 复选框
        self.__NoCalTitle = ""  # 不计标志 复选框

        self.dlgInforObj = None  # 工程信息的对象 初始化为空
        self.dlgQuanObj = None  # 工程量的对象 初始化为空
        self.dlgLoginObj = None  # 用户登录的对象 初始化为空
        # 工程信息字典(一个工程公用)
        self.inforDict = {'项目名称': '', '所属事业部': '', '项目所在省份': '', '项目所在市': '', '地区类型': '',
                          '设计院名称': '', '人防设计院名称': ''}  #, '编制单位': '','编制人': '','编制日期': '', '记事本': ''}
        self.singleInforDict = {}  # 单体信息  写入房号表第二个单元格
        # 工程量表（每个房号独立）传给主窗口）
        self.quantitiesDict = {}
        # # 主工程文件字典
        self.engineeringDict = {'工程信息': self.inforDict}

        # 控件初始化
        self.__init_tableView_0()  # 房号表 初始化
        self.__init_singleInfor()  # 单体信息表 初始化
        self.__init_tableView_1()  # 清单表 面积表 初始化
        self.__init_tableView_2()  # 清单表 措施项目清单 初始化
        self.__init_tableView_31()  # 清单表 分部分项 初始化
        self.__init_tableView_32()  # 清单表 分部分项 明细表 初始化
        self.__init_listView_undo()  # 撤销栏 历史操作 初始化
        # self.__init_bulidNumList()  # 房号 初始化
        # self.__init_tableView_4()  # 分类表 初始化
        self.__buildUI()  # 动态创建组件，添加到工具栏和状态栏

        # 信号与函数连接
        self.connectAll()

        # 分割器设置
        self.ui.splitter_main.setStretchFactor(0, 8)  # 房号 与 清单表  垂直分割比例 （索引,百分比）
        self.ui.splitter_main.setStretchFactor(1, 92)

        self.ui.splitter_1.setStretchFactor(0, 50)  # 房号与 历史操作
        self.ui.splitter_1.setStretchFactor(1, 50)

        self.ui.splitter_3.setStretchFactor(0, 70)  # 清单表和明细表的 水平分割比例 （索引,百分比）
        self.ui.splitter_3.setStretchFactor(1, 30)

        # 初始不可用控件
        self.ui.tableView_32.setEnabled(False)  # 不可用
        self.ui.actEdit_Copy.setEnabled(False)
        self.ui.actEdit_Paste.setEnabled(False)
        self.ui.act_Redo.setEnabled(False)
        # self.ui.pushButton_addBuild.setEnabled(False)
        self.ui.pushButton_delBuild.setEnabled(False)
        # self.ui.actDataPrepare.setEnabled(False)

        # self.ui.label_33.setText('<a href="https://www.baidu.com">百度网址</a>')
        # self.ui.label_33.setOpenExternalLinks(True)
        # self.ui.label_33.setTextInteractionFlags(Qt.NoTextInteraction)
        # self.ui.label_33.setTextInteractionFlags(Qt.TextSelectableByKeyboard)
        self.ui.label_66.setToolTip("占地面积的说明：是指整个建筑物最大的水平投影面积。\n"
                                    "用于测定一些特定工程量的含量，如：筏板、屋面做法等！")  # 占地面积的说明

        # self.ui.undoView.setEnabled(False)
        # self.ui.undoView.setVisible(False)
        # self.ui.tab_1_4.setVisible(True)


    # TODO  ==============初始化功能函数========================
    # 非常重要 传入一个新建的数据模型 全部赋值为空白 否则无法获取text（）
    def initItemModelBlank(self, itemModel):
        rows = itemModel.rowCount()
        cols = itemModel.columnCount()
        for x in range(rows):
            for y in range(cols):
                item = QStandardItem("")
                itemModel.setItem(x, y, item)  # 初始化表 itemModel 每项目为空

    def __buildUI(self):  # 窗体上动态添加组件     创建状态栏上的组件
        # 菜单栏构建
        # '''字体大小 调节按钮 '''
        self.__labelSize = QLabel(self)  # 调整字体大小文字
        self.__labelSize.setText(" 字大小：")
        self.ui.mainToolBar.addWidget(self.__labelSize)  # 添加到工具栏
        self.__spinFontSize = QSpinBox(self)  # 调整字体大小按钮到工具栏
        self.__spinFontSize.setMinimum(8)
        self.__spinFontSize.setMaximum(35)
        self.__spinFontSize.setValue(self.spinFontSize)  # 默认字号大小
        self.__spinFontSize.valueChanged.connect(self.spinFontSize_valueChanged)  # 变化时触发
        self.ui.mainToolBar.addWidget(self.__spinFontSize)  # 添加到工具栏
        # '''小数位数 调节按钮'''
        self.__labelSize = QLabel(self)  # 调整字体大小文字
        self.__labelSize.setText(" 小数位：")
        self.ui.mainToolBar.addWidget(self.__labelSize)  # 添加到工具栏
        self.__decimalPlaces = QSpinBox(self)  # 调整字体大小按钮到工具栏
        self.__decimalPlaces.setMinimum(2)
        self.__decimalPlaces.setMaximum(4)
        self.__decimalPlaces.setValue(self.decimalPlaces)  # 默认小数位数
        self.__decimalPlaces.valueChanged.connect(self.decimalPlaces_valueChanged)  # 变化时触发
        self.ui.mainToolBar.addWidget(self.__decimalPlaces)  # 添加到工具栏
        # 登录 注册框
        self.spacerWidget = QWidget(self)  # 继承一个间隔控件
        self.spacerWidget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        self.ui.mainToolBar.addWidget(self.spacerWidget)

        self.userlogin = QAction(QIcon(QPixmap(":/icons/images/Login.png")), "", self)
        self.userlogin.setToolTip("登录")
        self.ui.mainToolBar.addAction(self.userlogin)
        self.userlogin.triggered.connect(self.userlogin_triggered)
        # self.ui.mainToolBar.setToolButtonStyle(Qt.ToolButtonIconOnly)
        # self.__username = QLineEdit(self)
        # self.__username =
        # self.__username.setToolTip("输入用户名")  # 提示符
        # self.__username.setPlaceholderText("输入用户名")  # 占位符
        # self.__username.setFixedWidth(80)
        # self.ui.mainToolBar.addWidget(self.__username)

        # 状态栏构建
        self.__LabFile = QLabel(self)  # QLabel组件显示信息
        self.__LabFile.setMinimumWidth(200)  # 宽度
        self.__LabFile.setText("姿势正确！")
        self.ui.statusBar.addWidget(self.__LabFile)  # 状态栏 从左至右添加

        self.__LabFile2 = QLabel(self)  # QLabel组件显示信息
        # self.__LabFile2.setMinimumWidth(150)
        self.__LabFile2.setMaximumWidth(1800)
        self.__LabFile2.setText("这里是自己构建的状态栏2")
        self.ui.statusBar.addWidget(self.__LabFile2)  # 状态栏 从左至右添加

        self.__LabInfo = QLabel(self)  # QLabel组件显示信息
        # self.__LabInfo.setMinimumWidth(150)
        self.__LabInfo.setText("公众号《王欣阳小课堂》、QQ群 765198205")
        self.ui.statusBar.addPermanentWidget(self.__LabInfo)  # 状态栏最右边

    # 单体信息表 初始化
    def __init_singleInfor(self):
        formlist = ['', '别墅', '叠拼', '联排', '合院', '洋房', '小高层', '高层', '超高层', '单层非人防车库', '双层非人防车库',
                    '单层人防车库', '双层人防车库', '停车楼', '附属商业(沿街商业)', '大商业', '金街', '外街', 'SOHO办公',
                    '酒店式公寓', '幼儿园', '小学', '配套公共设施']
        self.ui.singleCBBox_1.addItems(formlist)
        # self.singleInforDict = {}  # 单体信息  写入房号表第二个单元格 singleInforSet函数 传参# str1 字典key str2 value
        self.ui.singleCBBox_1.currentIndexChanged.connect\
            (lambda: self.singleInforSet("业态", self.ui.singleCBBox_1.currentText()))

        seismicGrade = ['','特一级','一级','二级','三级','四级','非抗震']
        self.ui.singleCBBox_2.addItems(seismicGrade)
        self.ui.singleCBBox_2.currentIndexChanged.connect\
            (lambda: self.singleInforSet("抗震等级", self.ui.singleCBBox_2.currentText()))

        fortificationIntensity = ['','5','6','7','8','9']
        self.ui.singleCBBox_3.addItems(fortificationIntensity)
        self.ui.singleCBBox_3.currentIndexChanged.connect\
            (lambda: self.singleInforSet("设防烈度", self.ui.singleCBBox_3.currentText()))

        basicsType = ['','筏板基础','基础梁基础','独立基础','条线基础','承台+地梁基础','不包含基础']  # 基础类型
        self.ui.singleCBBox_4.addItems(basicsType)
        self.ui.singleCBBox_4.currentIndexChanged.connect\
            (lambda: self.singleInforSet("基础类型", self.ui.singleCBBox_4.currentText()))

        roofingType = ['','平屋面','坡屋面']  # 屋面类型
        self.ui.singleCBBox_5.addItems(roofingType)
        self.ui.singleCBBox_5.currentIndexChanged.connect\
            (lambda: self.singleInforSet("屋面形式", self.ui.singleCBBox_5.currentText()))

        structureType = ['','框架结构','剪力墙结构','剪力墙（全砼外墙）','砖混']  # 结构类型
        self.ui.singleCBBox_6.addItems(structureType)
        self.ui.singleCBBox_6.currentIndexChanged.connect\
            (lambda: self.singleInforSet("结构类型", self.ui.singleCBBox_6.currentText()))

        self.ui.spinBox_1.valueChanged.connect\
            (lambda: self.singleInforSet("首层层高", self.ui.spinBox_1.value()))

        self.ui.spinBox_2.valueChanged.connect\
            (lambda: self.singleInforSet("标准层高", self.ui.spinBox_2.value()))

        self.ui.spinBox_3.valueChanged.connect\
            (lambda: self.singleInforSet("地上层数", self.ui.spinBox_3.value()))

        self.ui.spinBox_4.valueChanged.connect\
            (lambda: self.singleInforSet("地下层数", self.ui.spinBox_4.value()))

        self.ui.spinBox_5.valueChanged.connect\
            (lambda: self.singleInforSet("PC率", self.ui.spinBox_5.value()))

        self.ui.spinBox_6.valueChanged.connect\
            (lambda: self.singleInforSet("新三板率", self.ui.spinBox_6.value()))

    # 房号表 初始化
    def __init_tableView_0(self):
        self.itemModel0 = QStandardItemModel(1, 12, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel0)  # 初始化 数据模型为 空值
        # 2021/11/21 修改 增加绘图工程量表2张
        headerList = ['房号', '单体信息表', '建筑面积计算表', '措施项目清单', '分部分项清单',
                      '工程量总表\n历史操作', '工程量总表', '分类表', '钢筋工程量文件表', '钢筋工程量表', '绘图工程量文件表', '绘图工程量表']
        self.tableView_0_headerList = headerList
        self.itemModel0.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel0 = QItemSelectionModel(self.itemModel0)  # itemModel 选择模型

        self.ui.tableView_0.setModel(self.itemModel0)  # 设置数据模型
        self.ui.tableView_0.setSelectionModel(self.selectionModel0)  # 设置选择模型
        # self.undoStack = QUndoStack(self)  # 用于undo redo 堆栈

        # tb=QTableView  # 表头换行
        self.ui.tableView_0.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_0.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        # self.ui.tableView_1.horizontalHeader().setHighlightSections(False)
        # self.ui.tableView_1.horizontalHeader().setStyleSheet(
        #     "QHeaderView::section{background-color:rgb(155, 194, 230);font:11pt '宋体';color: black;};")

        # 选择模式
        # oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        oneOrMore = QAbstractItemView.SingleSelection
        self.ui.tableView_0.setSelectionMode(oneOrMore)  # 可多选
        #
        # itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        # self.ui.tableView_0.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_0.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        # self.ui.tableView_0.verticalHeader().setDefaultSectionSize(60)  # 缺省行高 设置了行高自动 就失效了
        self.ui.tableView_0.setAlternatingRowColors(True)  # 交替行颜色
        # 设置表头边框样式
        self.ui.tableView_0.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"
            "QTableView::item{selection-background-color:#3399FF}")

        self.ui.tableView_0.setColumnWidth(0, 80)

        # self.ui.tableView_0.setColumnWidth(1, 200)
        # self.ui.tableView_0.setColumnWidth(2, 120)
        # self.ui.tableView_0.setColumnWidth(3, 100)
        # self.ui.tableView_0.setColumnWidth(4, 100)
        # self.ui.tableView_0.setColumnWidth(5, 100)
        # self.ui.tableView_0.setColumnWidth(6, 100)
        # self.ui.tableView_0.setColumnWidth(7, 100)
        # self.ui.tableView_0.setColumnWidth(8, 100)
        # self.ui.tableView_0.setColumnWidth(9, 100)
        # self.ui.tableView_0.setColumnWidth(10, 100)
        # self.ui.tableView_0.setColumnWidth(11, 100)
        ## 隐藏 除房号 以外的所有列
        self.ui.tableView_0.horizontalHeader().hideSection(1)
        self.ui.tableView_0.horizontalHeader().hideSection(2)
        self.ui.tableView_0.horizontalHeader().hideSection(3)
        self.ui.tableView_0.horizontalHeader().hideSection(4)
        self.ui.tableView_0.horizontalHeader().hideSection(5)
        self.ui.tableView_0.horizontalHeader().hideSection(6)
        self.ui.tableView_0.horizontalHeader().hideSection(7)
        self.ui.tableView_0.horizontalHeader().hideSection(8)
        self.ui.tableView_0.horizontalHeader().hideSection(9)
        self.ui.tableView_0.horizontalHeader().hideSection(10)
        self.ui.tableView_0.horizontalHeader().hideSection(11)
        # 设置默认房号
        item = QStandardItem('1#')
        self.itemModel0.setItem(0, 0, item)
        # 选中第一个为 默认
        self.selectionModel0.setCurrentIndex(self.itemModel0.index(0, 0), QItemSelectionModel.Select)

    # 建筑面积表 初始化
    def __init_tableView_1(self):
        self.itemModel1 = QStandardItemModel(3, 11, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel1)  # 初始化 数据模型为 空值
        headerList = ['序号', '地上\n地下', '楼层号', '层数',
                      '计算表达式\n中文括号【注释】\t\t\t\t尖括号<楼层>引用\t\t\t\t书名号《工程量表》引用',
                      '计算结果', '面积', '不计\n标志', '公式\n错误', '备注', '临时存放待处理表达式']  # 计算结果temp 用于暂存去除【注释】的结果
        self.tableView_1_headerList = headerList
        self.itemModel1.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel1 = QItemSelectionModel(self.itemModel1)  # itemModel 选择模型

        self.ui.tableView_1.setModel(self.itemModel1)  # 设置数据模型
        self.ui.tableView_1.setSelectionModel(self.selectionModel1)  # 设置选择模型
        # self.undoStack = QUndoStack(self)  # 用于undo redo 堆栈

        # tb=QTableView  # 表头换行
        self.ui.tableView_1.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_1.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        # self.ui.tableView_1.horizontalHeader().setHighlightSections(False)
        # self.ui.tableView_1.horizontalHeader().setStyleSheet(
        #     "QHeaderView::section{background-color:rgb(155, 194, 230);font:11pt '宋体';color: black;};")

        # 选择模式
        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_1.setSelectionMode(oneOrMore)  # 可多选
        #
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_1.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_1.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_1.setAlternatingRowColors(True)  # 交替行颜色

        # 设置表头边框样式
        self.ui.tableView_1.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"
            "QTableView::item{selection-background-color:#3399FF}")
        self.ui.tableView_1.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        # self.ui.tableView_1.setStyleSheet("selection-background-color:lightBlue")   # 单元格选中变色  且光标离开保留颜色

        self.ui.tableView_1.setColumnWidth(0, 40)
        self.ui.tableView_1.setColumnWidth(1, 40)
        self.ui.tableView_1.setColumnWidth(2, 100)
        self.ui.tableView_1.setColumnWidth(3, 30)
        # self.ui.tableView_1.horizontalHeader().hideSection(3)
        self.ui.tableView_1.setColumnWidth(4, 700)
        self.ui.tableView_1.setColumnWidth(5, 80)
        self.ui.tableView_1.setColumnWidth(6, 80)
        self.ui.tableView_1.setColumnWidth(7, 40)
        self.ui.tableView_1.setColumnWidth(8, 40)
        self.ui.tableView_1.setColumnWidth(9, 200)
        self.ui.tableView_1.horizontalHeader().hideSection(10)

        # 不计标志
        for i in range(self.itemModel1.rowCount()):
            item = QStandardItem(self.__NoCalTitle)  # 最后一列
            item.setFlags(self.__NoCalFlags)
            item.setCheckable(True)  # 非锁定
            item.setCheckState(Qt.Unchecked)  # 非勾选
            self.itemModel1.setItem(i, 7, item)  # 设置最后一列的item

        # 地上地下 下拉
        qualities = ["", "地上", "地下"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_1.setItemDelegateForColumn(1, self.UnitOfMeasurement)  # 地上地下

    # 措施项目清单 初始化
    def __init_tableView_2(self):
        self.itemModel2 = QStandardItemModel(1, 11, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel2)  # 初始化 数据模型为 空值
        headerList = ['序号', '措施费\n类别', '项目', '部位\n内容', '计量\n单位',
                      '计算表达式\n书名号《工程量表》引用', '工程量', '除税综合单价\n（元/m2)', '含税综合单价\n(元/m2)', '合价（元）', '备注']  # 计算结果temp 用于暂存去除【注释】的结果
        self.tableView_2_headerList = headerList
        self.itemModel2.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel2 = QItemSelectionModel(self.itemModel2)  # itemModel 选择模型
        self.ui.tableView_2.setModel(self.itemModel2)  # 设置数据模型
        self.ui.tableView_2.setSelectionModel(self.selectionModel2)  # 设置选择模型

        # tb=QTableView  # 表头换行
        self.ui.tableView_2.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_2.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_2.setSelectionMode(oneOrMore)  # 可多选
        #
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_2.setSelectionBehavior(itemOrRow)  # 单元格选择

        self.ui.tableView_2.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_2.setAlternatingRowColors(True)  # 交替行颜色

        # 设置表头边框样式
        self.ui.tableView_2.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"
            "QTableView::item{selection-background-color:#3399FF}")
        self.ui.tableView_2.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        # self.ui.tableView_1.setStyleSheet("selection-background-color:lightBlue")   # 单元格选中变色  且光标离开保留颜色

        self.ui.tableView_2.setColumnWidth(0, 40)
        self.ui.tableView_2.setColumnWidth(1, 100)
        self.ui.tableView_2.setColumnWidth(2, 200)
        self.ui.tableView_2.setColumnWidth(3, 200)
        # self.utableView_2_1.horizontalHeader().hideSection(3)
        self.ui.tableView_2.setColumnWidth(4, 40)
        self.ui.tableView_2.setColumnWidth(5, 200)
        self.ui.tableView_2.setColumnWidth(6, 80)
        self.ui.tableView_2.setColumnWidth(7, 100)
        self.ui.tableView_2.setColumnWidth(8, 100)
        self.ui.tableView_2.setColumnWidth(9, 100)
        self.ui.tableView_2.setColumnWidth(10, 200)

        category = ['', '整体措施费', '单项措施费', '其他项目清单']  # 类别
        self.comboDelegate = QmyComboBoxDelegate(self)
        self.comboDelegate.setItems(category, False)  # 不可编辑
        self.ui.tableView_2.setItemDelegateForColumn(1, self.comboDelegate)

        project = ['', '模板费用', '脚手架工程', '垂直运输', '总承包服务费', '外立面配合费', '如此项目为户内精装修项目']  # 项目
        self.comboDelegate = QmyComboBoxDelegate(self)
        self.comboDelegate.setItems(project, True)  # 不可编辑 分部分项清单合计
        self.ui.tableView_2.setItemDelegateForColumn(2, self.comboDelegate)

        # 计量单位 下拉
        qualities = ["", "m", "m2", "m3", "T", "kg", "个", "项"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_2.setItemDelegateForColumn(4, self.UnitOfMeasurement)  # 计量单位

        # 设置几行通用的
        # self.itemModel2.setItem(0, 0, QStandardItem('1'))
        # self.itemModel2.setItem(0, 1, QStandardItem('整体措施费'))
        # self.itemModel2.setItem(0, 4, QStandardItem('m2'))

        # item = QStandardItem('二')
        # self.itemModel2.setItem(2, 0, item)
        # item = QStandardItem('单项措施费')
        # self.itemModel2.setItem(2, 1, item)
        # item = QStandardItem('模板费用')
        # self.itemModel2.setItem(2, 2, item)
        # item = QStandardItem('单项措施费')
        # self.itemModel2.setItem(3, 1, item)
        # item = QStandardItem('脚手架工程')
        # self.itemModel2.setItem(3, 2, item)
        # item = QStandardItem('单项措施费')
        # self.itemModel2.setItem(4, 1, item)
        # item = QStandardItem('垂直运输')
        # self.itemModel2.setItem(4, 2, item)
        #
        # item = QStandardItem('三')
        # self.itemModel2.setItem(6, 0, item)
        # item = QStandardItem('其他项目清单')
        # self.itemModel2.setItem(6, 1, item)
        # item = QStandardItem('总承包服务费')
        # self.itemModel2.setItem(6, 2, item)

    # 分部分项表 初始化
    def __init_tableView_31(self):
        self.itemModel31 = QStandardItemModel(3, 20, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel31)  # 初始化 数据模型为 空值
        headerList = ['序号', '色标', '地上\n地下', '清单编码', '项目名称', '项目特征描述', '计量\n单位', '工程量',
                      '∑工程量明细表', '综合单价', '项目合价', '备注',
                      '人工费', '主材费', '辅材费', '机械费', '管理费、利润', '规费', '不含税综合单价', '增值税税金'
                      ]  # '∑\n工程量明细表' 需隐藏的列
        self.tableView_31_headerList = headerList
        self.itemModel31.setHorizontalHeaderLabels(headerList)  # 设置表头标题
        self.selectionModel31 = QItemSelectionModel(self.itemModel31)  # itemModel 选择模型

        # self.ui.act_Undo.setEnabled(False)   # 禁用撤销按钮
        # self.ui.act_Redo.setEnabled(False)   # 禁用恢复按钮
        # self.ui.tableView_31.clicked.connect(self.itemClicked)  # 用于undo redo
        # self.itemModel31.itemChanged.connect(self.itemChanged)   # 用于undo redo

        # self.undoStack = QUndoStack(self)  # 用于undo redo
        # undoView = QUndoView(self.undoStack)  # 用于undo redo
        # self.textBeforeEdit = ""  # 用于undo redo
        # self.UndoNumbers = 0  # 可撤销的次数 初始化

        # self.itemModel31.itemChanged.connect(self.do_cellChanged)  #

        self.ui.tableView_31.setModel(self.itemModel31)  # 设置数据模型
        self.ui.tableView_31.setSelectionModel(self.selectionModel31)  # 设置选择模型
        # self.undoStack = QUndoStack(self)
        # self.ui.undoView.setStack(self.undoStack)
        # self.ui.tableView_31=QTableView  # 表头字符换行
        # tb.verticalHeader()
        # self.ui.tableView_31.horizontalHeader().setHighlightSections(True)  # 选择的表头高亮
        self.ui.tableView_31.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))  # 表头字符对齐 换行
        # self.ui.tableView_31.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        # self.ui.tableView_31.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # 这个实现各列平均分配，并且占满整个tableview
        # self.ui.tableView_31.resizeRowsToContents()
        # self.ui.tableView_31.setWordWrap(True)  # 此设置为默认，需要调整行高

        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_31.setSelectionMode(oneOrMore)  # 可多选
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_31.setSelectionBehavior(itemOrRow)  # 单元格选择

        # self.ui.tableView_31.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.ui.tableView_31.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        # self.ui.tableView_31.verticalHeader().setDefaultSectionSize(58)  # 缺省行高 设置了行高自动 就失效了
        self.ui.tableView_31.setAlternatingRowColors(True)  # 交替行颜色
        # 设置表头边框样式
        # self.ui.tableView_31.setStyleSheet(
        #     "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9}")
        # self.ui.tableView_31.setStyleSheet("QHeaderView::section, QTableCornerButton::section{border:1px solid #014F84}""QTableView::item{border:1px solid #014F84}")

        # self.ui.tableView_31.setStyleSheet("QTableView::item:alternate:!selected{ background:lightBlue}")   # 交替行换色
        # self.ui.tableView_31.setStyleSheet("selection-background-color:lightgreen")  # 单个单元格 # 离开状态有停留
        # self.ui.tableView_31.setStyleSheet("selection-background-color:red")

        # self.ui.tableView_31.setSelectionBehavior(QTableView.SelectColumns)  # 单击某个项目时,将选择整个列
        # self.ui.tableView_31.setStyleSheet("selection-background-color:lightBlue")   # 单元格选中变色  且光标离开保留颜色
        # self.ui.tableView_31.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        # self.ui.tableView_31.setSelectionBehavior(QTableView.SelectColumns)  # 单击某个项目时,将选择整个列
        self.ui.tableView_31.setSelectionBehavior(QAbstractItemView.SelectRows)  # 单击某个项目时,将选择整个列
        # self.ui.tableView_31.setSelectionBehavior(QAbstractItemView.SelectColumns)  # 单击某个项目时,将选择整个列
        # self.ui.tableView_31.setStyleSheet(
        #     "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9}"  # #A9A9A9
        #     "QTableView::item{selection-background-color:#3399FF}")  # 表格颜色样式与 选中行 反色样式 lightBlue darkcyan #3399FF
        # self.ui.tableView_31.setStyleSheet(
        #     "QHeaderView::section{background-color:rgb(245, 245, 245);font:11pt 'Microsoft YaHei';color: black;}, "
        #     "QTableCornerButton::section{border:1px solid #A9A9A9}"  # #A9A9A9
        #     "QTableView::item{selection-background-color:#3399FF}")  # 表格颜色样式与 选中行 反色样式 lightBlue darkcyan #3399FF
        self.ui.tableView_31.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"  # #A9A9A9
            "QTableView::item{selection-background-color:#3399FF}")  # 表格颜色样式与 选中行 反色样式 lightBlue darkcyan #3399FF

        self.ui.tableView_31.setColumnWidth(0, 150)
        self.ui.tableView_31.setColumnWidth(1, 40)
        self.ui.tableView_31.setColumnWidth(2, 40)
        self.ui.tableView_31.setColumnWidth(3, 100)
        self.ui.tableView_31.setColumnWidth(4, 200)
        self.ui.tableView_31.setColumnWidth(5, 400)
        self.ui.tableView_31.setColumnWidth(6, 40)
        self.ui.tableView_31.setColumnWidth(7, 100)
        # self.ui.tableView_31.setColumnWidth(8, 200)
        self.ui.tableView_31.horizontalHeader().hideSection(8)
        self.ui.tableView_31.setColumnWidth(9, 100)
        self.ui.tableView_31.setColumnWidth(10, 100)
        self.ui.tableView_31.setColumnWidth(11, 200)
        # 以下人工到税金列的宽度 有 复选框 折叠决定  def on_checkBoxFoldPrice_clicked
        morenColunmzero = 0
        self.ui.tableView_31.setColumnWidth(12, morenColunmzero)  # 人工
        self.ui.tableView_31.setColumnWidth(13, morenColunmzero)  #
        self.ui.tableView_31.setColumnWidth(14, morenColunmzero)  #
        self.ui.tableView_31.setColumnWidth(15, morenColunmzero)  #
        self.ui.tableView_31.setColumnWidth(16, morenColunmzero)  #
        self.ui.tableView_31.setColumnWidth(17, morenColunmzero)  #
        self.ui.tableView_31.setColumnWidth(18, morenColunmzero)  #
        self.ui.tableView_31.setColumnWidth(19, morenColunmzero)  #

        # self.ui.tableView_31.setEditable(False)  # 下拉列表可编辑
        # self.ui.tableView_31.setEnabled(False)  # 表格控件可编辑

        # ~~~~~~~~~~~~~~创建自定义代理组件并设置~~~~~~~~~~~~

        # self.Quantity = QmyFloatSpinDelegate(-100000, 1000000, 3, self)  # 用于工程量 最小值 最大值 精度
        # self.ui.tableView_31.setItemDelegateForColumn(5, self.Quantity)
        #
        # self.price = QmyFloatSpinDelegate(-100000, 1000000, 2, self)  # 用于价格 最小值 最大值 精度
        # self.ui.tableView_31.setItemDelegateForColumn(6, self.price)

        # 计量单位 下拉
        qualities = ["", "m", "m2", "m3", "T", "kg", "个", "项", "座", "樘"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_31.setItemDelegateForColumn(6, self.UnitOfMeasurement)  # 计量单位
        # 色标 下拉
        self.rowcolor = ["", "1,红", "2,蓝", "3,黄", "4,绿"]
        self.UnitOfMeasurementrowcolor = QmyComboBoxDelegate(self)
        self.UnitOfMeasurementrowcolor.setItems(self.rowcolor, False)  # 可编辑
        self.ui.tableView_31.setItemDelegateForColumn(1, self.UnitOfMeasurementrowcolor)  # 色标
        # 地上地下 下拉
        qualities = ["", "地上", "地下"]
        self.UnitOfMeasurement = QmyComboBoxDelegate(self)
        self.UnitOfMeasurement.setItems(qualities, False)  # 不可编辑
        self.ui.tableView_31.setItemDelegateForColumn(2, self.UnitOfMeasurement)  # 地上地下

    # 分部分项明细表 初始化
    def __init_tableView_32(self):
        self.itemModel32 = QStandardItemModel(3, 10, self)  # 数据模型,初始化 不定义的话 会出错
        self.initItemModelBlank(self.itemModel32)  # 初始化 数据模型为 空值
        # 计算结果temp 用于暂存去除【注释】的结果
        headerList = ['楼层号', '部位', '计算表达式\n中文括号【注释】\t\t\t\t尖括号<楼层>引用\t\t\t\t书名号《工程量表》引用',
                      '计算结果real', '计算结果', '小计\n部 位', '小计\n楼 层', '不计\n标志', '公式\n错误', '备注']
        self.itemModel32.setHorizontalHeaderLabels(headerList)  # 设置表头标题

        self.selectionModel32 = QItemSelectionModel(self.itemModel32)  # itemModel 选择模型

        self.ui.tableView_32.setModel(self.itemModel32)  # 设置数据模型
        self.ui.tableView_32.setSelectionModel(self.selectionModel32)  # 设置选择模型

        # tb=QTableView  # 表头换行
        self.ui.tableView_32.horizontalHeader().setDefaultAlignment(
            Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.tableView_32.horizontalHeader().setStretchLastSection(True)  # 最后一列动态拉伸
        oneOrMore = QAbstractItemView.ExtendedSelection  # 选择模式
        self.ui.tableView_32.setSelectionMode(oneOrMore)  # 可多选
        #
        itemOrRow = QAbstractItemView.SelectItems  # 项选择模式
        self.ui.tableView_32.setSelectionBehavior(itemOrRow)  # 单元格选择

        # self.ui.tableView_32.verticalHeader().setDefaultSectionSize(28)  # 缺省行高
        self.ui.tableView_32.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents)  # 行高自动变 内容换行
        self.ui.tableView_32.setAlternatingRowColors(True)  # 交替行颜色

        # self.ui.tableView_32.setStyleSheet("QTableView{border:1px solid #014F84}")

        # 设置表头边框样式
        self.ui.tableView_32.setStyleSheet(
            "QHeaderView::section, QTableCornerButton::section{border:1px solid #A9A9A9;font:11pt}"
            "QTableView::item{selection-background-color:#3399FF}")
        self.ui.tableView_32.setSelectionBehavior(QTableView.SelectRows)  # 单击某个项目时,将选择整个行
        # self.ui.tableView_32.setStyleSheet("selection-background-color:lightBlue")   # 单单元格选中变色  且光标离开保留颜色

        self.ui.tableView_32.setColumnWidth(0, 80)
        self.ui.tableView_32.setColumnWidth(1, 160)
        self.ui.tableView_32.setColumnWidth(2, 600)
        # self.ui.tableView_32.setColumnWidth(3, 80)  # 结算结果的real  后期 隐藏
        self.ui.tableView_32.horizontalHeader().hideSection(3)  #
        self.ui.tableView_32.setColumnWidth(4, 80)  # 计算结果
        self.ui.tableView_32.setColumnWidth(5, 100)
        self.ui.tableView_32.setColumnWidth(6, 100)
        self.ui.tableView_32.setColumnWidth(7, 40)  # 不计  列
        self.ui.tableView_32.setColumnWidth(8, 40)  # 错误  列
        self.ui.tableView_32.setColumnWidth(9, 180)  # 备注  列
        for i in range(self.itemModel32.rowCount()):
            item = QStandardItem(self.__NoCalTitle)  # 最后一列
            item.setFlags(self.__NoCalFlags)
            item.setCheckable(True)  # 非锁定
            item.setCheckState(Qt.Unchecked)  # 非勾选
            self.itemModel32.setItem(i, 7, item)  # 设置最后一列的item

    # 历史操作 初始化
    def __init_listView_undo(self):
        self.itemModel_listView_undo = QStringListModel(self)
        # lll = ["1", "2", "3"]
        # self.itemModel_listView_undo.setStringList(lll)
        self.ui.listView_undo.setModel(self.itemModel_listView_undo)
        trig=(QAbstractItemView.DoubleClicked |QAbstractItemView.SelectedClicked)
        self.ui.listView_undo.setEditTriggers(trig)

    # 房号 初始化
    # def __init_bulidNumList(self):
    #     pass
        # self.modelBulidNumList = QStringListModel(self)
        # self.modelBulidNumList.setStringList(self.defaultbuildNumList)
        # self.ui.listView.setModel(self.modelBulidNumList)
        # self.ui.listView.setStyleSheet("QListView::item:selected {background-color:#3399FF}") # 设置选中项的背景色

    # 信号与函数连接
    def connectAll(self):
        self.selectionModel0.currentChanged.connect(self.selectionModel0_currentChanged)
        self.itemModel0.itemChanged.connect(self.itemModel0_itemChanged)  # 房号表的数据有变化，则触发该槽

        self.selectionModel1.currentChanged.connect(self.selectionModel1_currentChanged)
        self.itemModel1.itemChanged.connect(self.itemModel1_itemChanged)  # 面积表的数据有变化，则触发该槽

        self.selectionModel2.currentChanged.connect(self.selectionModel2_currentChanged)  # 主表 选择变化时 明细表触发 改变行数，显示内容
        self.itemModel2.itemChanged.connect(self.itemModel2_itemChanged)  # 主表有变化时触发

        self.selectionModel31.currentChanged.connect(self.selectionModel31_currentChanged)  #
        self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged)  # 主表有变化时触发

        self.ui.tableView_31.clicked.connect(self.tableView_31_clicked)  # 分部分项被点击

        self.selectionModel32.currentChanged.connect(self.selectionModel32_currentChanged)  # 附表被点击时触发
        self.itemModel32.itemChanged.connect(self.itemModel32_itemChanged)  # 明细表的数据有变化，则触发该槽



    def disconnectAll(self):
        self.selectionModel0.currentChanged.disconnect(self.selectionModel0_currentChanged)
        self.itemModel0.itemChanged.disconnect(self.itemModel0_itemChanged)  # 房号表的数据有变化，则触发该槽

        self.selectionModel1.currentChanged.disconnect(self.selectionModel1_currentChanged)
        self.itemModel1.itemChanged.disconnect(self.itemModel1_itemChanged)  # 面积表的数据有变化，则触发该槽

        self.selectionModel2.currentChanged.disconnect(self.selectionModel2_currentChanged)  # 主表 选择变化时 明细表触发 改变行数，显示内容
        self.itemModel2.itemChanged.disconnect(self.itemModel2_itemChanged)  # 主表有变化时触发

        self.selectionModel31.currentChanged.disconnect(self.selectionModel31_currentChanged)  #
        self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)  # 主表有变化时触发

        self.ui.tableView_31.clicked.disconnect(self.tableView_31_clicked)  # 分部分项被点击

        self.selectionModel32.currentChanged.disconnect(self.selectionModel32_currentChanged)  # 附表被点击时触发
        self.itemModel32.itemChanged.disconnect(self.itemModel32_itemChanged)  # 明细表的数据有变化，则触发该槽

    # TODO  ==============Undo Redo 功能函数==========================
    def makeConnections(self):
        self.ui.tableView_31.clicked.connect(self.itemClicked)
        self.itemModel31.itemChanged.connect(self.itemChanged)
        # self.quitButton.clicked.connect(self.close)
        self.ui.act_Undo.triggered.connect(self.undoStack.undo)  # Stack 堆栈
        self.ui.act_Redo.triggered.connect(self.undoStack.redo)  # Stack 堆栈
        # self.ui.undoButton.clicked.connect(self.undoStack.undo)  # Stack 堆栈
        # self.ui.redoButton.clicked.connect(self.undoStack.redo)

    # 断开信号链接 用于undo redo
    def disconnectSignal(self):
        self.itemModel31.itemChanged.disconnect(self.itemChanged)

    # 数据模型有变化 则触发 self.itemChanged 用于undo redo
    def connectSignal(self):
        self.itemModel31.itemChanged.connect(self.itemChanged)

    # QTreeView 被单击时 触发 用于undo redo
    def itemClicked(self, index):
        # print("itemClicked ")
        item = self.itemModel31.itemFromIndex(index)
        # print(item)
        self.textBeforeEdit = item.text()

    # 数据模型有变化 则触发 self.itemChanged 用于undo redo
    def itemChanged(self, item):
        try:
            command = CommandItemEdit(self.connectSignal, self.disconnectSignal, self.itemModel1, item, self.textBeforeEdit,
                                      "改 '{0}' to '{1}'".format(self.textBeforeEdit, item.text()))
            self.undoStack.push(command)
        except Exception as e:
            print("发生了undo itemChanged错误 ： ", e)

    # TODO  ==============Echarts数据可视化==========================

    # TODO  ==============自定义功能函数==========================
    # 多线程 广告信息的获取
    def advertise_info(self, info):
        if info:
            index = info[0]
            size = info[2]
            color = info[3]
            str1 = ast.literal_eval(color)
            r = str1[0]
            g = str1[1]
            b = str1[2]
            pe = QPalette()
            # pe.setColor(QPalette.WindowText, Qt.red)  # 设置字体颜色
            pe.setColor(QPalette.WindowText, QColor(r, g, b))  # 设置字体颜色
            self.__LabFile2.setPalette(pe)

            words = info[4]
            font = self.__LabFile2.font()
            font.setPointSize(size)

            self.__LabFile2.setFont(font)
            self.__LabFile2.setText(f'{index}、{words}')
        # self.thread.wait()

    # 二维表转置
    def transpose_2d(self, data):
        transposed = []
        cols = len(data)
        # 判断行数
        rows = 0
        for col in range(cols):
            temprows = len(data[col])
            if temprows > rows:
                rows = temprows
        # rows = len(data[0])
        for row in range(rows):
            list2 = []
            for col in range(cols):
                try:
                    itemstr = data[col][row]
                except:
                    itemstr = ''
                list2.append(itemstr)
            transposed.append(list2)
        return transposed

    # 房号命重复时红色  辅助
    def sameBuildName(self):
        rows = self.itemModel0.rowCount()
        listkeyword = []
        for row in range(rows):
            keystr = self.itemModel0.item(row, 0).text()
            item = QStandardItem(keystr)
            font = item.font()
            if keystr in listkeyword:  # 如果有重复
                # 变红色
                font.setBold(True)
                font.setPointSize(13)
                colorstr = QColor(200, 0, 0)
            else:   # 如果没有重复
                # 黑色
                font.setBold(False)
                font.setPointSize(12)
                colorstr = QColor(0, 0, 0)  # 黑色
            self.itemModel0.setItem(row, 0, QStandardItem(keystr))
            self.itemModel0.item(row, 0).setForeground(colorstr)  # 设置字体颜色
            listkeyword.append(keystr)

    # 打开工程时 工程信息 写入每个控件
    def engineeringInforSetControl(self, tempDict):
        res = tempDict.get('项目名称')
        if res:
            self.ui.lineEdit_1.setText(res)
        else:
            self.ui.lineEdit_1.setText('')

        res = tempDict.get('所属事业部')
        if res:
            self.ui.lineEdit_2.setText(res)
        else:
            self.ui.lineEdit_2.setText('')

        res = tempDict.get('项目所在省份')
        if res:
            self.ui.lineEdit_3.setText(res)
        else:
            self.ui.lineEdit_3.setText('')

        res = tempDict.get('项目所在市')
        if res:
            self.ui.lineEdit_4.setText(res)
        else:
            self.ui.lineEdit_4.setText('')

        res = tempDict.get('地区类型')
        if res:
            self.ui.lineEdit_5.setText(res)
        else:
            self.ui.lineEdit_5.setText('')

        res = tempDict.get('设计院名称')
        if res:
            self.ui.lineEdit_6.setText(res)
        else:
            self.ui.lineEdit_6.setText('')

        res = tempDict.get('人防设计院名称')
        if res:
            self.ui.lineEdit_7.setText(res)
        else:
            self.ui.lineEdit_7.setText('')


    # 房号表的单体信息 发生变化时 写入每个控件
    def singleInforSetControl(self, tempDict):
        res = tempDict.get('业态')
        if res:
            self.ui.singleCBBox_1.setCurrentText(res)
        else:
            self.ui.singleCBBox_1.setCurrentText('')

        res = tempDict.get('抗震等级')
        if res:
            self.ui.singleCBBox_2.setCurrentText(res)
        else:
            self.ui.singleCBBox_2.setCurrentText('')

        res = tempDict.get('设防烈度')
        if res:
            self.ui.singleCBBox_3.setCurrentText(res)
        else:
            self.ui.singleCBBox_3.setCurrentText('')

        res = tempDict.get('基础类型')
        if res:
            self.ui.singleCBBox_4.setCurrentText(res)
        else:
            self.ui.singleCBBox_4.setCurrentText('')

        res = tempDict.get('屋面形式')
        if res:
            self.ui.singleCBBox_5.setCurrentText(res)
        else:
            self.ui.singleCBBox_5.setCurrentText('')

        res = tempDict.get('结构类型')
        if res:
            self.ui.singleCBBox_6.setCurrentText(res)
        else:
            self.ui.singleCBBox_6.setCurrentText('')

        res = tempDict.get('首层层高')
        if res:
            self.ui.spinBox_1.setValue(res)
        else:
            self.ui.spinBox_1.setValue(0)

        res = tempDict.get('标准层高')
        if res:
            self.ui.spinBox_2.setValue(res)
        else:
            self.ui.spinBox_2.setValue(0)

        res = tempDict.get('地上层数')
        if res:
            self.ui.spinBox_3.setValue(res)
        else:
            self.ui.spinBox_3.setValue(0)

        res = tempDict.get('地下层数')
        if res:
            self.ui.spinBox_4.setValue(res)
        else:
            self.ui.spinBox_4.setValue(0)

        res = tempDict.get('PC率')
        if res:
            self.ui.spinBox_5.setValue(res)
        else:
            self.ui.spinBox_5.setValue(0)

        res = tempDict.get('新三板率')
        if res:
            self.ui.spinBox_6.setValue(res)
        else:
            self.ui.spinBox_6.setValue(0)

    # 单体信息表内的控件变化时 写入房号单体信息 （房号表第二个单元格）
    def singleInforSet(self, str1, str2):  # str1 字典key str2 value
        # self.singleInforDict = {}  # 单体信息  写入房号表第二个单元格
        row = self.selectionModel0.currentIndex().row()
        item = self.itemModel0.item(row, 1).text()
        if item:
            self.singleInforDict = ast.literal_eval(item)  # 还原字典
        else:
            self.singleInforDict = {}
        self.singleInforDict[str1] = str2
        # 写入房号表
        item = QStandardItem(str(self.singleInforDict))
        self.itemModel0.setItem(row, 1, item)

    # 房号表 写入  各类表格
    # 房号表被点击时 触发
    def buildNumTableList_to_mod(self, tablelist, model, noCol=None):  # 二维表， 模型，不计标志的列号
        if not model:
            return
        rows = len(tablelist)
        cols = len(tablelist[0])
        model.setRowCount(rows)  # 重置行数
        for row in range(rows):
            for col in range(cols):
                cellstr = tablelist[row][col]
                if col == noCol:  # 在“不计标志” 单元格
                    item = QStandardItem(self.__NoCalTitle)
                    item.setFlags(self.__NoCalFlags)
                    item.setCheckable(True)  # 非锁定
                    if cellstr == "1":
                        item.setCheckState(Qt.Checked)  # 勾选
                    else:
                        item.setCheckState(Qt.Unchecked)  # 勾选
                    model.setItem(row, col, item)  # 赋值
                else:
                    item = QStandardItem(cellstr)
                    model.setItem(row, col, item)

    # 各类表格写入 房号表
    # 各表有数据变化时，增加 删除 插入 粘贴行 时 触发
    def mod_to_buildNumTable(self, model, noCol=None):  # 模型，不计标志的列号
        curowBuildNum = self.selectionModel0.currentIndex().row()
        # print("房号表在 ：", curowBuildNum, " 行")
        curcolBuildNum = 0
        if model == self.itemModel1:
            # print("面积表有数据写入")
            curcolBuildNum = 2
        elif model == self.itemModel2:
            # print("措施表有数据写入")
            curcolBuildNum = 3
        elif model == self.itemModel31:
            # print("分部分项表有数据写入")
            curcolBuildNum = 4
        if curcolBuildNum == 0:
            return
        rows = model.rowCount()
        cols = model.columnCount()
        datalist = []
        for row in range(rows):
            for col in range(cols):
                item = model.item(row, col)
                if col == 0:
                    datalist.append([])
                if col == noCol:
                    if item.checkState() == Qt.Checked:
                        item = "1"
                    else:
                        item = ""
                    datalist[row].append(item)
                else:
                    datalist[row].append(item.text())
        items = str(datalist)
        item = QStandardItem(items)
        self.itemModel0.setItem(curowBuildNum, curcolBuildNum, item)  # 明细写入隐藏列

    # 刷新 建筑面积 统计
    def census_area(self):
        rows = self.itemModel1.rowCount()
        total_area = 0  # 总面积
        upper_area = 0  # 地上面积
        lower_area = 0  # 地下面积
        land_area = 0  # 占地面积，  面积最大的一层

        for row in range(rows):
            resultstr = self.itemModel1.item(row, 5).text()  # 取出每一行的“计算结果”
            if resultstr:
                resultfloat = float(resultstr)
                if resultfloat > land_area:
                    land_area = resultfloat

        for row in range(rows):
            numstr = self.itemModel1.item(row, 6).text()  # 取出每一行的面积
            if numstr:
                numfloat = float(numstr)
                if self.itemModel1.item(row, 1).text() == "地下":
                    lower_area += numfloat
                elif self.itemModel1.item(row, 1).text() == "地上":
                    upper_area += numfloat
                total_area += numfloat
        if upper_area + lower_area != total_area and total_area != 0:  # 如果地下 地上 都空 总面积有，就把所有面积都放在地上
            upper_area = total_area - lower_area
        # 格式化小数位数
        upper_area = round(upper_area, self.decimalPlaces)
        lower_area = round(lower_area, self.decimalPlaces)
        total_area = round(total_area, self.decimalPlaces)
        land_area = round(land_area, self.decimalPlaces)
        # 写入地上 地下 总面积 占地面积
        self.ui.label_37.setText(str(upper_area))
        self.ui.label_38.setText(str(lower_area))
        self.ui.label_33.setText(str(total_area))
        self.ui.label_66.setText(str(land_area))

    # 刷新 措施费统计
    def census_measures(self):
        rows = self.itemModel2.rowCount()
        total_measures = 0  # 总措施费
        whole_measures = 0  # 整体措施费
        single_measures = 0  # 单项措施费
        other_measures = 0  # 其他措施费
        for row in range(rows):
            money = self.itemModel2.item(row, 9).text()  # 取出每一行合价
            keyword = self.itemModel2.item(row, 1).text()  # 取出每一行合价
            if money:
                money = float(money)
                # 判断 类别
                if keyword == "整体措施费":
                    whole_measures += float(money)
                elif keyword == "单项措施费":
                    single_measures += float(money)
                elif keyword == "其他项目清单":
                    other_measures += float(money)
                total_measures += float(money)
        # 写入标签中
        self.ui.label_9.setText(str(round(whole_measures, self.decimalPlaces)))
        self.ui.label_10.setText(str(round(single_measures,self.decimalPlaces)))
        self.ui.label_11.setText(str(round(other_measures, self.decimalPlaces)))
        self.ui.label_3.setText(str(round(total_measures, self.decimalPlaces)))

    # 写入历史操作的功能
    def undoredo_listview_write(self):
        lastRow = self.itemModel_listView_undo.rowCount()
        self.itemModel_listView_undo.insertRow(lastRow)  # 在尾部插入一空行
        index = self.itemModel_listView_undo.index(lastRow, 0)  # 获取最后一行的ModelIndex
        self.itemModel_listView_undo.setData(index, f"{self.beforestrundo[1]} to {self.afterstrundo}")  # 设置显示文字
        self.ui.listView_undo.setCurrentIndex(index)  # 设置当前选中的行

    # 返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制
    # 表名 1 面积表 2 措施表 31 分部分项表 32 明细表 4 分类表
    def viewCurrentModel(self):
        # print("开始判断模型")
        tablename = None
        if self.ui.tableView_1.hasFocus():
            # print("表1 在当前界面")
            selectmodelRows = len(self.selectionModel1.selectedRows())
            if selectmodelRows:
                tablename = 1
                tableobj = self.ui.tableView_1
                model = self.itemModel1
                selectModel = self.selectionModel1
            else:
                print("表1 未获得焦点")
                return
        elif self.ui.tableView_2.hasFocus():
            # print("表2 在当前界面")
            selectmodelRows = len(self.selectionModel2.selectedRows())
            if selectmodelRows:
                tablename = 2
                tableobj = self.ui.tableView_2
                model = self.itemModel2
                selectModel = self.selectionModel2
            else:
                # print("表2 未获得焦点")
                return
        elif self.ui.tableView_31.hasFocus():
            # print("表31 在当前界面")
            selectmodelRows = len(self.selectionModel31.selectedRows())
            if selectmodelRows:
                tablename = 31
                tableobj = self.ui.tableView_31
                model = self.itemModel31
                selectModel = self.selectionModel31
            else:
                # print("表31 未获得焦点")
                return
        elif self.ui.tableView_32.hasFocus():
            # print("表32 在当前界面")
            selectmodelRows = len(self.selectionModel32.selectedRows())
            if selectmodelRows:
                tablename = 32
                tableobj = self.ui.tableView_32
                model = self.itemModel32
                selectModel = self.selectionModel32
            else:
                # print("表32 未获得焦点")
                return
        else:
            # print("所有表 未在可视界面")
            self.ui.statusBar.showMessage("没有表格被选中，无法增、删、插、复制行", 5000)
            return
        rowsIndexList = [i.row() for i in selectModel.selectedRows()]
        rowsIndexList.sort(reverse=True)
        return tablename, tableobj, model, selectModel, rowsIndexList
        # print("判断模型完成")

    # 不计标志变色 辅助 参数： 模型，不计标志的列，表达式的列号,最后两个列号是需要归0
    # 暂时只考虑明细表 其他表暂未考虑
    # def noCalcuColorChange(self, model, noCol, noColstr, resultRealcol=3, floor=5, position=6):
    def noCalcuColorChange(self, model, noCol, noColstr, zeroCol=[3, 5, 6]):
        rows = model.rowCount()
        for row in range(rows):
            boolstr = model.item(row, noCol)  # 不计标志 单元格对象
            strtemp = model.item(row, noColstr).text()  # 要变色的单元格文字
            if (boolstr.checkState() == Qt.Checked):  # 勾选了不计标志
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

    # 用于itemModel32_itemChanged 处理部位、楼层汇总
    def floorPositionSum(self):
        im2rows: int = self.itemModel32.rowCount()  # 声明变量类型的赋值语句
        # 部位position  楼层floor  str名称  RowStart起始行  RowFinish结束行
        positionRowStart, positionRowFinish = None, None
        floorRowStart, floorRowFinish = None, None
        for r in range(im2rows):  # 处理  楼层 部位 汇总
            # ~~~~~~~~~~~处理只有一行的情况~~~~~~~~~~~S
            if im2rows == 1:
                temp = self.itemModel32.item(0, 3).text()
                item = QStandardItem(str(temp))
                self.itemModel32.setItem(0, 5, item)  # 部位 把累加值 写入 起始行的那一格
                temp = self.itemModel32.item(0, 3).text()
                item = QStandardItem(str(temp))
                self.itemModel32.setItem(0, 6, item)  # 楼层 把累加值 写入 起始行的那一格
                continue
            # ~~~~~~~~~~~处理只有一行的情况~~~~~~~~~~~E
            positionStr = self.itemModel32.item(r, 1).text().strip()  # 获取部位字符
            floorStr = self.itemModel32.item(r, 0).text().strip()  # 获取楼层字符
            if positionStr and (positionRowStart == None) and (positionRowFinish == None):  # 如果部位名称  存在 且起始行为空时
                positionRowStart = r  # 第一次找到起始行
            elif positionStr and (positionRowStart != None) and (positionRowFinish == None):  # 找到部位关键字 起始行已存在
                positionRowFinish = r - 1  # 再次出现部位名，是前一次的结束行
            elif floorStr and (positionRowStart != None) and (positionRowFinish == None):  # 找到楼层关键字 起始行已存在
                positionRowFinish = r - 1  # 再次出现部位名，是前一次的结束行
            # elif floorStr and (positionRowStart != None) and (positionRowFinish == None):  # 找到楼层关键字 起始行已存在
            #     positionRowFinish = r - 1  # 再次出现部位名，是前一次的结束行
            if floorStr:  # 找到楼层关键字
                if not positionStr and positionRowStart != None and positionRowFinish != None:
                    pass
                elif not positionStr:
                    positionRowStart, positionRowFinish = None, None
                if (floorRowStart == None) and (floorRowFinish == None):  # 找到部位关键字 起始行已存在
                    floorRowStart = r  # 首次出现楼层名 部位名为空
                elif (floorRowStart != None) and (floorRowFinish == None):  # 找到部位关键字 起始行已存在
                    floorRowFinish = r - 1  # 首次出现楼层名 部位名为空
            #  ~~~~~~~处理循环到最后一行时~~~~~~~~
            if r == im2rows - 1:  # 如果到了最后一行  处理最后一行
                # print(f'到了最后一行,"positionRowStart"={positionRowStart}')
                if positionRowStart != None and positionRowStart == r:  # 处理部位 有起始行
                    positionRowFinish = r
                if positionRowStart != None and positionRowStart < r:
                    if (not positionStr and floorStr):  # 只有楼层
                        positionRowFinish = r - 1  # 这条是给上一个部位汇总 设置的
                        temp = self.itemModel32.item(r, 3).text()
                        item = QStandardItem(str(temp))
                        self.itemModel32.setItem(r, 5, item)  # 把累加值 写入 起始行的那一格'
                    elif (positionStr):  #  如果最后一行有 部位名字 则判断楼层结束行
                        positionRowFinish = r - 1  #  这条是给上一个部位汇总 设置的
                        temp = self.itemModel32.item(r, 3).text()
                        item = QStandardItem(str(temp))
                        self.itemModel32.setItem(r, 5, item)  # 把累加值 写入 起始行的那一格'
                    else:
                        positionRowFinish = r
                if floorRowStart != None and floorRowStart == r:  #  如果最后一行有楼层名字 则判断楼层结束行
                    floorRowFinish = r
                elif floorRowStart != None and floorRowStart < r:
                    if floorStr:
                        floorRowFinish = r - 1  #  这条是给上一个部位汇总 设置的
                        temp = self.itemModel32.item(r, 3).text()
                        item = QStandardItem(str(temp))
                        self.itemModel32.setItem(r, 6, item)  # 把累加值 写入 起始行的那一格'
                    else:
                        floorRowFinish = r
                if floorRowStart != None and (floorRowFinish == None):
                    floorRowFinish = r
            #  ~~~~~~~处理部位累加~~~~~~~~
            if positionRowStart != None and positionRowFinish != None:  # 有始有终时  处理 然后归0 处理部位汇总
                # print('开始汇总计算')
                positionSum = 0
                for positionrow in range(positionRowStart, positionRowFinish + 1):  # 在 开始 结束行 之间循环 含结束行
                    if not self.itemModel32.item(positionrow, 3).text():  # 如果为空 则赋值0
                        temp = 0
                    else:
                        temp = self.itemModel32.item(positionrow, 3).text()
                    positionSum += float(temp)  # 循环取出“计算结果” 累加
                item = QStandardItem(str(round(positionSum,2)))
                self.itemModel32.setItem(positionRowStart, 5, item)  # 把累加值 写入 起始行的那一格
                if r != im2rows - 1:  #
                    positionRowStart, positionRowFinish = positionRowFinish + 1, None  # 新的起始行赋值为老的结束行，新的结束行为空
            #  ~~~~~~~处理楼层累加~~~~~~~~
            if floorRowStart != None and floorRowFinish != None:  # 有始有终时  处理 然后归0 处理部位汇总
                # print('开始汇总计算')
                floorSum = 0
                for floorrow in range(floorRowStart, floorRowFinish + 1):  # 在 开始 结束行 之间循环 含结束行
                    if not self.itemModel32.item(floorrow, 3).text():  # 如果为空 则赋值0
                        temp = 0
                    else:
                        temp = self.itemModel32.item(floorrow, 3).text()
                    floorSum += float(temp)  # 循环取出“计算结果” 累加
                item = QStandardItem(str(round(floorSum,2)))
                self.itemModel32.setItem(floorRowStart, 6, item)  # 把累加值 写入 起始行的那一格
                floorRowStart, floorRowFinish = floorRowFinish + 1, None  # 新的起始行赋值为老的结束行，新的结束行为空
        # self.detailedTotal()  # 明细表中的楼层汇总数量 写入到主表工程量 与 本行工程量小计标签

    # 用于itemModel32_itemChanged  # 注释 中文括号 工程量表 处理，楼层引用暂不处理
    # 处理表达式的字符 返回值 "float", strtempfloat 、 "公式错误", "●" 、  "待处理", strtemp
    def expression_machining(self, str):
        # print("expression_machining", str)
        # ~~~~~~~~~~~~~~~~~~【注释】 中文圆括号 处理~~~~~~~~~~~~~~~~~~~~~
        strtemp = re.sub(r'\u3010[^\u3010]*\u3011', '', str)  # 【注释】 删除
        strtemp = re.sub('（', '(', strtemp)  # 中文圆括号替换英文 删除
        strtemp = re.sub('）', ')', strtemp)  # 中文圆括号替换英文 删除
        # print("# ~~~~~~~~~~~~~~~~~《工程量表》 查找替换~~~~~~~~~~~~~~~~")
        quanDictNum = len(self.quantitiesDict)  # 工程量总表有数据
        quantities = self.quantitiesDict.get("工程量总表")
        if quanDictNum and quantities:
            quantitiesRows = len(self.quantitiesDict.get("工程量总表"))
            # print("工程量表不为空")
            # 工程量表有数据  尝试搜索 可能有多个《工程量表》 所以用while 循环
            while True:  # 直到没有引用的符号
                strkey = re.search('\《[^\《]*?\》', strtemp)  # 先用正则搜索
                # strkey 返回 None 或者 匹配的对象 用group进一步处理
                if not strkey:  # 没有匹配正则 则退出while
                    break  # 退出while
                strkey = strkey.group()[1:-1]  # 脱去两侧的括号
                keystr = "开始找"
                # print('strkey找到，开始找引用工程量表', strkey)
                for rowr in range(quantitiesRows):  # 工程量表有数据 内搜索引用
                    floorrowstr = self.quantitiesDict.get("工程量总表")[rowr][2]  # 获取编号字符
                    if floorrowstr == strkey:
                        strtemp = re.sub('\《[^\《]*?\》', self.quantitiesDict.get("工程量总表")[rowr][6], strtemp,
                                          count=1)  # 把第6列楼层汇总 替换成[工程量表]字符，替换第一处
                        keystr = "找到了"
                        break  # 找到就跳出 工程量表的循环 for
                if keystr == "开始找":  # 有《》 但是里面对应的工程量 没找到
                    return "公式错误", "●"
        # 此时 注释 中文括号 工程量表引用 已经处理完毕 只剩下<楼层>引用
        # 尝试转换公式 处理楼层在 主函数处理
        try:  # 传入的非表达式  无法转换为eval则出错
            strtempstr = eval(strtemp)  # 表达式 转结果
            strtempfloat = round(strtempstr, self.decimalPlaces)  # 保留三位小数
            return "float", strtempfloat
        except:
            return "待处理", strtemp

    def m3ChangeRed(self, model, row=None):
        rows = model.rowCount()
        # 判断 计量单位所在的列号
        if self.tabindex == "建筑面积计算表":
            return
        elif self.tabindex == "措施项目清单":
            col = 4
        elif self.tabindex == "分部分项清单":
            col = 6
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

    # 色标 颜色 根据色标更改整行的 背景色 分部分项表
    # 用于读取存储记录时的 颜色刷新
    def rowsBackground(self, model, col=1, row=None):
        # print("色标 颜色 根据色标更改整行的")
        # print("1")
        rows = model.rowCount()
        # print("2")
        columns = model.columnCount()
        # print("3")
        if row != None:
            # print("单行 色标")
            text = model.item(row, col).text()
            # text2 = self.itemModel31.item(currRow, 0).text()
            if text.startswith('1'):
                colorstr = QColor(222, 28, 49)
            elif text.startswith('2'):
                colorstr = QColor(36, 116, 181)
            elif text.startswith('3'):
                colorstr = QColor(210, 180, 44)
            elif text.startswith('4'):
                colorstr = QColor(178, 207, 135)
            else:
                # colorstr = self.itemModel.item(0, 0).background()
                if row % 2:
                    colorstr = QColor(245, 245, 245)
                else:
                    colorstr = QColor(255, 255, 255)
            # item = QStandardItem("")
            # self.itemModel31.setItem(currRow, 1, item)
            for x in range(1, 2):
                model.item(row, x).setBackground(colorstr)  # 1列变色
            model.item(row, 1).setForeground(colorstr)  # 色标字体 背景 同色
        else:  # 初始化
            # print("多行 色标")
            for r in range(rows):
                text = model.item(r, col).text()
                # text2 = self.itemModel31.item(currRow, 0).text()
                if text.startswith('1'):
                    colorstr = QColor(222, 28, 49)
                elif text.startswith('2'):
                    colorstr = QColor(36, 116, 181)
                elif text.startswith('3'):
                    colorstr = QColor(210, 180, 44)
                elif text.startswith('4'):
                    colorstr = QColor(178, 207, 135)
                else:
                    # colorstr = self.itemModel.item(0, 0).background()
                    if r % 2:
                        colorstr = QColor(245, 245, 245)
                    else:
                        colorstr = QColor(255, 255, 255)
                # item = QStandardItem("")
                # self.itemModel31.setItem(currRow, 1, item)
                for x in range(1, 2):
                    model.item(r, x).setBackground(colorstr)  # 1列变色
                model.item(r, 1).setForeground(colorstr)  # 色标单元格 字体 背景 同色 隐藏字效果

    # 粗体 序号 为ABCD...  一二三...开头  分部分项表
    def boldcol0(self, model, row=None):  # self.boldcol0(self.itemModel31, currRow)
        # print("粗体0")
        col = 0
        rows = model.rowCount()
        columns = model.columnCount()
        if row != None:
            # print("# 单独一行加粗")
            curstr = model.item(row, col).text().strip()
            colorstr = model.item(row, col).background()  # 1001 记录原始背景色
            if curstr == "":
                # print("序号为空值！")
                return
            fristcurstr = curstr[0]
            item = QStandardItem(curstr)
            if fristcurstr in ["A", "B", "C", "D", "E"] or fristcurstr in ["一", "二", "三", "四", "五"]:
                font = item.font()
                font.setBold(True)
                font.setPointSize(self.spinFontSize + 1)  # 比设定字体大1号
            else:
                font = item.font()
                font.setBold(False)
            item.setFont(font)
            model.setItem(row, col, item)
            model.item(row, col).setBackground(colorstr)  # 1001 背景色还原
        else:  # 初始化
            # print("全部行 判断0")
            for r in range(rows):
                curstr = model.item(r, col).text().strip()
                colorstr = model.item(r, col).background()  # 1001 记录原始背景色
                if curstr == "":
                    # print("序号为空值！")
                    continue
                fristcurstr = curstr[0]
                item = QStandardItem(curstr)
                if fristcurstr in ["A", "B", "C", "D", "E"] or fristcurstr in ["一", "二", "三", "四", "五"]:
                    font = item.font()
                    font.setBold(True)
                    font.setPointSize(self.spinFontSize + 1)  # 比设定字体大2号
                else:
                    font = item.font()
                    font.setBold(False)
                item.setFont(font)
                model.setItem(r, col, item)
                model.item(r, col).setBackground(colorstr)  # 1001 背景色还原

    # 汇总清单表合价总和
    def itemModel31_summary(self):
        # print("计算 分部分项清单合计")
        sumnum = 0
        rows = self.itemModel31.rowCount()
        for row in range(rows):
            item1 = self.itemModel31.item(row, 7).text()  # 工程量
            try:
                item1 = float(item1)
            except:
                item1 = 0
            item2 = self.itemModel31.item(row, 9).text()  # 单价
            try:
                item2 = float(item2)
            except:
                item2 = 0
            item = self.itemModel31.item(row, 10).text()
            if item1 and item2:
                sumstr = round(item1 * item2, self.decimalPlaces)
                itemsum = QStandardItem(str(sumstr))
                self.itemModel31.setItem(row, 10, itemsum)
                sumnum += sumstr
            else:
                itemoo = QStandardItem("")
                self.itemModel31.setItem(row, 10, itemoo)  # 项目合价为空
        sumnum = round(sumnum, 2)
        self.ui.label_23.setText(str(sumnum))

    # 辅助 quantiti_TotalPrice(self): 不计标志 清零
    def noCalColquantiti(self, liststr, noCol, zeroCol=[3, 5, 6]):
        rows = len(liststr)
        for row in range(rows):
            strtemp = liststr[row][noCol]
            if strtemp == "1":  # 勾选了不计标志
                # 真实的计算结果real 部位 楼层 归0
                for x in zeroCol:
                    liststr[row][x] = ""
        return liststr

    # 辅助 quantiti_TotalPrice(self):  处理部位、楼层汇总
    def floorPositionSumquantiti(self, liststr):
        # im2rows: int = self.itemModel32.rowCount()  # 声明变量类型的赋值语句
        im2rows = len(liststr)
        # 部位position  楼层floor  str名称  RowStart起始行  RowFinish结束行
        positionRowStart, positionRowFinish = None, None
        floorRowStart, floorRowFinish = None, None
        for r in range(im2rows):  # 处理  楼层 部位 汇总
            # ~~~~~~~~~~~处理只有一行的情况~~~~~~~~~~~S
            if im2rows == 1:
                # temp = self.itemModel32.item(0, 3).text()
                temp = liststr[0][3]
                liststr[0][5] = temp
                liststr[0][5] = temp
                continue
            # ~~~~~~~~~~~处理只有一行的情况~~~~~~~~~~~E
            positionStr = liststr[r][1]
            # positionStr = self.itemModel32.item(r, 1).text().strip()  # 获取部位字符
            floorStr = liststr[r][0]
            # floorStr = self.itemModel32.item(r, 0).text().strip()  # 获取楼层字符
            if positionStr and (positionRowStart == None) and (positionRowFinish == None):  # 如果部位名称  存在 且起始行为空时
                positionRowStart = r  # 第一次找到起始行
            elif positionStr and (positionRowStart != None) and (positionRowFinish == None):  # 找到部位关键字 起始行已存在
                positionRowFinish = r - 1  # 再次出现部位名，是前一次的结束行
            elif floorStr and (positionRowStart != None) and (positionRowFinish == None):  # 找到楼层关键字 起始行已存在
                positionRowFinish = r - 1  # 再次出现部位名，是前一次的结束行
            # elif floorStr and (positionRowStart != None) and (positionRowFinish == None):  # 找到楼层关键字 起始行已存在
            #     positionRowFinish = r - 1  # 再次出现部位名，是前一次的结束行
            if floorStr:  # 找到楼层关键字
                if not positionStr and positionRowStart != None and positionRowFinish != None:
                    pass
                elif not positionStr:
                    positionRowStart, positionRowFinish = None, None
                if (floorRowStart == None) and (floorRowFinish == None):  # 找到部位关键字 起始行已存在
                    floorRowStart = r  # 首次出现楼层名 部位名为空
                elif (floorRowStart != None) and (floorRowFinish == None):  # 找到部位关键字 起始行已存在
                    floorRowFinish = r - 1  # 首次出现楼层名 部位名为空
            #  ~~~~~~~处理循环到最后一行时~~~~~~~~
            if r == im2rows - 1:  # 如果到了最后一行  处理最后一行
                # print(f'到了最后一行,"positionRowStart"={positionRowStart}')
                if positionRowStart != None and positionRowStart == r:  # 处理部位 有起始行
                    positionRowFinish = r
                if positionRowStart != None and positionRowStart < r:
                    if not positionStr and floorStr:  # 只有楼层
                        positionRowFinish = r - 1  # 这条是给上一个部位汇总 设置的
                        temp = liststr[r][3]
                        liststr[r][5] = temp
                    elif positionStr:  #  如果最后一行有 部位名字 则判断楼层结束行
                        positionRowFinish = r - 1  #  这条是给上一个部位汇总 设置的
                        temp = liststr[r][3]
                        liststr[r][5] = temp
                    else:
                        positionRowFinish = r
                if floorRowStart != None and floorRowStart == r:  #  如果最后一行有楼层名字 则判断楼层结束行
                    floorRowFinish = r
                elif floorRowStart != None and floorRowStart < r:
                    if floorStr:
                        floorRowFinish = r - 1  #  这条是给上一个部位汇总 设置的
                        temp = liststr[r][3]
                        liststr[r][6] = temp
                    else:
                        floorRowFinish = r
                if floorRowStart != None and (floorRowFinish == None):
                    floorRowFinish = r
            #  ~~~~~~~处理部位累加~~~~~~~~
            if positionRowStart != None and positionRowFinish != None:  # 有始有终时  处理 然后归0 处理部位汇总
                # print('开始汇总计算')
                positionSum = 0
                for positionrow in range(positionRowStart, positionRowFinish + 1):  # 在 开始 结束行 之间循环 含结束行
                    # if not self.itemModel32.item(positionrow, 3).text():  # 如果为空 则赋值0
                    if not liststr[positionrow][3]:  # 如果为空 则赋值0
                        temp = 0
                    else:
                        temp = liststr[positionrow][3]
                    positionSum += float(temp)  # 循环取出“计算结果” 累加
                item = str(positionSum)
                liststr[positionRowStart][5] = item
                if r != im2rows - 1:  #
                    positionRowStart, positionRowFinish = positionRowFinish + 1, None  # 新的起始行赋值为老的结束行，新的结束行为空
            #  ~~~~~~~处理楼层累加~~~~~~~~
            if floorRowStart != None and floorRowFinish != None:  # 有始有终时  处理 然后归0 处理部位汇总
                # print('开始汇总计算')
                floorSum = 0
                for floorrow in range(floorRowStart, floorRowFinish + 1):  # 在 开始 结束行 之间循环 含结束行
                    # if not self.itemModel32.item(floorrow, 3).text():  # 如果为空 则赋值0
                    if not liststr[floorrow][3]:  # 如果为空 则赋值0
                        temp = 0
                    else:
                        # temp = self.itemModel32.item(floorrow, 3).text()
                        temp = liststr[floorrow][3]
                    floorSum += float(temp)  # 循环取出“计算结果” 累加
                item = str(floorSum)
                # item = QStandardItem(str(floorSum))
                # self.itemModel32.setItem(floorRowStart, 6, item)  # 把累加值 写入 起始行的那一格
                liststr[floorRowStart][6] = item
                floorRowStart, floorRowFinish = floorRowFinish + 1, None  # 新的起始行赋值为老的结束行，新的结束行为空
        return liststr
        # self.detailedTotal()  # 明细表中的楼层汇总数量 写入到主表工程量 与 本行工程量小计标签

    # 辅助 刷新清单表每一行的 合价，用于广联达清单量表 更新后刷新对应的清单工程量
    def quantiti_TotalPrice(self):
        mainRows = self.itemModel31.rowCount()
        mainCols = self.itemModel31.columnCount()
        # 分部分项的每一行循环取出
        for mainRow in range(mainRows):
            detailedstr = self.itemModel31.item(mainRow, 8).text()
            if not detailedstr:  # 如果明细表为空 则跳至下一行
                continue

            strkey = re.search('\《[^\《]*?\》', detailedstr)  # 先用正则搜索
            # strkey 返回 None 或者 匹配的对象 用group进一步处理
            if not strkey:  # 没有匹配正则
                continue

            # 字符转为列表 下面主要处理 这个 detailedstrlist 明细表二维表
            detailedstrlist = ast.literal_eval(detailedstr)
            rows = len(detailedstrlist)
            columns = len(detailedstrlist[0])
            # ~~~~~~~~~~~~~~~“计算结果real 计算结果 面积 计算错误 ●”~~~~先清0~~~~~~~~~~~~~~~~~~~~~~~~
            # print("“计算错误 ●”~~~~先清0")
            for row in range(rows):
                detailedstrlist[row][3] = ""
                detailedstrlist[row][4] = ""
                detailedstrlist[row][5] = ""
                detailedstrlist[row][6] = ""
                detailedstrlist[row][8] = ""
            # ~~~~~~~~~~~~~~~“表达式处理”~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            # print("表达式处理")
            for row in range(rows):
                # 取出表达式的原始字符
                str_temp = detailedstrlist[row][2]  # 取出表达式
                if str_temp == "":  # 表达式为空时   '计算结果real'3, '计算结果'4, 都为空
                    detailedstrlist[row][3] = ""
                    detailedstrlist[row][4] = ""
                else:   # 表达式的有内容进一步处理
                    # print("# 注释 中文括号 工程量表 处理，楼层引用暂不处理")
                    # str_temp = ast.literal_eval(str_temp)  # 字符转list
                    result = self.expression_machining(str_temp)
                    if result[0] == "float":
                        # "如果可以转换为 数值"
                        item = str(result[1])  # 转成字符
                        detailedstrlist[row][3] = item
                        detailedstrlist[row][4] = item
                    elif result[0] == "公式错误":
                        detailedstrlist[row][3] = ""
                        detailedstrlist[row][4] = ""
                        detailedstrlist[row][8] = "●"
                    elif result[0] == "待处理":  # 把待处理的字符放在 temp 格
                        str_temp = result[1]
                        detailedstrlist[row][4] = str_temp
                        detailedstrlist[row][3] = ""
            # 处理可能存在的楼层引用前，先 楼层汇总一次
            # 辅助 quantiti_TotalPrice(self): 真实计算结果归0 部位 楼层归0
            detailedstrlist = self.noCalColquantiti(detailedstrlist, 7)
            # 计算部位 楼层 小计
            detailedstrlist = self.floorPositionSumquantiti(detailedstrlist)
            # 处理 <楼层> 引用
            for row in range(rows):
                # 取出表达式的原始字符  计算结果real
                # str_temp3 = self.itemModel32.item(row, 4).text().strip()
                str_temp3 = detailedstrlist[row][4]
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
                        # floorrowstr = self.itemModel32.item(rowr, 0).text().strip()  # 获取楼层字符
                        floorrowstr = detailedstrlist[rowr][0]
                        if floorrowstr == strkey:
                            # 把第6列楼层汇总 替换成[楼层]字符，替换第一处
                            # str_temp3 = re.sub('\<[^\<]*?\>', self.itemModel32.item(rowr, 6).text(), str_temp3, count=1)
                            str_temp3 = re.sub('\<[^\<]*?\>', detailedstrlist[rowr][6], str_temp3, count=1)
                            keystr = "找到了"
                            break  # 找到就跳出 楼层字符的循环 for
                    if keystr == "开始找":  # 说明没找到 有错误
                        detailedstrlist[row][3] = ""
                        detailedstrlist[row][4] = ""
                        detailedstrlist[row][8] = "●"
                        # item = QStandardItem("")
                        # self.itemModel32.setItem(row, 3, item)  # 赋值 给 计算结果temp列
                        # item = QStandardItem("")
                        # self.itemModel32.setItem(row, 4, item)
                        # item = QStandardItem('●')
                        # self.itemModel32.setItem(row, 8, item)
                        break  # 退出while
                # 试着处理计算式 可能有错误的引用
                try:  # 传入的非表达式  无法转换为eval则出错
                    strtempstr = eval(str_temp3)  # 表达式 转结果
                    strtempfloat = round(strtempstr, self.decimalPlaces)  # 保留三位小数
                    strtempfin = str(strtempfloat)  # 转成字符
                    detailedstrlist[row][3] = strtempfin
                    detailedstrlist[row][4] = strtempfin
                except:  # 去除【】 [] 以后 还不能表达式转结果  8列 错误列标记
                    detailedstrlist[row][3] = ""
                    detailedstrlist[row][4] = ""
                    detailedstrlist[row][8] = "●"
            # <楼层> 引用处理后 再 楼层汇总一次
            # print("楼层汇总第二次")
            detailedstrlist = self.noCalColquantiti(detailedstrlist, 7)
            detailedstrlist = self.floorPositionSumquantiti(detailedstrlist)
            detailedstrlist = self.noCalColquantiti(detailedstrlist, 7)
            # ~~~~~~~~~~~~~~~~~每一行数据写入总表~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            item = QStandardItem(str(detailedstrlist))
            self.itemModel31.setItem(mainRow, 8, item)
            sumlist = 0  # 本条清单子目的工程量小计
            for x in range(rows):
                strtemp = detailedstrlist[x][6]
                if strtemp:
                    sumlist += float(strtemp)
            if sumlist:
                item = QStandardItem(str(sumlist))
                self.itemModel31.setItem(mainRow, 7, item)
            else:
                item = QStandardItem("")
                self.itemModel31.setItem(mainRow, 7, item)

    # TODO  ==============event处理函数 事件==========================
    # 窗体关闭时询问
    def closeEvent(self, event):
        quanBtn = self.ui.actGlod_quan.isEnabled()
        if quanBtn:
            dlgTitle = "警告！"
            strInfo = "确定要退出吗？"
            defaultBtn = QMessageBox.No  # 缺省按钮
            result = QMessageBox.question(self, dlgTitle, strInfo,
                                          QMessageBox.Yes | QMessageBox.No, defaultBtn)
            if result == QMessageBox.Yes:
                event.accept()  # 窗口可关闭
            else:
                event.ignore()  # 窗口不能被关闭
        else:
            dlgTitle = "警告！"
            strInfo = "请退出，工程量表窗口！"
            defaultBtn = QMessageBox.Yes  # 缺省按钮
            result = QMessageBox.question(self, dlgTitle, strInfo,
                                          QMessageBox.Yes)
            event.ignore()  # 窗口不能被关闭

    # TODO  ==========由connectSlotsByName()自动连接的槽函数============
    # 追加模板子目  木模 砖模 铝模 二次结构模板等
    @pyqtSlot()
    def on_btn_Template_Suborder_clicked(self):
        print('追加模板子目')
        mainrow = self.itemModel31.rowCount()
        # print(mainrow)
        # 添加子目
        listTemp = [
            ["" for _ in range(20)],
            ["1", "", "地下", "", "模板", "模板", "m2", "0", "", "", "", "分析系统之用", "", "", "", "", "", "", "", ""],
            ["2", "", "地上", "", "模板", "模板", "m2", "0", "", "", "", "分析系统之用", "", "", "", "", "", "", "", ""],
            ["3", "", "地下", "", "模板", "二次结构模板", "m2", "0", "", "", "", "分析系统之用", "", "", "", "", "", "", "", ""],
            ["4", "", "地上", "", "模板", "二次结构模板", "m2", "0", "", "", "", "分析系统之用", "", "", "", "", "", "", "", ""],
            ["5", "", "地下", "", "模板", "铝模", "m2", "0", "", "", "", "分析系统之用", "", "", "", "", "", "", "", ""],
            ["6", "", "地上", "", "模板", "铝模", "m2", "0", "", "", "", "分析系统之用", "", "", "", "", "", "", "", ""]]
        for listx in listTemp:
            blanklist = []
            for x in listx:
                item = QStandardItem(x)
                blanklist.append(item)
            self.itemModel31.insertRow(mainrow, blanklist)
            mainrow += 1
        # 滚动条 移动到最后
        self.ui.tableView_31.verticalScrollBar().setSliderPosition(mainrow)
        self.ui.btn_Template_Suborder.setEnabled(False)  # 追加模板子目 控件不可用，防止重复按下

    # 智能识别已导入的面积表 导入面积导入
    @pyqtSlot()
    def on_btn_area_auto_clicked(self):
        print('智能识别已导入的面积表')
        buildrow = self.selectionModel0.currentIndex().row()  # 当前楼号所在行
        datas = self.itemModel0.item(buildrow, 11).text()  # 取出 绘图工程量 提取后的工程量
        if not datas:
            QMessageBox.information(self, "提示", "未找到提取到的“建筑面积”工程量！\n请提取后点击“刷新主程序”再识别！", QMessageBox.Cancel)
            return
        # print(type(datas))
        datas = ast.literal_eval(datas)
        # print(type(datas))
        rows = len(datas)
        area2list = []
        num = 1
        for row in range(rows):
            reportForm = datas[row][3]  # 报表名称
            if reportForm == "建筑面积":
                describe = datas[row][4][:2]  # 特征描述  地上_建筑面积  取前2个字符 地上 or 地下
                filename = datas[row][1]  # 文件来源
                detaileddatasstr = datas[row][7]  # 明细表的数据
                detaileddatas = ast.literal_eval(detaileddatasstr)
                for drow in range(len(detaileddatas)):  # 循环明细表的行数
                    floorname = detaileddatas[drow][0]  # 楼层名
                    if not floorname:
                        continue  # 跳过楼层空行
                    floorsum = detaileddatas[drow][4]  # 楼层总量
                    if not floorsum:
                        continue  # 跳过楼层空行
                    # 拼接成一条 面积表的数据
                    area1list = [num, describe, floorname, "", floorsum + '【' + filename + '】',
                                 floorsum, floorsum, '', '', '', '']
                    area2list.append(area1list)
                    num += 1
        if len(area2list) < 1:
            QMessageBox.information(self, "提示", "未找到提取到的“建筑面积”工程量！\n请提取后点击“刷新主程序”再识别！", QMessageBox.Cancel)
            return
        # print(area2list)
        # 写入面积计算表
        # 断开链接
        self.itemModel1.itemChanged.disconnect(self.itemModel1_itemChanged)
        rows = len(area2list)
        cols = len(area2list[0])
        mainrow = self.itemModel1.rowCount()  # 面积表的现有总行数

        for row in range(rows):
            templist = []
            for col in range(cols):
                item = str(area2list[row][col])
                item = QStandardItem(item)
                if col == 7:
                    item.setFlags(self.__NoCalFlags)
                    item.setCheckable(True)  # 非锁定
                    item.setCheckState(Qt.Unchecked)  # 非勾选
                templist.append(item)
            self.itemModel1.insertRow(mainrow, templist)
            mainrow += 1
        # 刷新 建筑面积 统计
        self.census_area()
        # 设置控件不可用 防止重复按下
        self.ui.btn_area_auto.setEnabled(False)
        # 恢复链接
        self.itemModel1.itemChanged.connect(self.itemModel1_itemChanged)
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel1)

    # 添加房号
    @pyqtSlot()
    def on_pushButton_addBuild_clicked(self):
        print("添加房号")
        rows = self.itemModel0.rowCount()
        cols = self.itemModel0.columnCount()
        itemlist = []  # QStandardItem 对象列表
        for col in range(cols):
            if col == 0:
                stritem = f"{rows+1}" + "#"
                item = QStandardItem(stritem)
            else:
                item = QStandardItem("")
            itemlist.append(item)  # 一行空数据
        self.itemModel0.appendRow(itemlist)  # 在最下面行的下面插入一行

    # 复制房号
    @pyqtSlot()
    def on_pushButton_copyBuild_clicked(self):
        print('复制房号')
        self.itemModel0.itemChanged.disconnect(self.itemModel0_itemChanged)
        currrow = self.selectionModel0.currentIndex().row()
        if currrow < 0:
            return
        rows = self.itemModel0.rowCount()
        cols = self.itemModel0.columnCount()
        listdata = []
        for col in range(cols):
            item = self.itemModel0.item(currrow, col).text()
            if col == 0:
                item = item + "-1"
            item = QStandardItem(item)
            # if item:
            #     item = item.text()
            # else:
            #     item = ""
            listdata.append(item)
        print(listdata)
        self.itemModel0.insertRow(rows, listdata)
        self.ui.statusBar.showMessage("‘复制房号’ 成功！", 5000)
        self.sameBuildName()      # 房号命重复时红色  辅助
        self.itemModel0.itemChanged.connect(self.itemModel0_itemChanged)

    # 删除房号
    @pyqtSlot()
    def on_pushButton_delBuild_clicked(self):
        print("删除房号")
        # self.itemModel0.itemChanged.disconnect(self.itemModel0_itemChanged)
        rows = self.itemModel0.rowCount()
        cols = self.itemModel0.columnCount()
        currrow = self.selectionModel0.currentIndex().row()
        if rows <= 1:
            self.ui.statusBar.showMessage("全部删光就没房号了，留一个吧！", 10000)
            return
        elif rows == len(self.selectionModel0.selectedRows()):
            self.ui.statusBar.showMessage("不许全删除房号，留一个吧！", 10000)
            return
        # 删除前 警告确认
        res = QMessageBox.warning(self, "警告", "请确认是否需要删除选中的行！", QMessageBox.Yes | QMessageBox.No)
        if QMessageBox.Yes == res:
            print("点了确认！")
            self.itemModel0.removeRow(currrow)
        elif QMessageBox.No == res:
            print("点了否！")
            return
        self.sameBuildName()      # 房号命重复时红色  辅助
        # self.itemModel0.itemChanged.connect(self.itemModel0_itemChanged)

    # 历史操作 清空
    @pyqtSlot()
    def on_pushButton_clearListview_clicked(self):
        print("历史操作 清空")
        self.__init_listView_undo()

    # 复选框  主表中 人材机管利规税 列 隐藏显示
    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_checkBoxFoldPrice_clicked(self, checked):
        if checked:  # 如果是 “非选中” 状态 则打开明细
            print('非选中')
            morenColunmWidth = 60
            self.ui.tableView_31.setColumnWidth(12, morenColunmWidth)  #
            self.ui.tableView_31.setColumnWidth(13, morenColunmWidth)  #
            self.ui.tableView_31.setColumnWidth(14, morenColunmWidth)  #
            self.ui.tableView_31.setColumnWidth(15, morenColunmWidth)  #
            self.ui.tableView_31.setColumnWidth(16, morenColunmWidth)  #
            self.ui.tableView_31.setColumnWidth(17, morenColunmWidth)  #
            self.ui.tableView_31.setColumnWidth(18, morenColunmWidth)  #
            self.ui.tableView_31.setColumnWidth(19, morenColunmWidth)  #
        else:  # "选中" 状态  则 需要隐藏 明细列 人材机管利规税  不能动态切换隐藏 只能设置列宽度
            print('选中')
            morenColunmzero = 0
            self.ui.tableView_31.setColumnWidth(12, morenColunmzero)  #
            self.ui.tableView_31.setColumnWidth(13, morenColunmzero)  #
            self.ui.tableView_31.setColumnWidth(14, morenColunmzero)  #
            self.ui.tableView_31.setColumnWidth(15, morenColunmzero)  #
            self.ui.tableView_31.setColumnWidth(16, morenColunmzero)  #
            self.ui.tableView_31.setColumnWidth(17, morenColunmzero)  #
            self.ui.tableView_31.setColumnWidth(18, morenColunmzero)  #
            self.ui.tableView_31.setColumnWidth(19, morenColunmzero)  #
            # self.ui.num_tableView1.horizontalHeader().hideSection(17)  # 隐藏列  明细表

    # 字体大小 调节按钮
    def spinFontSize_valueChanged(self):
        print("字体大小 调节按钮 更改了！")
        # 断开连接
        self.disconnectAll()
        # 只调节当前的表
        self.spinFontSize = self.__spinFontSize.value()
        font = QFont()
        # font.setFamily("微软雅黑")
        font.setPointSize(self.spinFontSize)
        if self.tabindex == "分部分项清单":
            self.ui.tableView_31.setFont(font)
            self.boldcol0(self.itemModel31)  # 分部工程名称 粗体 放大
            # self.ui.tableView_31.horizontalHeader().setFont(font)
            self.ui.tableView_32.setFont(font)
            # self.ui.tableView_32.horizontalHeader().setFont(font)
        elif self.tabindex == "措施项目清单":
            self.ui.tableView_2.setFont(font)
            # self.ui.tableView_2.horizontalHeader().setFont(font)
        elif self.tabindex == "建筑面积计算表":
            self.ui.tableView_1.setFont(font)
            # self.ui.tableView_1.horizontalHeader().setFont(font)
        # 恢复连接
        self.connectAll()

    # 小数位数 调节按钮
    def decimalPlaces_valueChanged(self):
        print("小数位数 调节按钮 更改了！")
        # 调节 所有的表
        self.decimalPlaces = self.__decimalPlaces.value()

    # 房号表 数据有变化
    def itemModel0_itemChanged(self, index):
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel0.rowCount()
        columns = self.itemModel0.columnCount()
        self.itemModel0.itemChanged.disconnect(self.itemModel0_itemChanged)
        if currCol == 0:
            textstr = self.itemModel0.item(currRow, 0).text().strip()
            if textstr:
                self.itemModel0.setItem(currRow,0, QStandardItem(textstr))
            elif self.currrow_buildNum_str != None:
                item = self.currrow_buildNum_str
                item = QStandardItem(item)
                self.itemModel0.setItem(currRow, 0, item)
                res = QMessageBox.warning(self, "警告", "房号名称不可为空字符！", QMessageBox.Yes | QMessageBox.No)
                if QMessageBox.Yes == res:
                    print("点了确认！")
                elif QMessageBox.No == res:
                    print("点了否！")
        self.sameBuildName()  # 房号名 同名时标记
        self.itemModel0.itemChanged.connect(self.itemModel0_itemChanged)

    # 面积表 数据有变化 处理
    def itemModel1_itemChanged(self, index):
        print("面积表数据更改了！")
        # "导入" 面积 控件可用
        self.ui.btn_area_auto.setEnabled(True)
        self.itemModel1.itemChanged.disconnect(self.itemModel1_itemChanged)  # 主表有变化时触发
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel1.rowCount()
        columns = self.itemModel1.columnCount()
        # 层数 如果输入的不是数值类型则清除 并提示, 如果是数字则显示整数
        if currCol == 3:
            itemtext = self.itemModel1.item(currRow, currCol).text().strip()
            try:
                item = int(float(itemtext))
                item = QStandardItem(str(item))
                self.itemModel1.setItem(currRow, currCol, item)
            except:
                item = QStandardItem("")
                self.itemModel1.setItem(currRow, currCol, item)
                self.ui.statusBar.showMessage("《层数》必须为数值型！", 5000)
            finally:
                pass
        # 写入历史操作
        if currCol not in [5, 6, 8]:
            if self.beforestrundo[0] == 1:
                self.afterstrundo = self.itemModel1.item(currRow, currCol).text()  # 用于历史操作的原始字符
                print(self.afterstrundo)
                self.undoredo_listview_write()
                self.beforestrundo = (1, self.itemModel1.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # 面积表 数据有变化 处理所有行行数据 辅助
        self.itemModel1_itemChanged_All()
        # 刷新 建筑面积 统计
        self.census_area()
        # 恢复链接
        self.itemModel1.itemChanged.connect(self.itemModel1_itemChanged)
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel1, 7)

    #  辅助 面积表 数据有变化 处理所有行行数据
    def itemModel1_itemChanged_All(self):
        rows = self.itemModel1.rowCount()
        # ~~~~~~~~~~~~~~~“计算结果 面积 计算错误 ● 临时列~~~~先清0~~~~~~~~~~~~~~~~~~~~~~~~
        # print("“计算错误 ●”~~~~先清0")
        for row in range(rows):
            # item = QStandardItem("").setTextAlignment(Qt.AlignCenter)  #右对齐" setTextAlignment(Qt::AlignCenter)
            item = QStandardItem("")
            self.itemModel1.setItem(row, 5, item)
            item = QStandardItem("")
            self.itemModel1.setItem(row, 6, item)
            item = QStandardItem("")
            self.itemModel1.setItem(row, 8, item)
            item = QStandardItem("")
            self.itemModel1.setItem(row, 10, item)
        # ~~~~~~~~~~~~~~~“表达式处理”~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # print("表达式处理")
        for row in range(rows):
            # 取出表达式的原始字符
            str_temp = self.itemModel1.item(row, 4).text().strip()
            if str_temp == "":  # 表达式为空时  '计算结果'4, 都为空
                item = QStandardItem("")
                self.itemModel1.setItem(row, 5, item)
            else:   # 表达式的有内容进一步处理
                # print("# 注释 中文括号 工程量表 处理，楼层引用暂不处理")
                result = self.expression_machining(str_temp)
                if result[0] == "float":
                    # "如果可以转换为 数值"
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel1.setItem(row, 5, item)  # 能转化为结果 则直接赋值 给 计算结果列
                    intstr = self.itemModel1.item(row, 3).text()
                    if intstr == "":
                        intstr = 1
                    item = str(result[1] * int(intstr))  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel1.setItem(row, 6, item)  # 面积
                elif result[0] == "公式错误":
                    item = QStandardItem("")  # 需要给一个str格式
                    self.itemModel1.setItem(row, 5, item)  # 清空
                    item = QStandardItem("●")
                    self.itemModel1.setItem(row, 8, item)  # 清空
                elif result[0] == "待处理":  # 把待处理的字符放在 temp 格
                    str_temp = result[1]
                    item = QStandardItem(str_temp)  # 需要给一个str格式
                    self.itemModel1.setItem(row, 10, item)  # 待处理 10列 隐藏
                    item = QStandardItem("")
                    self.itemModel1.setItem(row, 5, item)  # 清空
        # 处理不计标志 归0
        self.noCalcuColorChange(self.itemModel1, 7, 4, [6, 10])
        # 处理 <楼层> 引用
        for row in range(rows):
            # 取出表达式的原始字符  计算结果real
            str_temp3 = self.itemModel1.item(row, 10).text().strip()
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
                    floorrowstr = self.itemModel1.item(rowr, 2).text().strip()  # 获取楼层字符
                    if floorrowstr == strkey:
                        str_temp3 = re.sub('\<[^\<]*?\>', self.itemModel1.item(rowr, 6).text(), str_temp3,
                                          count=1)  # 把第6列楼层汇总 替换成[楼层]字符，替换第一处
                        keystr = "找到了"
                        break  # 找到就跳出 楼层字符的循环 for
                if keystr == "开始找":  # 说明没找到 有错误
                    item = QStandardItem("")
                    self.itemModel1.setItem(row, 5, item)  # 赋值 给 计算结果temp列
                    item = QStandardItem("")
                    self.itemModel1.setItem(row, 6, item)
                    item = QStandardItem('●')
                    self.itemModel1.setItem(row, 8, item)
                    break  # 退出while
            # 试着处理计算式 可能有错误的引用
            try:  # 传入的非表达式  无法转换为eval则出错
                strtempstr = eval(str_temp3)  # 表达式 转结果
                strtempfloat = round(strtempstr, self.decimalPlaces)  # 保留三位小数
                strtempfin = str(strtempfloat)  # 转成字符
                item = QStandardItem(strtempfin)  # 需要给一个str格式
                self.itemModel1.setItem(row, 5, item)  # 能转化为结果 则直接赋值 给 计算结果列
                intstr = self.itemModel1.item(row, 3).text()
                if intstr == "":
                    intstr = 1
                item = str(strtempfloat * int(intstr))  # 转成字符
                item = QStandardItem(item)
                self.itemModel1.setItem(row, 6, item)
            except:  # 去除【】 [] 以后 还不能表达式转结果  8列 错误列标记
                item = QStandardItem("")
                self.itemModel1.setItem(row, 5, item)  # 赋值 给 计算结果temp列
                item = QStandardItem("")
                self.itemModel1.setItem(row, 6, item)
                item = QStandardItem('●')
                self.itemModel1.setItem(row, 8, item)

    # 措施表 数据有变化 处理
    def itemModel2_itemChanged(self, index):
        print("措施表数据更改了！")
        self.itemModel2.itemChanged.disconnect(self.itemModel2_itemChanged)  # 主表有变化时触发
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel2.rowCount()
        columns = self.itemModel2.columnCount()
        # m3红色 m2青色
        if currCol == 4:
            self.m3ChangeRed(self.itemModel2, currRow)
        # 除税单价 含税单价 如果输入的不是数值类型则清除 并提示, 如果是数字则显示整数
        if currCol == 7 or currCol == 8:
            itemtext = self.itemModel2.item(currRow, currCol).text().strip()
            try:
                item = round(float(itemtext), self.decimalPlaces)  # 保留三位小数
                item = QStandardItem(str(item))
                self.itemModel2.setItem(currRow, currCol, item)
            except:
                item = QStandardItem("")
                self.itemModel2.setItem(currRow, currCol, item)
                self.ui.statusBar.showMessage("除税单价 含税单价 必须为数值型！", 8000)
            finally:
                pass
        # 写入历史操作
        if currCol not in [6, 9]:
            if self.beforestrundo[0] == 2:
                self.afterstrundo = self.itemModel2.item(currRow, currCol).text()  # 用于历史操作的原始字符
                print(self.afterstrundo)
                self.undoredo_listview_write()
                self.beforestrundo = (2, self.itemModel2.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # 辅助 措施表 数据有变化 处理所有行行数据
        self.itemModel2_itemChanged_All()
        # 刷新一下 措施费统计
        self.census_measures()
        # 恢复链接
        self.itemModel2.itemChanged.connect(self.itemModel2_itemChanged)
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel2)

    # 辅助 措施表 数据有变化 处理所有行行数据
    def itemModel2_itemChanged_All(self):
        rows = self.itemModel2.rowCount()
        # ~~~~~~~~~~~~~~~“合价~~~~先清0~~~~~~~~~~~~~~~~~~~~~~~~
        for row in range(rows):
            item = QStandardItem("")
            self.itemModel2.setItem(row, 9, item)
        # ~~~~~~~~~~~~~~~“表达式处理”~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # print("表达式处理")
        for row in range(rows):
            # 取出表达式的原始字符
            str_temp = self.itemModel2.item(row, 5).text().strip()
            if str_temp == "":  # 表达式为空时  合价为空
                self.itemModel2.setItem(row, 6, QStandardItem(""))
                self.itemModel2.setItem(row, 9, QStandardItem(""))
            else:  # 表达式的有内容进一步处理
                print("# 注释 中文括号 工程量表 处理，")
                result = self.expression_machining(str_temp)
                if result[0] == "float":
                    # "如果可以转换为 数值"
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel2.setItem(row, 6, item)
                    price = self.itemModel2.item(row, 8).text().strip()
                    if price:
                        item = float(price) * result[1]
                        item = round(item, 2)
                        item = str(item)  # 转成字符
                        item = QStandardItem(item)  # 需要给一个str格式
                        self.itemModel2.setItem(row, 9, item)  # 合价
                elif result[0] == "公式错误":
                    item = QStandardItem("公式错误")  # 需要给一个str格式
                    self.itemModel2.setItem(row, 6, item)  # 清空
                    item = QStandardItem("")
                    self.itemModel2.setItem(row, 9, item)  # 清空
                elif result[0] == "待处理":
                    item = QStandardItem("公式错误")  # 需要给一个str格式
                    self.itemModel2.setItem(row, 6, item)  # 清空
                    item = QStandardItem("")
                    self.itemModel2.setItem(row, 9, item)  # 清空

    # 分部分项表 数据有变化 处理
    def itemModel31_itemChanged(self, index):
        # self.disconnectAll()
        self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)  # 主表有变化时触发
        print("分部分项表数据更改了！")
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel31.rowCount()
        columns = self.itemModel31.columnCount()
        # print(self.beforestrundo)
        # 写入历史操作
        if currCol in [0, 1, 2, 3, 4, 5, 6, 11]:
            if self.beforestrundo[0] == 31:
                self.afterstrundo = self.itemModel31.item(currRow, currCol).text()  # 用于历史操作的原始字符
                print(self.afterstrundo)
                self.undoredo_listview_write()  # 写入历史操作
                self.beforestrundo = (31, self.itemModel31.item(currRow, currCol).text())  # 用于历史操作的原始字符

        # 项目名称更改了
        if currCol == 4:
            test = self.itemModel31.item(currRow, 4).text()
            self.ui.label_35.setText(f'{currRow + 1} ')
            self.ui.label_34.setText(f"项名称：{test}")

        # 粗体 序号 为ABCD...  一二三...开头
        if currCol == 0:
            print("序号首字 ABCDE一二三四五 显示粗体+2")
            self.boldcol0(self.itemModel31, currRow)

        # 色标 显示
        if currCol == 1:
            print("行变色 色标")
            self.rowsBackground(self.itemModel31, currCol, currRow)

        # m3红色 m2青色
        if currCol == 6:
            print("计量单位有数据变化")
            self.m3ChangeRed(self.itemModel31, currRow)

        # 工程量 单价如果输入的不是数值类型则清除 并提示, 如果是数字则按控件上的小数位显示
        if currCol == 7 or currCol == 9 or currCol == 10:
            itemtext = self.itemModel31.item(currRow, currCol).text().strip()
            try:
                fltemp = float(itemtext)
                item = round(fltemp, self.decimalPlaces)
                item = QStandardItem(str(item))
                self.itemModel31.setItem(currRow, currCol, item)
            except:
                itemo = QStandardItem("")
                self.itemModel31.setItem(currRow, currCol, itemo)  # 工程量 单价 为空
                itemo = QStandardItem("")
                self.itemModel31.setItem(currRow, 10, itemo)  # 项目合价为空
                self.ui.statusBar.showMessage("《工程量》 《单价》必须为数值型！", 5000)
            finally:
                pass
        # 辅助 分部分项表 每一行 量 价 合
        self.itemModel31_summary()
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel31)
        # self.connectAll()
        self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged) # 主表有变化时触发

    # 明细表 数据有变化 处理
    def itemModel32_itemChanged(self, index):
        self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)  # 主表有变化时触发
        self.itemModel32.itemChanged.disconnect(self.itemModel32_itemChanged)
        # self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)  # 主表有变化时触发
        currRow = index.row()
        currCol = index.column()
        rows = self.itemModel32.rowCount()
        columns = self.itemModel32.columnCount()
        # print(self.beforestrundo)
        # 写入历史操作
        if currCol in [0, 1, 2, 9]:
            if self.beforestrundo[0] == 32:
                self.afterstrundo = self.itemModel32.item(currRow, currCol).text()  # 用于历史操作的原始字符
                print(self.afterstrundo)
                self.undoredo_listview_write()
            self.beforestrundo = (32, self.itemModel32.item(currRow, currCol).text())  # 用于历史操作的原始字符
        # 辅助 明细表 数据有变化 处理所有行行数据
        self.itemModel32_itemChanged_All()
        print("每一行数据写入总表 完成")
        # 刷新一下 总计
        self.itemModel31_summary()
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel31)
        # 恢复响应
        self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged)
        self.itemModel32.itemChanged.connect(self.itemModel32_itemChanged)
        # self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged) # 主表有变化时触发

    # 辅助 明细表 数据有变化 处理所有行行数据
    def itemModel32_itemChanged_All(self):
        rows = self.itemModel32.rowCount()
        columns = self.itemModel32.columnCount()
        table31_row = self.selectionModel31.currentIndex().row()
        # ~~~~~~~~~~~~~~~“计算结果real 计算结果 面积 计算错误 ●”~~~~先清0~~~~~~~~~~~~~~~~~~~~~~~~
        print("“计算错误 ●”~~~~先清0")
        for row in range(rows):
            item = QStandardItem("")
            self.itemModel32.setItem(row, 3, item)
            item = QStandardItem("")
            self.itemModel32.setItem(row, 4, item)
            item = QStandardItem("")
            self.itemModel32.setItem(row, 5, item)
            item = QStandardItem("")
            self.itemModel32.setItem(row, 6, item)
            item = QStandardItem("")
            self.itemModel32.setItem(row, 8, item)
        # ~~~~~~~~~~~~~~~“表达式处理”~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        print("表达式处理")
        for row in range(rows):
            # 取出表达式的原始字符
            str_temp = self.itemModel32.item(row, 2).text().strip()
            if str_temp == "":  # 表达式为空时   '计算结果real'3, '计算结果'4, 都为空
                item = QStandardItem("")
                self.itemModel32.setItem(row, 3, item)
                item = QStandardItem("")
                self.itemModel32.setItem(row, 4, item)
            else:   # 表达式的有内容进一步处理
                print("# 注释 中文括号 工程量表 处理，楼层引用暂不处理")
                result = self.expression_machining(str_temp)
                if result[0] == "float":
                    # "如果可以转换为 数值"
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel32.setItem(row, 4, item)  # 能转化为结果 则直接赋值 给 计算结果列
                    item = str(result[1])  # 转成字符
                    item = QStandardItem(item)  # 需要给一个str格式
                    self.itemModel32.setItem(row, 3, item)  # 结算结果real
                elif result[0] == "公式错误":
                    item = QStandardItem("")  # 需要给一个str格式
                    self.itemModel32.setItem(row, 3, item)  # 清空
                    item = QStandardItem("")
                    self.itemModel32.setItem(row, 4, item)  # 清空
                    item = QStandardItem("●")
                    self.itemModel32.setItem(row, 8, item)  # 清空
                elif result[0] == "待处理":  # 把待处理的字符放在 temp 格
                    str_temp = result[1]
                    item = QStandardItem(str_temp)  # 需要给一个str格式
                    self.itemModel32.setItem(row, 4, item)  # 清空
                    item = QStandardItem("")
                    self.itemModel32.setItem(row, 3, item)  # 清空
        # 处理可能存在的楼层引用前，先 楼层汇总一次
        print("楼层汇总第一次")
        self.noCalcuColorChange(self.itemModel32, 7, 2)  # “不计标志”变色 真实计算结果归0 部位 楼层归0
        self.floorPositionSum()  # 计算部位 楼层 小计
        # 处理 <楼层> 引用
        for row in range(rows):
            # 取出表达式的原始字符  计算结果real
            str_temp3 = self.itemModel32.item(row, 4).text().strip()
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
                    floorrowstr = self.itemModel32.item(rowr, 0).text().strip()  # 获取楼层字符
                    if floorrowstr == strkey:
                        str_temp3 = re.sub('\<[^\<]*?\>', self.itemModel32.item(rowr, 6).text(), str_temp3,
                                          count=1)  # 把第6列楼层汇总 替换成[楼层]字符，替换第一处
                        keystr = "找到了"
                        break  # 找到就跳出 楼层字符的循环 for
                if keystr == "开始找":  # 说明没找到 有错误
                    item = QStandardItem("")
                    self.itemModel32.setItem(row, 3, item)  # 赋值 给 计算结果temp列
                    item = QStandardItem("")
                    self.itemModel32.setItem(row, 4, item)
                    item = QStandardItem('●')
                    self.itemModel32.setItem(row, 8, item)
                    break  # 退出while
            # 试着处理计算式 可能有错误的引用
            try:  # 传入的非表达式  无法转换为eval则出错
                strtempstr = eval(str_temp3)  # 表达式 转结果
                strtempfloat = round(strtempstr, self.decimalPlaces)  # 保留三位小数
                strtempfin = str(strtempfloat)  # 转成字符
                item = QStandardItem(strtempfin)  # 需要给一个str格式
                self.itemModel32.setItem(row, 4, item)  # 能转化为结果 则直接赋值 给 计算结果列
                item = QStandardItem(strtempfin)
                self.itemModel32.setItem(row, 3, item)
            except:  # 去除【】 [] 以后 还不能表达式转结果  8列 错误列标记
                item = QStandardItem("")
                self.itemModel32.setItem(row, 3, item)  # 赋值 给 计算结果temp列
                item = QStandardItem("")
                self.itemModel32.setItem(row, 4, item)
                item = QStandardItem('●')
                self.itemModel32.setItem(row, 8, item)
        # <楼层> 引用处理后 再 楼层汇总一次
        print("楼层汇总第二次")
        self.noCalcuColorChange(self.itemModel32, 7, 2)  # “不计标志”变色
        self.floorPositionSum()
        self.noCalcuColorChange(self.itemModel32, 7, 2)  # “不计标志”变色
        # ~~~~~~~~~~~~~~~~~每一行数据写入总表~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        listdata = []  # 用于存放数据的空列表
        sumlist = 0  # 本条清单子目的工程量小计
        for x in range(rows):
            for y in range(columns):
                tempstr1 = self.itemModel32.item(x, y)  # 暂时不考虑 前后有空格的数据
                if y == 0:
                    listdata.append([])  # 如果是每一行的开始，则新增一个空列表
                # 不计标志 如果勾选填1 要不然填空
                if y == 7:
                    if (tempstr1.checkState() == Qt.Checked):  # 勾选了不计标志
                        item = '1'
                        listdata[x].append(item)
                    else:
                        item = ''
                        listdata[x].append(item)
                else:
                    listdata[x].append(tempstr1.text())
            strtemp = self.itemModel32.item(x, 6).text()
            if strtemp:
                sumlist += float(strtemp)
        items = str(listdata)
        item = QStandardItem(items)
        self.itemModel31.setItem(table31_row, 8, item)  # 明细写入隐藏列
        if sumlist:
            item = QStandardItem(str(sumlist))
            self.itemModel31.setItem(table31_row, 7, item)
            self.ui.label_27.setText(str(sumlist))
        else:
            item = QStandardItem("")
            self.itemModel31.setItem(table31_row, 7, item)
            self.ui.label_27.setText(str(0.0))

    # 房号表被点击时 触发
    def selectionModel0_currentChanged(self, index):
        print("房号表被点击时")
        # 房号删除按钮 可用
        self.ui.pushButton_delBuild.setEnabled(True)
        self.currrow_buildNum_str = None  # 房号名删除时 回复默认的房号名
        currrow = index.row()  # 当前行
        cols = self.itemModel0.columnCount()
        # 用于记录 房号编辑为空时 恢复原有的字符（房号不能为空）
        tempstr = self.itemModel0.item(currrow, 0).text().strip()
        if tempstr:
            self.currrow_buildNum_str = tempstr
        # 断开所有连接
        self.disconnectAll()
        # 主表 取消表的选中状态  次表初始化
        self.ui.tableView_31.clearSelection()
        self.__init_tableView_32()
        self.ui.tableView_32.setEnabled(False)  # 明细表可用
        self.ui.label_35.setText('-1')
        self.ui.label_34.setText("项名称：")
        self.ui.label_27.setText("0.00")
        # 横向 循环每一个单元格（一个单位格就是一张表）
        NumModelDict = {2: self.itemModel1, 3: self.itemModel2, 4: self.itemModel31}
        for col in NumModelDict:
            tabledatastr = self.itemModel0.item(currrow, col).text()
            model = NumModelDict.get(col)
            if tabledatastr:  # 如果表有数据
                tabledatastr = ast.literal_eval(tabledatastr)  # 字符变二维表
                # 判断 “不计标志” 的列
                noCol = None
                if model == self.itemModel1:
                    print("面积表有数据写入")
                    noCol = 7
                elif model == self.itemModel2:
                    print("措施表有数据写入")
                    noCol = None
                elif model == self.itemModel31:
                    print("分部分项表有数据写入")
                    noCol = None
                self.buildNumTableList_to_mod(tabledatastr, model, noCol)
            else:  # 如果表没有数据 则初始化
                if model == self.itemModel1:
                    self.__init_tableView_1()  # 面积表 初始化
                elif model == self.itemModel2:
                    self.__init_tableView_2()  # 措施项目清单 初始化
                elif model == self.itemModel31:
                    self.__init_tableView_31()  # 分部分项 初始化
                    self.__init_tableView_32()  # 明细表 初始化
        # 写入公共变量
        self.quantitiesDict = {}
        item = self.itemModel0.item(currrow, 5).text()
        if item:
            item = ast.literal_eval(item)
            self.quantitiesDict["历史操作"] = item
        else:
            self.quantitiesDict["历史操作"] = ""

        item = self.itemModel0.item(currrow, 6).text()
        if item:
            item = ast.literal_eval(item)
            self.quantitiesDict["工程量总表"] = item
        else:
            self.quantitiesDict["工程量总表"] = ""

        item = self.itemModel0.item(currrow, 7).text()
        if item:
            item = ast.literal_eval(item)
            self.quantitiesDict["分类表"] = item
        else:
            self.quantitiesDict["分类表"] = ""

        item = self.itemModel0.item(currrow, 8).text()
        if item:
            item = ast.literal_eval(item)
            self.quantitiesDict["钢筋工程量文件表"] = item
        else:
            self.quantitiesDict["钢筋工程量文件表"] = ""

        item = self.itemModel0.item(currrow, 9).text()
        if item:
            item = ast.literal_eval(item)
            self.quantitiesDict["钢筋工程量表"] = item
        else:
            self.quantitiesDict["钢筋工程量表"] = ""

        item = self.itemModel0.item(currrow, 10).text()
        if item:
            item = ast.literal_eval(item)
            self.quantitiesDict["绘图工程量文件表"] = item
        else:
            self.quantitiesDict["绘图工程量文件表"] = ""

        item = self.itemModel0.item(currrow, 11).text()
        if item:
            item = ast.literal_eval(item)
            self.quantitiesDict["绘图工程量表"] = item
        else:
            self.quantitiesDict["绘图工程量表"] = ""

        # 读取单体信息表
        tempstr = self.itemModel0.item(currrow, 1).text()
        if tempstr:
            tempDict = ast.literal_eval(tempstr)
        else:
            tempDict = {}
        self.singleInforSetControl(tempDict)

        # 刷新 建筑面积 统计
        self.census_area()
        # 刷新一下 措施费统计
        self.census_measures()
        # 刷新 m2m3 色标 项次变色 分部分项合价总和
        self.ui.tabWidget_2.setCurrentIndex(2)
        self.tabindex = "措施项目清单"
        self.m3ChangeRed(self.itemModel2)
        self.ui.tabWidget_2.setCurrentIndex(3)
        self.tabindex = "分部分项清单"
        self.itemModel31_summary()
        self.m3ChangeRed(self.itemModel31)
        self.boldcol0(self.itemModel31)
        self.rowsBackground(self.itemModel31)  # 色标 颜色
        # 隐藏行 显示出来
        rows = self.itemModel31.rowCount()
        for row in range(rows):
            self.ui.tableView_31.showRow(row)
        # 恢复所有连接
        self.connectAll()

    def selectionModel1_currentChanged(self, index):
        print("建筑面积表被点击！")
        self.beforestrundo = (1, self.itemModel1.item(index.row(), index.column()).text())  # 用于历史操作的原始字符
        self.ui.actEdit_Copy.setEnabled(True)  # 复制按钮可用

    def selectionModel2_currentChanged(self, index):
        print("措施表表被点击！")
        self.beforestrundo = (2, self.itemModel2.item(index.row(), index.column()).text())  # 用于历史操作的原始字符
        self.ui.actEdit_Copy.setEnabled(True)  # 复制按钮可用

    def tableView_31_clicked(self, index):
        self.ui.tableView_32.setEnabled(True)  # 明细表可用

    def selectionModel31_currentChanged(self, index):  # index当前索引 index2离开的索引
        self.ui.actEdit_Copy.setEnabled(True)  # 复制按钮可用
        self.ui.btn_Template_Suborder.setEnabled(True)  # 追加模板子目 控件不可用，防止重复按下
        # self.beforestrundo = self.itemModel31.indexFromItem(index).text()  # 用于历史操作的原始字符
        self.beforestrundo = (31, self.itemModel31.item(index.row(), index.column()).text())  # 用于历史操作的原始字符
        print(self.beforestrundo)

        print("分部分项表被点击！", index.row(), " 行")
        # print("分部分项表被点击！", self.selectionModel31.currentIndex().row(), " 行")
        self.itemModel32.itemChanged.disconnect(self.itemModel32_itemChanged)  # 明细表的数据有变化
        self.ui.tableView_32.setEnabled(True)  # 明细表可用

        rowstr = index.row()  # 主表 在哪一行
        temp8List = self.itemModel31.item(rowstr, 8).text().strip()  # 获取当前行第7列数据 明细表
        temp4 = self.itemModel31.item(rowstr, 4).text().strip()  # 获取 清单项目名称
        temp7 = self.itemModel31.item(rowstr, 7).text()  # 获取 清单 当前行工程量
        self.ui.label_35.setText(f'{rowstr + 1} ')
        self.ui.label_34.setText(f"项名称：{temp4}")
        self.ui.label_27.setText(temp7)

        if temp8List:  # 如果明细表数据不为空,则载入明细到明细表
            templist = ast.literal_eval(temp8List)  # str形式列表转化为真正的二维list
            rows = len(templist)  # 二维表共计多少行
            cols = len(templist[0])
            self.itemModel32.setRowCount(rows)  # 1 重置行数
            # print('开始循环读取 写入')
            for x in range(rows):  # 循环每一行 x 代表一行 列表形式的数据
                for y in range(cols):  # 循环每一个列  y代表每个数据
                    str1 = templist[x][y]
                    if y == 7:
                        item = QStandardItem(self.__NoCalTitle)
                        item.setFlags(self.__NoCalFlags)
                        item.setCheckable(True)  # 非锁定
                        if str1 == "1":
                            item.setCheckState(Qt.Checked)  # 勾选
                        else:
                            item.setCheckState(Qt.Unchecked)  # 勾选
                        self.itemModel32.setItem(x, y, item)  # 赋值
                    else:
                        item = QStandardItem(str1)  # 需要给一个str格式
                        self.itemModel32.setItem(x, y, item)  # 赋值
            self.noCalcuColorChange(self.itemModel32, 7, 2)  # “不计标志”变色
        else:  # 第七列数据为空  则初始化一个明细表
            rows = 10
            self.itemModel32.setRowCount(rows)
            self.initItemModelBlank(self.itemModel32)
            for i in range(rows):
                item = QStandardItem(self.__NoCalTitle)  # 最后一列
                item.setFlags(self.__NoCalFlags)
                item.setCheckable(True)  # 非锁定
                item.setCheckState(Qt.Unchecked)  # 非勾选
                self.itemModel32.setItem(i, 7, item)  # 设置最后一列的item
        # print(rowstr)
        # 选中附表的第一行为当前行
        self.ui.tableView_32.verticalScrollBar().setSliderPosition(0)

        self.itemModel32.itemChanged.connect(self.itemModel32_itemChanged)  # 明细表的数据有变化

    def selectionModel32_currentChanged(self, index):
        self.ui.actEdit_Copy.setEnabled(True)  # 复制按钮可用
        self.ui.btn_Template_Suborder.setEnabled(True)  # 追加模板子目 控件不可用，防止重复按下
        self.beforestrundo = (32, self.itemModel32.item(index.row(), index.column()).text())  # 用于历史操作的原始字符
        print(self.beforestrundo)
        print("明细表被点击！")
        mainrow = self.selectionModel31.currentIndex().row()
        test = self.itemModel31.item(mainrow, 4).text()
        self.ui.label_35.setText(f'{mainrow + 1} ')
        self.ui.label_34.setText(f"项名称：{test}")

    # def selectionModel4_currentChanged(self):
    #     print("分类表被点击！")

    # 选项卡被点击时触发 清单表 分析表
    def tabClicked_1(self, index):
        print('index', index)
        result = self.ui.tabWidget_1.tabText(index)
        print(result, "选项卡 被点击")  # 根据索引号返回选项卡的名字
        # 粘贴行 菜单不可用
        self.ui.actEdit_Paste.setEnabled(False)
        self.tabindex = result

        # 数据准备 点击后 以下为刷新
        if index == 1:  # 当分析系统被点击时
            if self.filename == None:
                self.ui.tabWidget_1.setCurrentIndex(0)
                QMessageBox.information(self, "提示", "需要先保存文件，才可以分析数据", QMessageBox.Yes)
                return
            path = os.path.splitext(self.filename)[0] + '_data'
            # 窗口刷新  控件名：[html去后缀名字 ， 随意不重复]
            # TODO  DataCal类中的  依次执行下方的生成html,并起个名字
            htmlLayoutnameDict = {self.ui.dataAna_Layout_0: ["总览表", "A"],  # 总览表
                                  self.ui.dataAna_Layout_1: ["清单分析", "B"],  # 清单分析
                                  self.ui.dataAna_Layout_21: ["土建模型分析", "C1"],  # 模型分析
                                  self.ui.dataAna_Layout_22: ["钢筋模型分析", "C2"],  # 模型分析
                                  self.ui.dataAna_Layout_3: ["土建清单报价汇总表", "D"],  # 钢筋分析
                                  self.ui.dataAna_Layout_4: ["钢筋预结算单", "E"],  # 钢筋分析
                                  }
            for layout,value in htmlLayoutnameDict.items():
                n = 1
                for i in range(layout.count()):  # 删除窗口内所有的控件
                    layout.itemAt(i).widget().deleteLater()
                value[1] = QWebEngineView()
                value[1].load(QUrl(QFileInfo(path + '\\' + value[0] + '.html').absoluteFilePath()))
                layout.addWidget(value[1])

    # 选项卡被点击时触发 面积 措施 其他 分布 分类表
    def tabClicked_2(self, index):
        result = self.ui.tabWidget_2.tabText(index)
        print(result, "选项卡 被点击")  # 根据索引号返回选项卡的名字
        # 粘贴行 菜单不可用
        self.ui.actEdit_Paste.setEnabled(False)
        self.tabindex = result
        # if result == "单体信息表":
        #     self.inforDict


    # TODO  =============Json 模块相关功能===============================
    def open_JSON_getdata(self, path):
        with open(path, "r", encoding='utf-8') as f:
            data = json.load(f)
        return data

    def save_data_toJSON(self, path, data):
        with open(path, "w", encoding='utf-8') as f:
            # indent 超级好用，格式化保存字典，默认为None，小于0为零个空格
            # sort_keys 排序 输出
            # f.write(json.dumps(dict, indent=4))
            # json.dump(dict, f, indent=1)  # 传入文件描述符，和dumps一样的结果
            json.dump(data, f, indent=4, sort_keys=True, ensure_ascii=False)  # 传入文件描述符，和dumps一样的结果

    # TODO  =============自定义菜单按钮===============================
    # 登录
    def userlogin_triggered(self):
        print("登录按钮被按下！")
        self.dlgLoginObj = None  # 用户登录的对象 初始化为空
        if (self.dlgLoginObj == None):  # 未创建对话框
            self.dlgLoginObj = QmyDialogLogin(self)
            self.dlgLoginObj.setDict()
        ret = self.dlgLoginObj.exec()  # 以模态方式运行对话框
        # ret = self.dlgLoginObj.show()  # 以非模态方式运行对话框
        if (ret == QDialog.Accepted):
            result = self.dlgLoginObj.getDict()

    # 测试1
    @pyqtSlot(bool)
    def on_actTest_triggered(self):
        print("测试按钮1")
        self.itemModel0.removeRow(0)

    # 测试2
    @pyqtSlot()
    def on_actTest2_triggered(self):
        print("测试按钮2")
        self.ui.tab_1_3.setVisible(True)

    # 测试3
    @pyqtSlot()
    def on_actTest3_triggered(self):
        print("测试按钮3")

    # 工程信息
    @pyqtSlot()
    def on_actInfor_triggered(self):
        print("工程信息被按下！")
        if (self.dlgInforObj == None):  # 未创建对话框
            self.dlgInforObj = QmyDialoginformation(self)
            self.dlgInforObj.setDict(self.inforDict)
            # self.__infordict  空字典，打开工程时定义
        ret = self.dlgInforObj.exec()  # 以模态方式运行对话框
        # ret = self.dlgInforObj.show()  # 以非模态方式运行对话框
        if (ret == QDialog.Accepted):  # 如果确定按钮被按下
            self.inforDict = self.dlgInforObj.getDict()  # 得到一个最新的信息表字典
            self.engineeringDict["工程信息"] = self.inforDict  # 更新主工程字典文件 temp 是字典
            self.ui.lineEdit_1.setText(self.inforDict['项目名称'])
            self.ui.lineEdit_2.setText(self.inforDict['所属事业部'])
            self.ui.lineEdit_3.setText(self.inforDict['项目所在省份'])
            self.ui.lineEdit_4.setText(self.inforDict['项目所在市'])
            self.ui.lineEdit_5.setText(self.inforDict['地区类型'])
            self.ui.lineEdit_6.setText(self.inforDict['设计院名称'])
            self.ui.lineEdit_7.setText(self.inforDict['人防设计院名称'])

    # 打开工程
    @pyqtSlot()
    def on_actOpen_triggered(self):
        print("打开工程")
        if self.filename:
            curPath = self.filename
        else:
            curPath = QDir.currentPath()  # 获取系统当前目录
        dlgTitle = "打开诚清单文件"  # 对话框标题
        filt = "工程文件(*.cqd);;所有文件(*.*)"  # 文件过滤器
        # filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        if filename:
            self.filename = filename
            print('filename地址真实有效！', self.filename)
            # self.filename = filename
            self.setWindowTitle(self.edition + '    ' + str(self.filename))  # 设置窗口名称为 全路径

            # 获取 JSON 数据
            data = self.open_JSON_getdata(self.filename)
            self.engineeringDict = data
            result = self.engineeringDict.get('工程信息')
            self.dlgInforObj = None  # 工程信息的对象 初始化为空
            if result:
                self.inforDict = result
            else:
                self.inforDict = {'项目名称': '', '所属事业部': '', '项目所在省份': '', '项目所在市': '', '地区类型': '',
                                  '设计院名称': '', '人防设计院名称': ''}
            self.engineeringInforSetControl(self.inforDict)  # 打开工程时 工程信息 写入每个控件
            # data 数据写入模型
            datalist2 = self.engineeringDict.get('房号信息')  # 二维list
            # 断开连接
            self.disconnectAll()
            # 初始化一些数据
            # self.selectionModel0.clear()
            # self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 清单表页面
            # self.ui.tabWidget_2.setCurrentIndex(3)  # 切换到tabindex 3 分部分项清单页面
            # self.__init_tableView_0()  # 房号表 初始化
            # # self.__init_tableView_1()  # 清单表 面积表 初始化
            # # self.__init_tableView_2()  # 清单表 措施项目清单 初始化
            # self.__init_tableView_31()  # 清单表 分部分项 初始化
            # self.__init_tableView_32()  # 清单表 分部分项 明细表 初始化
            # self.__init_listView_undo()  # 撤销栏 历史操作 初始化
            self.templist = []  # 复制粘贴行用的临时列表变量
            # self.singleInforDict = {}  # 单体信息  写入房号表第二个单元格
            # self.quantitiesDict = {}
            # self.engineeringDict = {'工程信息': self.inforDict}
            # 写入数据
            rows = len(datalist2)
            cols = len(datalist2[0])
            self.itemModel0.setRowCount(rows)
            for row in range(rows):
                list1 = []
                for col in range(cols):
                    item = QStandardItem(datalist2[row][col])
                    self.itemModel0.setItem(row, col, item)
            # 刷新单体信息的多个控件
            singleinforstr = datalist2[0][1]
            if singleinforstr:
                sDict = ast.literal_eval(singleinforstr)  # str还原成Dict
                self.singleInforSetControl(sDict)
            # 恢复连接
            self.connectAll()
            # 刷新用
            self.selectionModel0_currentChanged(self.itemModel0.index(0, 0))
            # self.on_pushButton_addBuild_clicked()
            # rows = self.itemModel0.rowCount()
            # self.selectionModel0.setCurrentIndex(self.itemModel0.index(rows - 1, 0), QItemSelectionModel.Select)
            # self.itemModel0.removeRow(rows - 1)

    # 保存工程 按钮
    @pyqtSlot()
    def on_actSave_triggered(self):
        print("保存工程")
        if self.filename == None:
            # print("第一次保存 self.filename")
            curPath = QDir.currentPath()  # 获取系统当前目录
            dlgTitle = "保存文件"  # 对话框标题
            filt = "工程文件(*.cqd);;所有文件(*.*)"  # 文件过滤器

            filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
            if filename:
                self.filename = filename
        if self.filename != None:
            # print('filename地址真实有效！', self.filename)
            # self.filename = filename
            self.setWindowTitle(self.edition + '    ' + str(self.filename))  # 设置窗口名称为 全路径
            # 获取data 数据
            buildNumList = []
            rowstable0 = self.itemModel0.rowCount()
            colstable0 = self.itemModel0.columnCount()
            for row in range(rowstable0):
                list = []
                for col in range(colstable0):
                    item = self.itemModel0.item(row, col).text()
                    list.append(item)
                buildNumList.append(list)
            self.engineeringDict['房号信息'] = buildNumList
            # print("保存前打印一下工程总表字典： \n ", self.engineeringDict)
            # 备份机制
            backmainf = '.\data\_backfolder'
            if not os.path.exists(backmainf):
                os.mkdir(backmainf)
            backsubf = backmainf + '\\' + os.path.basename(self.filename).split('.')[0]
            if not os.path.exists(backsubf):
                os.mkdir(backsubf)
            backfullpath = backsubf + '\\' + os.path.basename(self.filename).split('.')[0] + \
                           str(time.strftime("(%Y-%m-%d_%H-%M-%S)", time.localtime())) + ".cqd"
            self.save_data_toJSON(backfullpath, self.engineeringDict)
            # 存JSON
            self.save_data_toJSON(self.filename, self.engineeringDict)
            # self.backpickle(self.engineeringDict)  # 备份文件
            # QMessageBox.critical(self, "错误", "表格行未被选中，无法复制行", QMessageBox.Cancel)
            QMessageBox.information(self,"提示","保存成功！",QMessageBox.Ok)

    # 另存为工程
    @pyqtSlot()
    def on_actSaveAs_triggered(self):
        print("另存工程文件")
        self.filename = None  # 文件保存路径
        self.on_actSave_triggered()

    # 新建工程
    @pyqtSlot()
    def on_actNew_triggered(self):
        # if self.filename != None:
        res = QMessageBox.warning(self, "提醒", "是否新建工程！", QMessageBox.Yes | QMessageBox.No)
        if QMessageBox.Yes == res:
            print("点了确认！")
            # self.on_actSave_triggered()  # 先保存一次
        elif QMessageBox.No == res:
            print("点了否！")
            return
        # 断开连接
        self.disconnectAll()
        # 新建工程
        self.setWindowTitle(self.edition)
        # 公共变量
        self.filename = None  # 文件保存路径
        self.templist = []  # 复制粘贴行用的临时列表变量
        self.dlgInforObj = None  # 工程信息的对象 初始化为空
        self.inforDict = {'项目名称': '', '所属事业部': '', '项目所在省份': '', '项目所在市': '', '地区类型': '',
                          '设计院名称': '', '人防设计院名称': ''}
        self.engineeringInforSetControl(self.inforDict)  # 打开工程时 工程信息 写入每个控件
        self.singleInforDict = {}  # 单体信息  写入房号表第二个单元格
        self.singleInforSetControl(self.singleInforDict)
        self.quantitiesDict = {}
        self.engineeringDict = {'工程信息': self.inforDict}
        # 控件初始化
        self.ui.tabWidget_1.setCurrentIndex(0)  # 切换到tabindex 0 清单表页面
        self.ui.tabWidget_2.setCurrentIndex(3)  # 切换到tabindex 3 分部分项清单页面
        self.__init_tableView_0()  # 房号表 初始化
        # self.__init_tableView_1()  # 清单表 面积表 初始化
        # self.__init_tableView_2()  # 清单表 措施项目清单 初始化
        # self.__init_tableView_31()  # 清单表 分部分项 初始化
        # self.__init_tableView_32()  # 清单表 分部分项 明细表 初始化
        # self.__init_listView_undo()  # 撤销栏 历史操作 初始化
        # 恢复连接
        self.connectAll()
        # 刷新用
        self.selectionModel0_currentChanged(self.itemModel0.index(0, 0))

    # 查找清单
    @pyqtSlot()
    def on_actFindList_triggered(self):
        self.ui.tabWidget_1.setCurrentIndex(0)
        self.ui.tabWidget_2.setCurrentIndex(3)
        print("搜索按钮")
        text, okPressed = QInputDialog.getText(self, "分部分项清单—搜索",
                                               "搜索范围包括：色标 地上地下 编码 名称 特征 单位 备注以及 对应的明细表\n请输入关键词:（分隔 ， 中英文混用皆可）",
                                               QLineEdit.Normal, "")
        rows = self.itemModel31.rowCount()
        if okPressed and text != '':
            print("OK")
            list1 = text.split(sep='，')
            list1 = ','.join(list1)
            keyList = list1.split(sep=',')
            for row in range(rows):
                itemstr0 = self.itemModel31.item(row, 0).text()
                itemstr1 = self.itemModel31.item(row, 1).text()
                itemstr2 = self.itemModel31.item(row, 2).text()
                itemstr3 = self.itemModel31.item(row, 3).text()
                itemstr4 = self.itemModel31.item(row, 4).text()
                itemstr5 = self.itemModel31.item(row, 5).text()
                itemstr6 = self.itemModel31.item(row, 6).text()
                itemstr8 = self.itemModel31.item(row, 8).text()
                itemstr11 = self.itemModel31.item(row, 11).text()
                str1_8 = itemstr0 + itemstr1 + itemstr2 + itemstr3 + itemstr4 + itemstr5 + itemstr6 + itemstr8 + itemstr11
                if str1_8 == "":
                    self.ui.tableView_31.hideRow(row)
                    continue
                else:
                    for keystr in keyList:
                        if keystr in str1_8:
                            self.ui.tableView_31.showRow(row)
                            break  # 跳出for
                        else:
                            self.ui.tableView_31.hideRow(row)
        # 如果搜索的是空格，或者 按下了cancel按钮则
        elif not okPressed or text == "":
            print("Cancel")
            for row in range(rows):
                self.ui.tableView_31.showRow(row)

    # 智能导入清单(XC) 智能导入清单（新城控股标准清单措施项目的导入）
    @pyqtSlot()
    def on_actImportIQ_triggered(self):
        res = QMessageBox.information(self, "提醒", "《新城标准清单》 一键导入，将会清空当前措施项目清单的数据！\n 是否现在导入？",
                                      QMessageBox.Yes | QMessageBox.No)
        if QMessageBox.Yes == res:
            print("点了确认！")
            # self.on_actSave_triggered()  # 先保存一次
        elif QMessageBox.No == res:
            print("点了否！")
            return
        # 获取路径
        if self.filename == None:
            curPath = QDir.currentPath()  # 获取系统当前目录
        else:
            curPath = self.filename
        dlgTitle = "请选择《新城标准清单》"  # 对话框标题
        filt = "清单文件(*.xls *.xlsx);;所有文件(*.*)"  # 文件过滤器
        # filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        if not filename:
            QMessageBox.information(self, "提示", "不想一键导就算了，想好再点！", QMessageBox.Cancel)
            return
        # 打开文件 前 先做清空
        # 断开 措施费模型连接
        # 断开 措施费模型连接
        self.disconnectAll()
        # self.itemModel2.itemChanged.disconnect(self.itemModel2_itemChanged)
        self.__init_tableView_2()
        # 措施项目清单 写入 房号表
        # self.mod_to_buildNumTable(self.itemModel2)
        # self.on_actNew_triggered()  # 新建工程
        sheetnamekeyDict = {'整体措施费': ['整体措施'],
                            '单项措施费': ['单项措施', '措施费汇总表'],
                            '其他项目清单': ['其他项目', '其它项目'],
                            # '分部分项清单': ['工程量清单及计价表', '工程量清单', '分部分项'],
                            # '分部分项清单PC': ['PC构件', 'PC工程量']
                            }
        # 打开文件
        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
            app.screen_updating = False  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
            wb = app.books.open(filename, update_links=False)  # 更新连接 否
            shts = wb.sheets  # 所有工作表对象的列表
        except:
            app.quit()
            QMessageBox.critical(self, "错误", "未检测到电脑的OFFICE软件！\n请勿使用wps处理过的表格文件！否则可能无法识别！", QMessageBox.Cancel)
            return
        self.measureslist = []  # 存放措施项目的列表
        self.subitemlist = []  # 存放分部分项的列表
        # sheetname 判断是哪张表的key名
        for sheetname, sheetnamekeylist in sheetnamekeyDict.items():  # 循环执行要提取表的名字
            flag = False
            #  shtname 用来判断的备选名
            for shtname in sheetnamekeylist:
                if flag:  # 跳出 for shtname
                    break
                for sht in shts:  # 循环每个工作表的 对象
                    # print(type(sht), sht)
                    # print(sht.visibility)
                    if sht.api.Visible != -1:  # 工作表隐藏  -1显示 0隐藏 2深度隐藏
                        continue
                    if shtname in sht.name:  # 如果找到想要的表
                        self.handleEachDetailed_list(sheetname, sht.name, sht)  # 子程序处理
                        flag = True  # 跳出 for shtname
                        break  # 跳出 for num in range(sheetsNum): 循环
        wb.close()
        app.display_alerts = True  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        app.screen_updating = True  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        app.quit()
        # 写入模型
        if len(self.measureslist) < 1:  # 存放措施项目的列表
            QMessageBox.critical(self, "错误", "啥也没提取到！", QMessageBox.Cancel)
            return
        # print('self.measureslist', self.measureslist)
        # 写入模型
        nums = len(self.measureslist)
        mainrow = self.itemModel2.rowCount()
        for row in range(nums):
            templist = []
            for col in range(len(self.measureslist[0])):
                item = self.measureslist[row][col]
                if item:
                    item = QStandardItem(str(item))
                else:
                    item = QStandardItem('')
                templist.append(item)
            self.itemModel2.insertRow(mainrow, templist)
            mainrow += 1
        # 辅助 措施表 数据有变化 处理所有行行数据
        self.itemModel2_itemChanged_All()
        # 刷新一下 措施费统计
        self.census_measures()
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel2)
        # m2m3变色
        self.ui.tabWidget_2.setCurrentIndex(2)
        self.tabindex = "措施项目清单"
        self.m3ChangeRed(self.itemModel2)
        # 恢复 措施费模型连接
        # self.itemModel2.itemChanged.connect(self.itemModel2_itemChanged)
        self.connectAll()

    # 处理每一张有效的 清单文件内的表  辅助 智能导入清单(XC) 一键导入
    def handleEachDetailed_list(self, sheetname,  shtrealname, shtOBJ):
        try:
            meablanklist = ['' for _ in range(11)]  # 措施费的空行
            # subblanklist = ['' for _ in range(20)]  # 分部分项的空行
            rows = shtOBJ.used_range.row + shtOBJ.used_range.rows.count - 1  # 获取最大行号 推荐用法
            cols = shtOBJ.used_range.column + shtOBJ.used_range.columns.count - 1  # 获取最大行号 推荐用法
            if sheetname == "整体措施费":  # 只处理 一行有效数据的格式
                # 找关键列号
                cruxColNum = {'单价': [["含税综合单价", "综合单价"], '', '', ''], '工程量': [["招标业态面积", "工程量"], '', '', '']}
                # [备选名字],'','',''  备选, 行号, 列号,存放数据
                for key, value in cruxColNum.items():
                    flag = False
                    for kstr in value[0]:  # 循环每一个备选关键字
                        if flag == True:
                            break
                        for row in range(2, 11):  # 3~10行 循环
                            if flag == True:
                                break
                            result = shtOBJ.range(row, 1).row_height  # 隐藏的行高为0 既可以排除隐藏 又可以排除小行高的行
                            if result < 3:
                                continue
                            for col in range(1, 20):
                                result2 = shtOBJ.range(1, col).column_width  # 返回列宽
                                if result2 < 0.8:  # 列宽 小于0.8 基本不可见，隐藏列的列宽为0
                                    continue
                                cellstr = shtOBJ.range(row,col).value
                                if cellstr == None:
                                    continue
                                elif kstr in str(cellstr):
                                    cruxColNum[key][2] = col  # 写入列号
                                    cruxColNum[key][3] = shtOBJ.range(row + 1, col).value  # 写入数值
                                    flag = True
                                    break
                print(cruxColNum)
                # 设置模型的值
                result = cruxColNum.get('单价')[3]
                result2 = cruxColNum.get('工程量')[3]
                if result2 and type(result2) == float:
                    result2 = round(result2, 2)
                else:
                    result2 = ''
                if result != "" and type(result) == float:
                    result = round(result, 5)  # 单价小数位数
                    templist = ['1', '整体措施费', '', '', 'm2', result2, result2, '', result, '', '']
                    self.measureslist.append(templist)
                    self.measureslist.append(meablanklist)  # 加个空行

            elif sheetname == "单项措施费":
                # 找关键列号
                cruxColNum = {'项目': [["项目"], '', '', ''], '部位': [["部位"], '', '', ''],
                              '单位': [['单位', f'计量\n单位','计量单位'], '', '', ''],
                              '单价': [[f"含税综合单价\n(元/m2)", '基准含税综合单价', "含税综合单价", "综合单价"], '', '', ''],
                              '工程量': [["工程量"], '', '', '']}
                # [备选名字],'','',''  备选, 行号, 列号,存放数据
                for key, value in cruxColNum.items():
                    flag = False
                    for kstr in value[0]:  # 循环每一个备选关键字
                        if flag == True:
                            break
                        for row in range(2, 11):  # 3~10行 循环
                            if flag == True:
                                break
                            result = shtOBJ.range(row, 1).row_height  # 隐藏的行高为0 既可以排除隐藏 又可以排除小行高的行
                            if result < 3:
                                continue
                            for col in range(1, 20):
                                result2 = shtOBJ.range(1, col).column_width  # 返回列宽
                                if result2 < 1:  # 列宽 小于0.8 基本不可见，隐藏列的列宽为0
                                    continue
                                cellstr = shtOBJ.range(row, col).value
                                if cellstr == None:
                                    continue
                                elif kstr == str(cellstr):
                                    cruxColNum[key][1] = row  # 写入行号
                                    cruxColNum[key][2] = col  # 写入列号
                                    flag = True
                                    break
                # print('单项措施费', cruxColNum)
                # 循环读取数据
                n = 0
                flag = False
                str31 = ""
                for row in range(3, rows+1):
                    result = shtOBJ.range(row, 1).row_height  # 隐藏的行高为0 既可以排除隐藏 又可以排除小行高的行
                    if result < 3:
                        continue

                    resultcol = cruxColNum.get('项目')[2]
                    if resultcol:
                        project = shtOBJ.range(row, resultcol).value  # 模板 垂直 脚手架等
                        if project:
                            str2 = project
                            if "合计" in project:
                                break

                    resultcol = cruxColNum.get('部位')[2]
                    if resultcol:
                        place = shtOBJ.range(row, resultcol).value  # 高层 小高层等
                        if place:
                            str31 = place
                        # else:
                        #     continue  # 没有 部位 则跳过

                        str32 = shtOBJ.range(row, resultcol + 1).value  # 木模 铝模  必须要有
                        if str31 and str32:
                            str3 = str(str31) + '_' + str(str32)
                        else:
                            str3 = str31

                    resultcol = cruxColNum.get('单价')[2]
                    if resultcol:
                        str8 = shtOBJ.range(row, resultcol).value  #必须要有
                        if type(str8) == float:
                            str8 = round(str8, 5)
                        elif str8 == None:  # 只跳过空行
                            # pass
                            continue

                    resultcol = cruxColNum.get('单位')[2]
                    if resultcol:
                        str4 = shtOBJ.range(row, resultcol).value  # 必须要有

                    resultcol = cruxColNum.get('工程量')[2]
                    if resultcol:
                        str5 = shtOBJ.range(row, resultcol).value  #必须要有  str5=str6
                        if str5 != "" and type(str5) == float:
                            str5 = round(str5, 3)
                        else:
                            str5 = ''
                    # 判断是否拼接成一条子目  单价为数值 且不为0  有项目名 有部位名
                    if str8 != "" and str2 and str31:
                        n += 1
                        templist = [n, '单项措施费', str2, str3, str4, str5, str5, '', str8, '', '']
                        self.measureslist.append(templist)
                        flag = True
                if flag == True:
                    self.measureslist.append(meablanklist)

            elif sheetname == "其他项目清单":
                # 找关键列号
                cruxColNum = {'项目': [["项目"], '', '', ''],
                              '单位': [['单位', f'计量\n单位','计量单位'], '', '', ''],
                              '单价': [[f"含税综合单价\n(元/m2)", '基准含税综合单价', "含税综合单价", "综合单价"], '', '', ''],
                              '工程量': [["工程量"], '', '', '']}
                # [备选名字],'','',''  备选, 行号, 列号,存放数据
                for key, value in cruxColNum.items():
                    flag = False
                    for kstr in value[0]:  # 循环每一个备选关键字
                        if flag == True:
                            break
                        for row in range(2, 11):  # 3~10行 循环
                            if flag == True:
                                break
                            result = shtOBJ.range(row, 1).row_height  # 隐藏的行高为0 既可以排除隐藏 又可以排除小行高的行
                            if result < 3:
                                continue
                            for col in range(1, 20):
                                result2 = shtOBJ.range(1, col).column_width  # 返回列宽
                                if result2 < 1:  # 列宽 小于0.8 基本不可见，隐藏列的列宽为0
                                    continue
                                cellstr = shtOBJ.range(row, col).value
                                if cellstr == None:
                                    continue
                                elif kstr == str(cellstr):
                                    cruxColNum[key][1] = row  # 写入行号
                                    cruxColNum[key][2] = col  # 写入列号
                                    flag = True
                                    break
                # print('单项措施费', cruxColNum)
                # 循环读取数据
                n = 0
                flag = False
                for row in range(3, rows+1):
                    result = shtOBJ.range(row, 1).row_height  # 隐藏的行高为0 既可以排除隐藏 又可以排除小行高的行
                    if result < 3:
                        continue

                    resultcol = cruxColNum.get('项目')[2]
                    if resultcol:
                        project = shtOBJ.range(row, resultcol).value  # 模板 垂直 脚手架等
                        if project:
                            str2 = project

                    resultcol = cruxColNum.get('单位')[2]
                    if resultcol:
                        str4 = shtOBJ.range(row, resultcol).value  # 必须要有

                    resultcol = cruxColNum.get('单价')[2]
                    if resultcol:
                        str8 = shtOBJ.range(row, resultcol).value  #必须要有
                        if str8 and type(str8) == float:
                            str8 = round(str8, self.decimalPlaces)
                        else:
                            str8 = ''

                    resultcol = cruxColNum.get('工程量')[2]
                    if resultcol:
                        str5 = shtOBJ.range(row, resultcol).value  #必须要有  str5=str6
                        if str5 and type(str5) == float:
                            str5 = round(str5, self.decimalPlaces)
                        else:
                            str5 = ''

                    # 判断是否拼接成一条子目  单价为数值 且不为0  有项目名 有部位名
                    if str8 and str2:
                        n += 1
                        templist = [n, '其他项目清单', str2, '', str4, str5, str5, '', str8, '', '']
                        self.measureslist.append(templist)
                        flag = True
                if flag == True:
                    self.measureslist.append(meablanklist)

            # elif sheetname == "分部分项清单":
            #     # 找关键列号
            #     cruxColNum = {'序号': [["编号"], '', '', ''],
            #                   '项目名称': [["内容"], '', '', ''],
            #                   '项目特征描述': [["项目特征"], '', '', ''],
            #                   '计量单位': [['计量单位', '计量\n单位','单位'], '', '', ''],
            #                   '工程量': [["工程量"], '', '', ''],
            #                   '综合单价': [["含税综合单价(RMB)", "含税综合单价", "基准含税综合单价(RMB)"], '', '', ''],
            #                   '备注': [["备注"], '', '', ''],
            #                   '人工费': [["人工费"], '', '', ''],
            #                   '主材费': [["材料费"], '', '', ''],
            #                   '辅材费': [["辅材费"], '', '', ''],
            #                   '机械费': [["机械费"], '', '', ''],
            #                   '管理费、利润': [["管理费、利润"], '', '', ''],
            #                   '规范、措施费': [["规费"], '', '', ''],
            #                   '不含税综合单价': [["除税综合单价\n(RMB)"], '', '', ''],
            #                   '增值税税金': [["增值税税金(B)"], '', '', ''],
            #                   }
            #     # [备选名字],'','',''  备选, 行号, 列号,存放数据
            #     for key, value in cruxColNum.items():
            #         flag = False
            #         for kstr in value[0]:  # 循环每一个备选关键字
            #             if flag == True:
            #                 break
            #             for row in range(3, 10):  # 3~10行 循环
            #                 if flag == True:
            #                     break
            #                 result = shtOBJ.range(row, 1).row_height  # 隐藏的行高为0 既可以排除隐藏 又可以排除小行高的行
            #                 if result < 3:
            #                     continue
            #                 for col in range(1, 40):
            #                     result2 = shtOBJ.range(1, col).column_width  # 返回列宽
            #                     if result2 < 1:  # 列宽 小于0.8 基本不可见，隐藏列的列宽为0
            #                         continue
            #                     cellstr = shtOBJ.range(row, col).value
            #                     if cellstr == None:
            #                         continue
            #                     elif kstr == str(cellstr):
            #                         cruxColNum[key][1] = row  # 写入行号
            #                         cruxColNum[key][2] = col  # 写入列号
            #                         flag = True
            #                         break
            #     # print('单项措施费', cruxColNum)
            #     # 循环读取数据
            #     # n = 0
            #     flag = False
            #     for row in range(3, rows+1):
            #         if row > 1000:  # 大于1000行的清单 截取
            #             break
            #         # # 隐藏的行高为0 既可以排除隐藏 又可以排除小行高的行  暂时取消
            #         # result = shtOBJ.range(row, 1).row_height
            #         # if result < 3:
            #         #     continue
            #         resultcol = cruxColNum.get('序号')[2]
            #         if resultcol:
            #             str0 = shtOBJ.range(row, resultcol).value
            #         else:
            #             str0 = ''
            #
            #         resultcol = cruxColNum.get('项目名称')[2]  # 可能有竖向合并单元格
            #         if resultcol:  # 找到列号
            #             res = shtOBJ.range(row, resultcol).value
            #             if res:
            #                 str4 = res
            #                 flag = True
            #             elif flag == False:  # 从来没有找到过
            #                 str4 = ''
            #         else:
            #             str4 = ''
            #
            #         resultcol = cruxColNum.get('项目特征描述')[2]
            #         if resultcol:
            #             str5 = shtOBJ.range(row, resultcol).value
            #         else:
            #             str5 = ''
            #
            #         resultcol = cruxColNum.get('计量单位')[2]
            #         if resultcol:
            #             str6 = shtOBJ.range(row, resultcol).value
            #         else:
            #             str6 = ''
            #
            #         resultcol = cruxColNum.get('工程量')[2]
            #         if resultcol:
            #             str7 = shtOBJ.range(row, resultcol).value
            #             if str7 and type(str7) == float:
            #                 str7 = round(str7, self.decimalPlaces)
            #             else:
            #                 str7 = ''
            #         else:
            #             str7 = ''
            #
            #         resultcol = cruxColNum.get('综合单价')[2]
            #         if resultcol:
            #             str9 = shtOBJ.range(row, resultcol).value  #必须要有
            #             if str9 and type(str9) == float:
            #                 str9 = round(str9, self.decimalPlaces)
            #             else:
            #                 str9 = ''
            #         else:
            #             str9 = ''
            #
            #         resultcol = cruxColNum.get('备注')[2]
            #         if resultcol:
            #             str11 = shtOBJ.range(row, resultcol).value
            #         else:
            #             str11 = ''
            #
            #         resultcol = cruxColNum.get('人工费')[2]
            #         if resultcol:
            #             str12 = shtOBJ.range(row, resultcol).value
            #             if str12 and type(str12) == float:
            #                 str12 = round(str12, self.decimalPlaces)
            #             else:
            #                 str12 = ''
            #         else:
            #             str12 = ''
            #
            #         resultcol = cruxColNum.get('主材费')[2]
            #         if resultcol:
            #             str13 = shtOBJ.range(row, resultcol).value
            #             if str13 and type(str13) == float:
            #                 str13 = round(str13, self.decimalPlaces)
            #             else:
            #                 str13 = ''
            #         else:
            #             str13 = ''
            #
            #         resultcol = cruxColNum.get('辅材费')[2]
            #         if resultcol:
            #             str14 = shtOBJ.range(row, resultcol).value
            #             if str14 and type(str14) == float:
            #                 str14 = round(str14, self.decimalPlaces)
            #             else:
            #                 str14 = ''
            #         else:
            #             str14 = ''
            #
            #         resultcol = cruxColNum.get('机械费')[2]
            #         if resultcol:
            #             str15 = shtOBJ.range(row, resultcol).value
            #             if str15 and type(str15) == float:
            #                 str15 = round(str15, self.decimalPlaces)
            #             else:
            #                 str15 = ''
            #         else:
            #             str15 = ''
            #
            #         resultcol = cruxColNum.get('管理费、利润')[2]
            #         if resultcol:
            #             str16 = shtOBJ.range(row, resultcol).value
            #             if str16 and type(str16) == float:
            #                 str16 = round(str16, self.decimalPlaces)
            #             else:
            #                 str16 = ''
            #         else:
            #             str16 = ''
            #
            #         resultcol = cruxColNum.get('规范、措施费')[2]
            #         if resultcol:
            #             str17 = shtOBJ.range(row, resultcol).value
            #             if str17 and type(str17) == float:
            #                 str17 = round(str17, self.decimalPlaces)
            #             else:
            #                 str17 = ''
            #         else:
            #             str17 = ''
            #
            #         resultcol = cruxColNum.get('不含税综合单价')[2]
            #         if resultcol:
            #             str18 = shtOBJ.range(row, resultcol).value
            #             if str18 and type(str18) == float:
            #                 str18 = round(str18, self.decimalPlaces)
            #             else:
            #                 str18 = ''
            #         else:
            #             str18 = ''
            #
            #         resultcol = cruxColNum.get('增值税税金')[2]
            #         if resultcol:
            #             str19 = shtOBJ.range(row, resultcol).value
            #             if str19 and type(str19) == float:
            #                 str19 = round(str19, self.decimalPlaces)
            #             else:
            #                 str19 = ''
            #         else:
            #             str19 = ''
            #         # 所有行 包含隐藏行 拼接成一条子目
            #         templist = [str0, '', '', '', str4, str5, str6, str7, '', str9, '',
            #                     str11, str12, str13, str14, str15, str16, str17, str18, str19]
            #         self.subitemlist.append(templist)
            #     if flag == True:
            #         self.subitemlist.append(subblanklist)
            #
            # elif sheetname == "分部分项清单PC":
            #     pass
        except:
            QMessageBox.critical(self, "错误", "只支持《新城公司》的标准化格式\n部分措施项目可能未导入成功，请检查！", QMessageBox.Cancel)
            return

    # 导入清单
    @pyqtSlot()
    def on_actImport_triggered(self):
        print("导入 清单")
        self.tabindex = "分部分项清单"
        self.ui.tabWidget_1.setCurrentIndex(0)
        self.ui.tabWidget_2.setCurrentIndex(3)
        dlgTableExcel = QmyDialogImportExcel()  # 局部变量，构建时不能传递self
        ret = dlgTableExcel.exec()  # 模态方式运行对话框
        if (ret == QDialog.Accepted):
            detailedData, datarow = dlgTableExcel.getData()  # 返回匹配好的 二维列表 需要转置
            if not detailedData or not datarow:
                return
            # 二维表转置
            # print(detailedData)
            detailedData = self.transpose_2d(detailedData)
            print(detailedData)
            zrow = self.itemModel31.rowCount()  # 主表的最大行
            cols = self.itemModel31.columnCount()
            for row in range(len(detailedData)):
                itemlist = []
                for col in range(len(detailedData[row])):
                    item = detailedData[row][col]
                    if col == 7 or col == 9:   # 如果是工程量 单价 列  处理小数位
                        try:
                            item = float(item)
                            item = round(item, self.decimalPlaces)
                            item = str(item)
                        except:
                            item = str(item)
                    else:
                        item = str(item)
                    item = QStandardItem(item)
                    itemlist.append(item)
                self.itemModel31.appendRow(itemlist)
            QMessageBox.information(self, "成功导入:", f"{datarow} 行清单表文件！\n 开始计算！", QMessageBox.Ok)
            # 断开连接
            # self.disconnectAll()
            self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)
            print('写入数据后的一些总计 颜色的处理')
            self.itemModel31_summary()  # 汇总清单表合价总和
            self.m3ChangeRed(self.itemModel31)  # m3 m2 变色
            self.boldcol0(self.itemModel31)  # 粗体 序号 为ABCD...  一二三...开头  分部分项表
            self.rowsBackground(self.itemModel31)  # 色标 颜色
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel31)
            # 恢复连接
            # self.connectAll()
            self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged)

    # 导出清单
    @pyqtSlot()
    def on_actExport_triggered(self):
        print("导出清单表")
        curPath = QDir.currentPath()  # 获取系统当前目录
        dlgTitle = "导出文件"  # 对话框标题
        filt = "工程文件(*.xls)"  # 文件过滤器
        print('1')
        filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        print('2')
        if not filename:
            return
        # self.ui.tabWidget.setCurrentIndex(0)  # 切换到tabindex 0 页面
        # self.tabindex = 0  # 清单表
        # buildnamelist = []
        buildrows = self.itemModel0.rowCount()
        sheetnameDict = {2: '建筑面积计算表', 3: '措施项目清单', 4: '分部分项清单',
                         6: '工程量总表', 7: '分类表', 9: '钢筋工程量表', 11: '绘图工程量表'}
        # 创建一个workbook 设置编码
        workbook = xlwt.Workbook(encoding='utf-8')
        flag = False
        for buildrow in range(buildrows):  # 循环每个房号
            buildname = self.itemModel0.item(buildrow, 0).text()

            for buildcol, sheetname in sheetnameDict.items():  # 循环每一个表
                datasstr = self.itemModel0.item(buildrow, buildcol).text()
                if not datasstr:
                    continue
                else:
                    datas = ast.literal_eval(datasstr)
                # 1、创建一个worksheet
                flag = True
                worksheet = workbook.add_sheet(buildname + '_' + sheetname)  # 房号加表名
                # 获取表头列表
                if sheetname == "建筑面积计算表":
                    headerList = self.tableView_1_headerList
                    colwidth = [1500, 3000, 3000, 1500, 12000, 3000, 3000, 3000, 3000, 3000, 1]  # *11
                    noCollist = [10]

                elif sheetname == "措施项目清单":
                    headerList = self.tableView_2_headerList
                    colwidth = [1500, 4000, 6000, 3000, 3000, 8000, 3000, 4000, 4000, 3000, 5000]  # *11
                    noCollist = [999]

                elif sheetname == "分部分项清单":
                    headerList = self.tableView_31_headerList
                    colwidth = [5000, 2000, 2000, 3000, 5000, 9000, 3000, 3000, 1, 3000,
                                3000, 3000, 3000, 3000, 3000, 3000, 4000, 4000, 4000, 3000]  # *20
                    noCollist = [8]

                elif sheetname == "工程量总表":
                    headerList = ['序号', '来源', '引用编码\n项目名称', '项目名称', '项目特征描述', '计量\n单位', '工程量', '∑工程量明细表\n隐藏', '备注']
                    colwidth = [1500, 4000, 6000, 1, 5000, 3000, 3000, 1, 5000]  # *9
                    noCollist = [3, 7]

                elif sheetname == "分类表":
                    headerList = ['序号', '楼层/部位', '清单名称', '计量\n单位', '计算表达式\n中文括号【注释】',
                      '工程量', '不计\n标志', '公式\n错误', '备注']
                    colwidth = [1500, 4000, 12000, 3000, 8000, 3000, 3000, 3000, 5000]  # *9
                    noCollist = [999]

                elif sheetname == "钢筋工程量表":
                    headerList = ['序号', '来源', '引用编码\n隐藏', '报表名称', '特征描述', '计量\n单位', '工程量', '∑工程量明细表\n隐藏', '备注']
                    colwidth = [1500, 4000, 1, 5000, 6000, 3000, 3000, 1, 5000]  # *9
                    noCollist = [2, 7]

                elif sheetname == "绘图工程量表":
                    headerList = ['序号', '来源', '引用编码\n隐藏', '报表名称', '特征描述', '计量\n单位', '工程量', '∑工程量明细表\n隐藏', '备注']
                    colwidth = [1500, 4000, 1, 5000, 6000, 3000, 3000, 1, 5000]  # *9
                    noCollist = [2, 7]
                # 2、写入表头
                for hcol in range(len(headerList)):
                    if hcol in noCollist:
                        continue  # 跳过隐藏的列
                    worksheet.write(0, hcol, label=headerList[hcol])
                # 3、写入数据
                rows = len(datas)
                cols = len(datas[0])
                for col in range(cols):
                    if col in noCollist:
                        continue  # 跳过隐藏的列
                    for row in range(rows):
                        item = datas[row][col]
                        worksheet.write(row + 1, col, label=item)
                    # 4、格式列宽
                    worksheet.col(col).width = colwidth[col]
        if flag == True:
            # 保存
            try:
                workbook.save(filename)
                print("导出文件成功")
                QMessageBox.information(self, "成功", "清单表导出成功", QMessageBox.Ok)
            except:
                QMessageBox.warning(self, "导出失败", "请确保路径下的文件处于“非打开”状态！", QMessageBox.Ok)
        else:
            QMessageBox.warning(self, "失败", "没有数据，要导出什么？", QMessageBox.Ok)

    # 复制行
    @pyqtSlot()
    def on_actEdit_Copy_triggered(self):
        # 返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            self.ui.statusBar.showMessage("没有表格被选中，无法增、删、插、复制行", 5000)
            return
        print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        self.templist = []
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        rows = len(selectModel.selectedRows())  # 选中了多少行
        strmess = [r+1 for r in rowsIndexList]  # 每个索引+1
        strmess.sort(reverse=False)
        self.ui.statusBar.showMessage(f"复制成功了{strmess} 行", 5000)
        print("写入用于复制行的 临时变量 self.templist")
        for row in range(rows):  # 循环了选中的行对象
            for col in range(cols):
                if col == 0:
                    # print("第一列开始，新建一个空列表")
                    self.templist.append([])
                tempItem = model.item(strmess[row]-1, col)
                if (tablename == 1 and col == 7) or (tablename == 32 and col == 7):
                    # print("有 不计标志 的 单元格")
                    if tempItem.checkState() == Qt.Checked:
                        lineStr = "1"
                    else:
                        lineStr = "0"
                    self.templist[row].append(lineStr)
                else:
                    # print("普通单元格写入！templist")
                    self.templist[row].append(tempItem.text())  # rowslist[row]依次取出行号
        print(self.templist)
        # 粘贴行 菜单可用
        self.ui.actEdit_Paste.setEnabled(True)

    # 粘贴行
    @pyqtSlot()
    def on_actEdit_Paste_triggered(self):
        print("粘贴行")
        if not self.templist:  # 如果复制列表为空 则不执行
            self.ui.statusBar.showMessage("没有复制原数据！", 5000)
        print("临时变量为： ", self.templist)
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            self.ui.statusBar.showMessage("选择要粘贴的行、再点击‘粘贴行’", 5000)
            return
        print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        print("当前取回的表为： ", tablename)
        print("判断一共多少列")
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        colsdata = len(self.templist[0])
        if cols != colsdata:
            self.ui.statusBar.showMessage("复制源与粘贴处 不为同一张表’", 5000)
            return
        print("# 开始插入行")
        rows = len(self.templist)
        # 断开所有 链接
        self.disconnectAll()
        for row in range(rows):
            itemlist = []  # QStandardItem 对象列表
            for i in range(cols):
                if (tablename == 1 and i == 7) or (tablename == 32 and i == 7):
                    print("#  不计标志 复选框")
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
        self.ui.statusBar.showMessage("‘粘贴行’ 成功！", 5000)
        # 刷新总价 变色等
        if tablename == 1:
            # 面积表 数据有变化 处理所有行行数据 辅助
            self.itemModel1_itemChanged_All()
            # 刷新 建筑面积 统计
            self.census_area()
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel1, 7)
        elif tablename == 2:
            # 辅助 措施表 数据有变化 处理所有行行数据
            self.itemModel2_itemChanged_All()
            # 刷新一下 措施费统计
            self.census_measures()
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel2)
        elif tablename == 31:
            self.rowsBackground(model)  # 色标
            self.boldcol0(model)  # 粗体
            self.m3ChangeRed(model)  # 粗体
            # 刷新分部分项合价总和
            self.itemModel31_summary()  # 刷新分部分项合价总和
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel31)
        elif tablename == 32:
            # 辅助 明细表 数据有变化 处理所有行行数据
            self.itemModel32_itemChanged_All()
            # 刷新分部分项合价总和
            self.itemModel31_summary()  # 刷新分部分项合价总和
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel31)
        # print("# 释放 临时变量")
        self.templist = []
        # 粘贴行 菜单不可用
        self.ui.actEdit_Paste.setEnabled(False)
        # 恢复 链接
        self.connectAll()

    # 删除行
    @pyqtSlot()
    def on_actDelRow_triggered(self):
        # 返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            self.ui.statusBar.showMessage("没有表格被选中，无法增、删、插、复制行", 5000)
            return
        # 取出返回值
        tablename, tableobj, model, selectModel, rowsIndexList = result
        # print("共计多少行 ", model.rowCount())
        # print("选中了行索引列表 排序后 ", rowsIndexList)
        # 如果只有最后一行也不可删除
        if model.rowCount() <= 1:
            self.ui.statusBar.showMessage("就剩下一行了，你忍心？", 5000)
            return
        elif model.rowCount() == len(selectModel.selectedRows()):
            self.ui.statusBar.showMessage("你居然想全部删光光，太残暴了！", 5000)
            return
        # 删除前 警告确认
        res = QMessageBox.warning(self, "警告", "请确认是否需要删除选中的行！", QMessageBox.Yes | QMessageBox.No)
        if QMessageBox.Yes == res:
            print("点了确认！")
        elif QMessageBox.No == res:
            print("点了否！")
            return
        # # 断开所有 链接
        self.disconnectAll()
        # 行删除执行
        for row in rowsIndexList:
            model.removeRow(row)
        # 刷新总价 变色等
        if tablename == 1:
            # 面积表 数据有变化 处理所有行行数据 辅助
            self.itemModel1_itemChanged_All()
            # 刷新 建筑面积 统计
            self.census_area()
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel1, 7)
        elif tablename == 2:
            # 辅助 措施表 数据有变化 处理所有行行数据
            self.itemModel2_itemChanged_All()
            # 刷新一下 措施费统计
            self.census_measures()
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel2)
        elif tablename == 31:
            self.rowsBackground(model)  # 色标
            self.boldcol0(model)  # 粗体
            self.m3ChangeRed(model)  # 粗体
            # 刷新分部分项合价总和
            self.itemModel31_summary()  # 刷新分部分项合价总和
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel31)
            self.ui.tableView_32.setEnabled(False)  # 明细表不可用
        elif tablename == 32:
            # 辅助 明细表 数据有变化 处理所有行行数据
            self.itemModel32_itemChanged_All()
            # 刷新分部分项合价总和
            self.itemModel31_summary()  # 刷新分部分项合价总和
            # 各类表格写入 房号表
            self.mod_to_buildNumTable(self.itemModel31)
        # # 恢复 链接
        self.connectAll()

    # 插入行 一行
    @pyqtSlot()
    def on_actIntRow_triggered(self):
        # print("返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制")
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            self.ui.statusBar.showMessage("没有表格被选中，无法增、删、插、复制行", 5000)
            return
        # print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        print("当前取回的表为： ", tablename)
        # print("判断一共多少列")
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        # 开始插入行
        rows = len(selectModel.selectedRows())
        # # 断开所有 链接
        # self.disconnectAll()
        itemlist = []  # QStandardItem 对象列表
        for i in range(cols):
            item = QStandardItem("")
            #  不计标志 复选框
            if (tablename == 1 and i == 7) or (tablename == 32 and i == 7):
                item.setFlags(self.__NoCalFlags)
                item.setCheckable(True)  # 非锁定
                item.setCheckState(Qt.Unchecked)  # 非勾选
            itemlist.append(item)  # 一行空数据
        model.insertRow(rowsIndexList[-1], itemlist)  # 在最下面行的下面插入一行
        # 定位选中的单元格
        # curIndex = model.index(rowsIndexList[0], 0)
        # # thisselectionModel.clearSelection()  # 清除选择模型 点击2次 插入行 死机退出
        # selectModel.setCurrentIndex(curIndex, QItemSelectionModel.select)
        # # 恢复 链接
        # self.connectAll()

    # 增加行,在最后增加行
    @pyqtSlot()
    def on_actAppRow_triggered(self):
        # print("返回表名、模型、选择模型、降序后的选中行索引列表 用于行的 增 删 插 复制")
        result = self.viewCurrentModel()
        # 判断表有无焦点
        if not result:
            self.ui.statusBar.showMessage("没有表格被选中，无法增、删、插、复制行", 5000)
            return
        # print("取出返回值")
        tablename, tableobj, model, selectModel, rowsIndexList = result
        print("当前取回的表为： ", tablename)
        # print("判断一共多少列")
        cols = model.columnCount()  # 先要获得列数 一行多少单元格
        # 开始在最后增加一行
        # 断开所有 链接
        # self.disconnectAll()
        itemlist = []  # QStandardItem 对象列表
        for i in range(cols):
            item = QStandardItem("")
            #  不计标志 复选框
            if (tablename == 1 and i == 7) or (tablename == 32 and i == 7):
                item.setFlags(self.__NoCalFlags)
                item.setCheckable(True)  # 非锁定
                item.setCheckState(Qt.Unchecked)  # 非勾选
            itemlist.append(item)  # 一行空数据
        model.appendRow(itemlist)  # 在最下面行的下面插入一行
        # 恢复 链接
        # self.connectAll()

    # 工程量表 打开时，自动灰显，防止重复按下  辅助 工程量表 按钮
    def do_setActLocateEnable(self, enable):
        self.ui.tableView_0.setEnabled(enable)
        self.ui.pushButton_addBuild.setEnabled(enable)
        self.ui.pushButton_copyBuild.setEnabled(enable)
        self.ui.pushButton_delBuild.setEnabled(enable)
        self.ui.actGlod_quan.setEnabled(enable)
        # self.ui.actEdit_Copy.setEnabled(enable)
        # self.ui.actEdit_Paste.setEnabled(enable)
        # self.ui.actDelRow.setEnabled(enable)
        # self.ui.actIntRow.setEnabled(enable)
        # self.ui.actAppRow.setEnabled(enable)

    # 工程量表 按钮被点击
    @pyqtSlot()
    def on_actGlod_quan_triggered(self):
        # if self.dlgQuanObj == None:  # 未创建对话框
        #     self.dlgQuanObj = QmyDialogQuantities(self)
        self.dlgQuanObj = QmyDialogQuantities(self)
        # 把公共变量（字典）传给 工程量表
        buildnameRow = self.selectionModel0.currentIndex().row()  # 房号表 当前行
        buildname = self.itemModel0.item(buildnameRow, 0).text()  # 房号表 第1单元格内容
        quantitiesDict = {}
        item = self.itemModel0.item(buildnameRow, 5).text()
        if item:
            item = ast.literal_eval(item)
            quantitiesDict["历史操作"] = item

        item = self.itemModel0.item(buildnameRow, 6).text()
        if item:
            item = ast.literal_eval(item)
            quantitiesDict["工程量总表"] = item

        item = self.itemModel0.item(buildnameRow, 7).text()
        if item:
            item = ast.literal_eval(item)
            quantitiesDict["分类表"] = item

        item = self.itemModel0.item(buildnameRow, 8).text()
        if item:
            item = ast.literal_eval(item)
            quantitiesDict["钢筋工程量文件表"] = item

        item = self.itemModel0.item(buildnameRow, 9).text()
        if item:
            item = ast.literal_eval(item)
            quantitiesDict["钢筋工程量表"] = item

        item = self.itemModel0.item(buildnameRow, 10).text()
        if item:
            item = ast.literal_eval(item)
            quantitiesDict["绘图工程量文件表"] = item

        item = self.itemModel0.item(buildnameRow, 11).text()
        if item:
            item = ast.literal_eval(item)
            quantitiesDict["绘图工程量表"] = item

        if len(quantitiesDict):
            self.dlgQuanObj.setDictToModel(quantitiesDict)

        if buildname:
            self.dlgQuanObj.setBuildName(buildname)
        # if self.quantitiesDict:
        #     self.dlgQuanObj.setDictToModel(self.quantitiesDict)
        self.dlgQuanObj.changeActionEnable.connect(self.do_setActLocateEnable)  # 防止主界面菜单重复点击
        # 工程量表 确定关闭 或者 “刷新数据”被点击
        self.dlgQuanObj.quantitiesTableLists.connect(self.do_writeListToBuildTable)

        # ret = self.dlgQuanObj.exec()  # 以模态方式运行对话框
        ret = self.dlgQuanObj.show()  # 以非模态方式运行对话框
        # if ret == QDialog.Accepted:
        #     print("工程量表 按下了 OK")
        #     self.quantitiesDict = self.dlgQuanObj.emitAllTableToMain()  # 取回工程量总表的各个表字典形式
        #     self.do_writeListToBuildTable(self.quantitiesDict)
            # self.do_setActLocateEnable(True)  # 控件打开
        # else:
        #     pass
            # self.do_setActLocateEnable(True)  # 控件打开

    # 工程量表 确定关闭 或者 “刷新数据”被点击执行
    def do_writeListToBuildTable(self, dictdata):
        buildrow = self.selectionModel0.currentIndex().row()
        print(f"工程量表的 刷新主程序被按下, 房号表 在 ：{buildrow} 行")
        self.quantitiesDict = dictdata
        # 断开所有链接
        self.disconnectAll()
        # 清单表内几张表 写入房号表
        tableView_0 = self.quantitiesDict.get("历史操作")  # 历史操作表
        if tableView_0:
            item0 = str(tableView_0)
            item0 = QStandardItem(item0)
            self.itemModel0.setItem(buildrow, 5, item0)
        else:
            self.itemModel0.setItem(buildrow, 5, QStandardItem(''))

        tableView_11 = self.quantitiesDict.get("工程量总表")  # 工程量总表
        if tableView_11:
            item11 = str(tableView_11)
            item11 = QStandardItem(item11)
            self.itemModel0.setItem(buildrow, 6, item11)

        tableView_2 = self.quantitiesDict.get("分类表")  # 分类表
        if tableView_2:
            item2 = str(tableView_2)
            item2 = QStandardItem(item2)
            self.itemModel0.setItem(buildrow, 7, item2)

        tableView_3 = self.quantitiesDict.get("钢筋工程量文件表")  # 钢筋工程量文件表
        if tableView_3:
            item3 = str(tableView_3)
            item3 = QStandardItem(item3)
            self.itemModel0.setItem(buildrow, 8, item3)
        else:
            item3 = QStandardItem('')
            self.itemModel0.setItem(buildrow, 8, item3)

        tableView_3 = self.quantitiesDict.get("钢筋工程量表")  # 钢筋工程量表
        if tableView_3:
            item3 = str(tableView_3)
            item3 = QStandardItem(item3)
            self.itemModel0.setItem(buildrow, 9, item3)

        tableView_41 = self.quantitiesDict.get("绘图工程量文件表")  # 绘图工程量文件表
        if tableView_41:
            item41 = str(tableView_41)
            item41 = QStandardItem(item41)
            self.itemModel0.setItem(buildrow, 10, item41)
        else:  # 可能有空表
            item41 = QStandardItem('')
            self.itemModel0.setItem(buildrow, 10, item41)

        tableView_42 = self.quantitiesDict.get("绘图工程量表")  # 绘图工程量表
        if tableView_42:
            item42 = str(tableView_42)
            item42 = QStandardItem(item42)
            self.itemModel0.setItem(buildrow, 11, item42)

        # 辅助 面积表 数据有变化 处理所有行行数据
        self.itemModel1_itemChanged_All()
        # 刷新 建筑面积 统计
        self.census_area()
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel1, 7)
        # 辅助 措施表 数据有变化 处理所有行行数据
        self.itemModel2_itemChanged_All()
        # 刷新一下 措施费统计
        self.census_measures()
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel2)
        # 刷新 清单表内 明细表 的工程量引用
        # self.ui.tableView_32.clearSelection()
        # self.ui.tableView_32.setEnabled(False)  # 明细表可用
        print("刷新 清单表内 明细表")
        self.quantiti_TotalPrice()
        self.itemModel32.setRowCount(0)
        # 刷新分部分项合价总和
        self.itemModel31_summary()
        # self.itemModel0.itemChanged.connect(self.itemModel0_itemChanged)
        # 各类表格写入 房号表
        self.mod_to_buildNumTable(self.itemModel31)
        # 恢复所有链接
        self.connectAll()

    # 开始分析计算  为分析系统的数据可视化提供后台数据
    @pyqtSlot()
    def on_actDataPrepare_triggered(self):
        if self.filename == None:
            QMessageBox.information(self, "提示", "需要先保存文件，才可以分析数据", QMessageBox.Yes)
            return
        else:  # 生成可视化路径文件夹
            res = QMessageBox.information(self, "提示", "分析计算前自动保存工程文件，是否开始？", QMessageBox.Yes | QMessageBox.No)
            if QMessageBox.Yes == res:
                print("点了确认！")
                self.on_actSave_triggered()  # 先保存一次
                DataCal(self.filename)  # 所有html生成程序都在 dataVisualization文件中
            elif QMessageBox.No == res:
                print("点了否！")
                return
        self.ui.tabWidget_1.setCurrentIndex(0)
        self.ui.tabWidget_1.setCurrentIndex(1)

    # 分部分项表工程量设0(保留计算痕迹)
    @pyqtSlot()
    def on_actPartialanditemizedzeroing_triggered(self):
        print("分部分项表工程量设0(保留计算痕迹)")
        # 断开连接
        self.selectionModel31.currentChanged.disconnect(self.selectionModel31_currentChanged)  #
        self.itemModel31.itemChanged.disconnect(self.itemModel31_itemChanged)  # 主表有变化时触发
        rows = self.itemModel31.rowCount()
        for row in range(rows):
            self.itemModel31.setItem(row, 7, QStandardItem(""))
        # 辅助 分部分项表 每一行 量 价 合
        self.itemModel31_summary()
        # 恢复连接
        self.selectionModel31.currentChanged.connect(self.selectionModel31_currentChanged)  #
        self.itemModel31.itemChanged.connect(self.itemModel31_itemChanged)  # 主表有变化时触发

    # 导入房号
    @pyqtSlot()
    def on_actImportBuild_triggered(self):
        print("导入房号")
        if self.filename:
            curPath = self.filename
        else:
            curPath = QDir.currentPath()  # 获取系统当前目录
        dlgTitle = "导入诚清单房号"  # 对话框标题
        filt = "工程文件(*.cqd);;所有文件(*.*)"  # 文件过滤器
        # filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        if filename:
            # self.filename = filename
            # 获取 JSON 数据
            data = self.open_JSON_getdata(filename)
            # data 数据写入模型
            datalist2 = data.get('房号信息')  # 二维list
            # 断开连接
            self.disconnectAll()
            # 初始化一些数据
            # 写入数据
            rows = len(datalist2)
            cols = len(datalist2[0])
            rowsbuild = self.itemModel0.rowCount()
            for row in range(rows):
                list1 = []
                for col in range(cols):
                    item = QStandardItem(datalist2[row][col])
                    list1.append(item)
                self.itemModel0.insertRow(rowsbuild + row, list1)
            # 刷新单体信息的多个控件
            # singleinforstr = datalist2[0][1]
            # if singleinforstr:
            #     sDict = ast.literal_eval(singleinforstr)  # str还原成Dict
            #     self.singleInforSetControl(sDict)
            self.sameBuildName()  # 房号命重复时红色  辅助
            # 恢复连接
            self.connectAll()
            # 刷新用
            # self.selectionModel0_currentChanged(self.itemModel0.index(0, 0))

    # 1;记录逻辑
    @pyqtSlot()
    def on_actAutomaticRecord_triggered(self):
        pass

    # 2;智能匹配
    @pyqtSlot()
    def on_actAutomaticMatching_triggered(self):
        pass

    # 撤销
    @pyqtSlot()
    def on_act_Undo_triggered(self):
        print("undo 被点击了")
        self.ui.act_Redo.setEnabled(True)
        pass

    # 重做
    @pyqtSlot()
    def on_act_Redo_triggered(self):
        pass

    # TODO  ============窗体加载程序 ================================
    # 开机画面辅助 加载
    def load_data(self, sp):
        applist = ['[清单系统]', '[工程量清单系统]', '[数据可视化]', ]
        for i in range(1, 3):  # 模拟主程序加载过程 加载时间 秒
            time.sleep(1)  # 加载数据 间隔一秒
            sp.showMessage("正在加载程序{1}... {0}%".format(i * 10, applist[i]), Qt.AlignHCenter | Qt.AlignBottom, Qt.white)  # white  cyan  black
            # 设置字体
            sp.setFont(QFont('微软雅黑', 10))
            qApp.processEvents()  # 允许主进程处理事件

    # 开机画面辅助 退出菜单响应
    # def fun_Exit(self):
    #     response_quit = QApplication.instance()
    #     response_quit.quit()

# TODO  ============本窗体测试程序 ================================
if __name__ == "__main__":  # 用于当前窗体测试
    app = QApplication(sys.argv)  # 创建GUI应用程序
    form = QmyMainWindow()  # 创建窗体
    form.show()
    sys.exit(app.exec_())