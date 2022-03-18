import sys

from PyQt5.QtWidgets import QApplication, QDialog, QAbstractItemView, QInputDialog

##from PyQt5.QtWidgets import  QDialog

from PyQt5.QtCore import pyqtSlot, Qt

##from PyQt5.QtCore import  pyqtSlot,Qt

##from PyQt5.QtWidgets import  

##from PyQt5.QtGui import

from ui_QWDialogRegister import Ui_Dialog
from pickle import dump, load


class QmyDialogRegister(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)  # 调用父类构造函数，创建窗体
        self.ui = Ui_Dialog()  # 创建UI对象
        self.ui.setupUi(self)  # 构造UI界面
        print('对象创建了')

        self.setFixedSize(self.width(), self.height())
        self.setWindowFlag(Qt.WindowStaysOnTopHint)  # StayOnTop显示 保持前端
        # self.loadpickle() # 取泡菜

        # self.__model=QStringListModel()
        # self.ui.listView.setModel(self.__model)
        #
        # self.ui.listView.setAlternatingRowColors(True)
        # self.ui.listView.setDragDropMode(QAbstractItemView.InternalMove)
        # self.ui.listView.setDefaultDropAction(Qt.MoveAction)

    def __del__(self):  # 析构函数
        print("QmyDialoginformation 对象被删除了")

    ##  ==============自定义功能函数============
    def setDict(self):  # 从主窗口得到 字典数据
        print('setDict')

    def getDict(self):
        print('getDict')

    # def savepickle(self, data):  # 存泡菜
    #     pickle_file = open(r'data\工程信息.pkl', 'wb')
    #
    #     dump(data, pickle_file)  # 将列表倒入文件
    #     pickle_file.close()  # 关闭pickle文件
    #
    # def loadpickle(self):  # 取泡菜
    #     pickle_file = open(r'data\工程信息.pkl', 'rb')
    #     self.infordict = load(pickle_file)  # 取得匹配表数据 self.matchingList 字典形式｛默认表：[[],[],[]]｝
    #     pickle_file.close()

    # def setHeaderList(self,headerStrList):
    #    self.__model.setStringList(headerStrList)

    # def headerList(self):
    #    return self.__model.stringList()

    # TODO  ==========由connectSlotsByName() 自动连接的槽函数==================
    @pyqtSlot(bool)  # 立刻注册
    def on_pushButton_clicked(self):
        print('立刻注册')

    # @pyqtSlot(bool)  # 新用户注册
    # def on_pushButton_2_clicked(self):
    #     print('新用户注册')

        # print('保存 按钮')
        # self.infordict = {}
        # print('读取开始')
        # str1 = self.ui.lineEdit_1.text()
        # str2 = self.ui.lineEdit_2.text()
        # str3 = self.ui.lineEdit_3.text()
        # str4 = self.ui.lineEdit_4.text()
        # str5 = self.ui.textEdit.toPlainText()
        # # print('读取完毕')
        # self.infordict['工程名称']=str1
        # self.infordict['编制单位']=str2
        # self.infordict['编制人']=str3
        # self.infordict['创建时间']=str4
        # self.infordict['记事本']=str5


##  =============自定义槽函数===============================


##  ============窗体测试程序 ================================
if __name__ == "__main__":  # 用于当前窗体测试
    app = QApplication(sys.argv)  # 创建GUI应用程序
    form = QmyDialogRegister()  # 创建窗体
    form.show()
    sys.exit(app.exec_())
