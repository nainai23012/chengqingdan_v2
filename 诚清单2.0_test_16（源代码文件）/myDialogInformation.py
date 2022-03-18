import sys

from PyQt5.QtWidgets import QApplication, QDialog, QAbstractItemView, QInputDialog, QLineEdit

##from PyQt5.QtWidgets import  QDialog

from PyQt5.QtCore import pyqtSlot

##from PyQt5.QtCore import  pyqtSlot,Qt

##from PyQt5.QtWidgets import  

##from PyQt5.QtGui import

from ui_QWDialoginformation import Ui_Dialog
from pickle import dump, load


class QmyDialoginformation(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)  # 调用父类构造函数，创建窗体
        self.ui = Ui_Dialog()  # 创建UI对象
        self.ui.setupUi(self)  # 构造UI界面
        print('对象创建了')
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
    def setDict(self, infordict):  # 从主窗口得到 字典数据
        print(infordict)
        self.ui.lineEdit_1.setText(infordict['项目名称'])
        self.ui.lineEdit_2.setText(infordict['所属事业部'])
        self.ui.lineEdit_3.setText(infordict['项目所在省份'])
        self.ui.lineEdit_4.setText(infordict['项目所在市'])
        self.ui.lineEdit_5.setText(infordict['地区类型'])
        self.ui.lineEdit_6.setText(infordict['设计院名称'])
        self.ui.lineEdit_7.setText(infordict['人防设计院名称'])
        # self.ui.textEdit.setPlainText(infordict['记事本'])

    def getDict(self):
        self.infordict = {}
        str1 = self.ui.lineEdit_1.text()
        str2 = self.ui.lineEdit_2.text()
        str3 = self.ui.lineEdit_3.text()
        str4 = self.ui.lineEdit_4.text()
        str5 = self.ui.lineEdit_5.text()
        str6 = self.ui.lineEdit_6.text()
        str7 = self.ui.lineEdit_7.text()
        # str8 = self.ui.textEdit.toPlainText()
        print('读取完毕')
        self.infordict['项目名称'] = str1
        self.infordict['所属事业部'] = str2
        self.infordict['项目所在省份'] = str3
        self.infordict['项目所在市'] = str4
        self.infordict['地区类型'] = str5
        self.infordict['设计院名称'] = str6
        self.infordict['人防设计院名称'] = str7
        # self.infordict['记事本'] = str8
        return self.infordict

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
    @pyqtSlot(bool)  # 立刻登录
    def on_btnLogin_clicked(self):
        print('立刻登录')

    @pyqtSlot(bool)  # 新用户注册
    def on_btnRegister_clicked(self):
        print('新用户注册')
        res = QInputDialog.getText(self, "Get text", "Your name:", QLineEdit.Normal, "www")  # 参数以此为 标题 正文 占位
        item, okPressed = QInputDialog.getItem(self, "Get item", "Color:", items, 0, False)

    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_btnOK_clicked(self):  # 确定 按钮
        print('信息表里的确定 按钮')

    @pyqtSlot(bool)  # 修饰符指定参数类型，用于overload型的信号
    def on_btnSave_clicked(self):  # 保存 按钮
        pass
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
    form = QmyDialoginformation()  # 创建窗体
    form.show()
    sys.exit(app.exec_())
