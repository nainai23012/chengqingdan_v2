# -*- coding: utf-8 -*-

##  GUI应用程序主程序入口

import sys

from PyQt5.QtWidgets import QApplication, QSplashScreen
from PyQt5.QtGui import QIcon, QPixmap, QFont
from PyQt5.QtCore import pyqtSlot, pyqtSignal, Qt
from myMainWindow import QmyMainWindow


# 重写QSplashScreen类 用于开机界面
class MySplashScreen(QSplashScreen):
    # 鼠标点击事件
    def mousePressEvent(self, event):
        pass
        # window.show()  # 显示主窗口
        # splash.finish(window)  # 隐藏启动界面


# TODO exe打包命令 切换到程序目录下  pyinstaller -F -w -i C1.ico appMain.py
app = QApplication(sys.argv)
# 设置LOGO
icon = QIcon(":/icons/images/bitbug_favicon (8).ico")
app.setWindowIcon(icon)
# 设置启动界面
splash = MySplashScreen()
# 初始图片
splash.setPixmap(QPixmap(r':/icons/images/start_frame.png'))  # 设置背景图片
# splash=QSplashScreen(QPixmap(r'D:\vippython2\xlwings 模块\工程量清单助手 test-20\QtApp\images\start_frame.png'))  # 设置背景图片
# 初始文本
splash.showMessage("加载... 0%", Qt.AlignHCenter | Qt.AlignBottom, Qt.white)  # white  cyan  black
# # 设置字体
splash.setFont(QFont('微软雅黑', 10))
# 显示启动界面
splash.show()
app.processEvents()  # 处理主进程事件
# 主窗口
window = QmyMainWindow()
window.load_data(splash)  # 加载数据
window.show()
splash.finish(window)  # 隐藏启动界面
splash.deleteLater()
app.exec_()

# app = QApplication(sys.argv) 创建GUI应用程序
# icon = QIcon(":/icons/images/bitbug_favicon (8).ico")
# app.setWindowIcon(icon)
# mainform=QmyMainWindow()        #创建主窗体
# mainform.show()                 #显示主窗体
# sys.exit(app.exec_())
