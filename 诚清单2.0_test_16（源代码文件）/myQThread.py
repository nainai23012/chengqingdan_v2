# -*- coding: utf-8 -*-
# 保存工程文件时，发送信息给mysql 不造成卡顿
import sys
import os
import time
import json
import socket
import re  # 导入正则表达式模块
import ast  # 列表包裹在引号内，提取出来变成列表  字符转列表 字符转字典等
import getpass  # 获取当前用户名
import platform  # 获取操作系统版本相关
import wmi  # 获取硬盘 cpu 主板 mac bios等硬件信息模块
import pymysql  # mysql服务器
# 导入自定义模块
import some_infor
# 导入PYQT5 相关模块
from PyQt5.QtWidgets import qApp,QApplication, QMainWindow, QUndoStack, QUndoCommand, \
    QMessageBox, QSpinBox, QLabel, QTableView, QCheckBox, QAbstractItemView, QHeaderView, \
    QColorDialog, QDialog, QInputDialog, QLineEdit
from PyQt5.QtCore import pyqtSlot, pyqtSignal, Qt, QItemSelectionModel, QStringListModel, \
    QTimer, QTime, QDateTime, QThread
from PyQt5.QtGui import QFont, QColor, QPalette, QStandardItemModel, QStandardItem

#          # 用于多线程 用户信息
#         self.thread = Worker()
#         self.thread.advertise.connect(self.advertise_info)
#         self.thread.start()



class Eng_doc(QThread):
    # advertise = pyqtSignal(list)  # 打开程序时 从服务器接收数据 元组形式 广而告之；advertise

    def __init__(self, parent=None):
        super(Eng_doc, self).__init__(parent)
        self.working = True
        # self.num = 0

    def __del__(self):
        self.working = False
        self.wait()

    def run(self):
        while self.working == True:
            print("开始执行开机任务支线！")
            diskstr = some_infor.get_diskNum()  # 硬盘信息的查询  返回硬盘号
            print('硬盘号:', diskstr)
            self.user_infor_send(diskstr)  # 电脑硬件信息发送
            self.user_login_send(diskstr)  # 登录信息发送
            some_infor.advertise_infor_receive()  # 更新广告
            self.advertise_emit_main()  # 本地取广告 并定时发送给主程序 TODO 无限循环 放在最后
            self.working = False


    # TODO  ============文件处理 相关 ================================保存 打开 新建 另存 备份等
    # # 获取文件大小
    # def get_file_size(self, filepath):
    #     """
    #     获取文件大小，结果保留两位小数，单位MB
    #     """
    #     f = os.path.getsize(filepath)
    #     f = f / float(1024 * 1024)
    #     return round(f, 2)
    #
    #
    # # 获取文件创建时间
    # def get_file_create_time(self, filepath):
    #     """
    #     获取文件创建时间
    #     """
    #     print("获取文件创建时间")
    #     tf = os.path.getctime(filepath)
    #     t = time.localtime(tf)
    #     # 时间戳转换方法
    #     return time.strftime('%Y-%m-%d %H:%M:%S', t)
    #
    # # 获取文件最新修改时间
    # def get_file_modify_time(self, filepath):
    #     """
    #     获取文件修改时间
    #     """
    #     tf = os.path.getmtime(filepath)
    #     t = time.localtime(tf)
    #     # 时间戳转换方法
    #     return time.strftime('%Y-%m-%d %H:%M:%S', t)
    #
    # # 获取文件访问时间
    # def get_file_visit_time(self, filepath):
    #     """
    #     获取文件访问时间
    #     """
    #     tf = os.path.getatime(filepath)
    #     t = time.localtime(tf)
    #     # 时间戳转换方法
    #     return time.strftime('%Y-%m-%d %H:%M:%S', t)


    # TODO  ============MySQL 数据库相关 ================================ 连接
    # def user_infor_send(self, diskstr):  # 电脑硬件信息发送
    #     print("电脑硬件信息发送")
    #     if not diskstr:
    #         return
    #     try:
    #         diskstr = str(diskstr)
    #         conn = pymysql.connect(host='rm-bp114m07t2e13i30i9o.mysql.rds.aliyuncs.com',
    #                                port=3306, user='use_cqd000', password='Cqd123456', db='chengqingdan2021',
    #                                charset='utf8')
    #         # status = conn.server_status
    #         cursor = conn.cursor()
    #         sql = '''
    #             select * from user_infor where disknum=%s;
    #         '''
    #         effect_row = cursor.execute(sql, [diskstr])
    #         print("查到了数据 ", effect_row, "条！")
    #         cuu = cursor.fetchall()
    #         if effect_row:  # 如果查的到数据 说明有
    #             pass
    #         else:  # 写入数据
    #             sql = '''
    #                 INSERT INTO user_infor(disknum) VALUES (%s);
    #             '''
    #             row = cursor.execute(sql, [diskstr])
    #             conn.commit()  # 提交
    #             print("写入", row)
    #         cursor.close()  # 4.关闭游标
    #         conn.close()  # 5.关闭连接
    #     except Exception as e:
    #         print("user_infor_send 连接失败！", e)

    def user_login_send(self, diskstr):  # 登录信息发送
        print("登录信息发送")
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
            print("查到了数据 ", effect_row, "条！")
            cuu = cursor.fetchall()

            conn.commit()  # 提交
            cursor.close()  # 4.关闭游标
            conn.close()  # 5.关闭连接
        except Exception as e:
            print("user_login_send 连接失败！", e)

    # # 本地取广告 并发送给主程序
    # def advertise_emit_main(self):
    #     print('# 本地取广告 并发送给主程序')
    #     try:
    #         with open(r".\data\main_data\advertiseDict.db", "r", encoding='utf-8') as f:
    #             advertiseDict = json.load(f)  # 字典 列表
    #         if advertiseDict:
    #             cuu = advertiseDict['advertise']  # 取出元组
    #         if cuu:
    #             rows = len(cuu)   # 一共有多少条数据
    #             while True:  # 无限循环
    #                 for row in range(rows):
    #                     linetuple = cuu[row]
    #                     continuedtime = linetuple[1]
    #                     self.advertise.emit(linetuple)
    #                     time.sleep(continuedtime)
    #     except Exception as e:
    #         print("取广告失败！", e)