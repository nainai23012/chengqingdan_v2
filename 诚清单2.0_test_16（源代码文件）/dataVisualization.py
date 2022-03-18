# -*- coding: utf-8 -*-
# 为数据可视化模块 生成html文件 需要主程序提供 filepath 保存过后的文件路径
import sys
import os
import time
import json
import random
# import pickle  # 泡菜模块
# import socket
import re  # 导入正则表达式模块
import ast  # 列表包裹在引号内，提取出来变成列表  字符转列表 字符转字典等
# import getpass  # 获取当前用户名
# import platform  # 获取操作系统版本相关
# import wmi  # 获取硬盘 cpu 主板 mac bios等硬件信息模块
# import pymysql  # mysql服务器
# import csv
import xlrd, xlwt
# import xlwings as xw
# from pprint import pprint
# 数据可视化分析的图表模块
from pyecharts import options as opts
from pyecharts.options import ComponentTitleOpts
from pyecharts.charts import Bar, Bar3D, Grid, Line, Page, Pie, Tab
from pyecharts.faker import Faker
from pyecharts.components import Table
from pyecharts.globals import ThemeType  # 样式
# 导入自定义模块
import some_infor

class DataCal:
    # 初始化方法，初始化了成员变量
    def __init__(self, filepath):
        self.filepath = filepath  # 工程文件路径
        self.bmtsDatas = None  # 主材汇总表的二维表
        self.bmtsDatas2 = None  # 主材汇总表的二维表     占地面积 指标用
        self.areaDict = {}  # 面积表 {1#:{总面积:0,地上面积:0,地下面积:0,占地面积:0,},}
        self.measuresPriceDict = {}  # 措施费单价  {1#:{整体措施费：0，模板费：0，其他项目清单：0}}
        self.subItemPriceDict = {}  # 分部分项单价  {1#:{名称+特征+地上地下+单位：0，...}}
        self.subItemQuantitiesDict = {}  # 分部分项工程量  {1#:{名称+特征+地上地下+单位：0，...}}
        self.typeWeightDict = {}  # 类别直径重量字典{1#:{梁_地上:0,板_地下:0},}
        self.levelDWeightDict = {}  # 级别直径重量字典{1#:{A6_地上:0,C8_地下:0},}
        self.levelWeightDict = {}  # 级别重量字典{1#:{A_地上:0,C_地下:0},}
        self.weightDict = {}  # 重量字典{1#:{A_地上:0,C_地下:0},}
        self.oneTwoWeightDict = {}  # 一次二次结构重量字典{1#:{地上_一次:0,地上_二次:0,地上_网片:0},}

        self.connectDict = {}  # 接头数量字典{1#:{电渣焊_地上:0,机械连接_地下:0},}
        self.connectDDict = {}  # 接头直径数量字典{1#:{电渣焊10_地上:0,机械连接20_地下:0},}
        self.eTypeConnectDDict = {}  # 电渣焊 构件直径
        self.mTypeConnectDDict = {}  # # 机械连接 构件直径

        self.levelDWeightDatas = None  # 级别直径重量二维表
        self.levelWeightDatas = None  # 级别重量二维表
        self.typeWeightDatas = None  # 类别重量二维表
        self.weightDatas = None  # 重量二维表
        self.oneTwoWeightDatas = None  # 一次二次结构重量二维表

        self.quo_sum_datas = None  # 报价汇总表二维表 ['房号', '垂直运输', '脚手架工程', '模板费用', '总面积', '整体措施费', '分部分项', '其他项目清单', '含税合价', '单方造价']
        self.connectDatas = None  # 接头数量二维表
        self.connectDDatas = None  # 接头直径数量二维表
        self.eTypeConnectDDatas = None  # 电渣焊 构件直径 接头数量二维表
        self.mTypeConnectDDatas = None  # 机械连接 构件直径 接头数量二维表
        # print(self.filepath)
        # 生成html文件的文件夹路径
        self.htmlfolder = os.path.splitext(self.filepath)[0] + '_data'
        if not os.path.exists(self.htmlfolder):
            os.mkdir(self.htmlfolder)
        # 读取数据
        self.engineeringDict = self.loadDatas(self.filepath)  # 读取主要数据
        # 生成钢筋预结算单
        self.iron_bill()
        # print(self.engineeringDict['房号信息'][0][8])  # [8] 钢筋工程量文件表
        # 获取面积、报价汇总表 初始化  赋值给公共变量 分部分项
        self.quotation_summary()  # 最后赋值给公共变量 self.areaDict  self.quo_sum_datas self.measuresPriceDict
        # ['房号', '垂直运输', '脚手架工程', '模板费用', '总面积', '整体措施费', '分部分项', '其他项目清单', '含税合价', '单方造价']
        # 生成土建清单报价汇总表
        self.civilengineering_bill()
        # 总览表 overview
        self.main_html()
        # 清单表分析
        self.detailedListAnalysis()
        # 土建模型分析
        self.soilWoodModelAnalysis()
        # 钢筋模型分析
        self.ironModelAnalysis()


    def loadDatas(self, path):  # 读取数据
        with open(path, "r", encoding='utf-8') as f:
            data = json.load(f)
        return data

    # TODO ~~~~~~~~~~~~~~~~~~~~初始化获取面积、报价汇总表 初始化~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # 报价汇总表 初始化
    def quotation_summary(self):
        # self.quo_sum_datas = None
        # self.areaDict = {}
        # self.areaDict = {1#:{总面积：200，地上面积：50，地下面积：100，占地面积：10},2#....}
        buildcounts = len(self.engineeringDict['房号信息'])
        summaryDict = {}
        for buildnum in range(buildcounts):
            buildName = self.engineeringDict['房号信息'][buildnum][0]  # 每一个房号名字
            areadatas = self.engineeringDict['房号信息'][buildnum][2]  # 面积表
            measuresdatas = self.engineeringDict['房号信息'][buildnum][3]  # 措施表
            branchdatas = self.engineeringDict['房号信息'][buildnum][4]  # 分部分项表
            if buildName:
                summaryDict[buildName] = {}
                self.areaDict[buildName] = {}  # 面积字典
                self.measuresPriceDict[buildName] = {}  # 措施费单价字典
            # total_area = self.areaDict.get(buildName).get("总面积")  # 取出总面积
            # keystr = "总面积"
            # if total_area:
            #     summaryDict[buildName][keystr] = total_area
            # else:
            #     summaryDict[buildName][keystr] = 0
            # 建筑面积 处理
            if areadatas:
                areadatas = ast.literal_eval(areadatas)
                total_area = 0  # 总面积
                upper_area = 0  # 地上面积
                lower_area = 0  # 地下面积
                land_area = 0  # 占地面积，  面积最大的一层
                for row in range(len(areadatas)):  # 取出占地面积
                    resultstr = areadatas[row][5]  # 取出每一行的“计算结果”
                    if resultstr:
                        resultfloat = float(resultstr)
                        if resultfloat > land_area:
                            land_area = resultfloat

                for row in range(len(areadatas)):  # 地上、地下、总面积
                    numstr = areadatas[row][6]  # 取出每一行的面积
                    if numstr:
                        numfloat = float(numstr)
                        if areadatas[row][1] == "地下":
                            lower_area += numfloat
                        elif areadatas[row][1] == "地上":
                            upper_area += numfloat
                        total_area += numfloat
                if upper_area + lower_area != total_area and total_area != 0:  # 如果地下 地上 都空 总面积有，就把所有面积都放在地上
                    upper_area = total_area - lower_area
                # 修正小数位数
                upper_area = round(upper_area, 2)
                lower_area = round(lower_area, 2)
                total_area = round(total_area, 2)
                land_area = round(land_area, 2)
                # 写入面积字典
                self.areaDict[buildName]["总面积"] = total_area
                self.areaDict[buildName]["地上面积"] = upper_area
                self.areaDict[buildName]["地下面积"] = lower_area
                self.areaDict[buildName]["占地面积"] = land_area
            else:  # 如果没有面积 则写入 0
                # 写入面积字典
                self.areaDict[buildName]["总面积"] = 0
                self.areaDict[buildName]["地上面积"] = 0
                self.areaDict[buildName]["地下面积"] = 0
                self.areaDict[buildName]["占地面积"] = 0
            # 措施表 处理
            if measuresdatas:
                measureslists = ast.literal_eval(measuresdatas)  # 措施表
                # whole = 0  # 整体措施费
                # singlestrdict = {}  # 单项清单字符列表
                # single = 0  # 单项措施费
                # other = 0  # 其他项目清单
                for row in range(len(measureslists)):
                    str8 = measureslists[row][8]  # 单价 字符型
                    str9 = measureslists[row][9]  # 合价 字符型
                    if not str9:  # 没有合价就跳过
                        continue
                    float8 = round(float(str8), 3)  # 单价小数位处理
                    float9 = float(str9)
                    str1 = measureslists[row][1]  # 整体措施费、单项措施费、其他项目清单
                    str2 = measureslists[row][2]  # 模板、脚手架、垂直 动态获取
                    if str1 == "整体措施费":
                        keystr = "整体措施费"
                        result = summaryDict[buildName].get(keystr)
                        if result == None:
                            summaryDict[buildName][keystr] = float9
                        else:
                            summaryDict[buildName][keystr] += float9
                        # 措施费单价表
                        self.measuresPriceDict[buildName][keystr] = float8
                    elif str1 == "其他项目清单":
                        keystr = "其他项目清单"
                        result = summaryDict[buildName].get(keystr)
                        if result == None:
                            summaryDict[buildName][keystr] = float9
                        else:
                            summaryDict[buildName][keystr] += float9
                        # 措施费单价表
                        self.measuresPriceDict[buildName][str2] = float8
                    elif str1 == "单项措施费":
                        keystr = "单项措施费"
                        if str2:  # 模板、脚手架、垂直 动态获取
                            result = summaryDict[buildName].get(str2)
                            if result == None:
                                summaryDict[buildName][str2] = float9
                            else:
                                summaryDict[buildName][str2] += float9
                        # 措施费单价表
                        self.measuresPriceDict[buildName][str2] = float8
                        # else:  # 没有  就取 单项措施费 的合计
                        result = summaryDict[buildName].get(keystr)
                        if result == None:
                            summaryDict[buildName][keystr] = float9
                        else:
                            summaryDict[buildName][keystr] += float9
            else:
                summaryDict[buildName]["整体措施费"] = 0
                summaryDict[buildName]["其他项目清单"] = 0
                summaryDict[buildName]["单项措施费"] = 0
                # 措施费单价表
                self.measuresPriceDict[buildName]["整体措施费"] = 0
                self.measuresPriceDict[buildName]["其他项目清单"] = 0
                self.measuresPriceDict[buildName]["单项措施费"] = 0
            # 分部分项表处理
            if branchdatas:
                keystr = "分部分项"
                branchlists = ast.literal_eval(branchdatas)
                for row in range(len(branchlists)):
                    str10 = branchlists[row][10]  # 合价
                    try:
                        flo10 = float(str10)
                    except:
                        flo10 = 0
                    result = summaryDict[buildName].get(keystr)
                    if result == None:
                        summaryDict[buildName][keystr] = flo10
                    else:
                        summaryDict[buildName][keystr] += flo10
            else:
                summaryDict[buildName]["分部分项"] = 0
        print(summaryDict)
        # 获取表头
        templist = []
        temp = [v for k, v in summaryDict.items()]
        for te in temp:
            for k in te.keys():
                templist.append(k)
        headers = list(set(templist))    # 去重
        # headers.remove("单项措施费")  # 单项措施费 不放在土建清单汇总表内
        # headers.sort()
        headers.insert(0, "房号")  # 最前面插入 房号 字段
        header_end = ["含税合价", "总面积", "单方造价"]
        headers.extend(header_end)  # 在最后面 扩展列表
        # 创建一个包含表头的 二维表
        datas = []
        datas.append(headers)  # 添加首行表头
        for k, v in summaryDict.items():
            buildname = k
            datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
            datas[-1][0] = buildname  # 写入房号
            for ka, va in v.items():  #   # {整体措施费：100，总面积：100}
                # if ka == "单项措施费":
                #     continue
                col = datas[0].index(ka)  # 获得列号
                datas[-1][col] = round(float(va), 2)
        print(datas)
        # 处理 合价与 总面积 单方造价
        rows = len(datas)
        cols = len(datas[0])
        for row in range(1, rows):  # 第一行表头 跳过
            sumprice = 0  # 价格合计
            result = self.areaDict.get(datas[row][0])
            if result:
                sumarea = self.areaDict[datas[row][0]]["总面积"]  # 面积合计
                datas[row][-2] = sumarea
            else:
                sumarea = 0
            for col in range(1, cols - 3):  # 最后 合价 单方 跳过
                if datas[0][col] == "单项措施费":  # 单项措施费有多个子目组成 跳过
                    continue
                price = datas[row][col]
                if isinstance(price, float):
                    sumprice += price
            datas[row][-3] = round(sumprice, 2)  # 合价 回填
            if sumarea:  # 如果有面积
                datas[row][-1] = round(sumprice / sumarea, 0)  # 单方 回填
            else:
                datas[row][-1] = 0
                # print(datas)
        # self.quo_sum_datas = datas
        self.quo_sum_datas = datas  #

    # TODO ~~~~~~~~~~~~~~~~~~~~总览表~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # 测试饼图 参数为html 文件名
    def main_html(self):
        # print(f"内部传参 路径为：{self.filepath}")
        # 获取总面积
        if len(self.areaDict) < 1:
            print("无面积表")
            return
        if self.quo_sum_datas == None:
            print("无报价汇总表")
            return
        datas = self.quo_sum_datas  # 调用数据
        # 获取总面积
        buildname = [ke for ke in self.areaDict.keys()]
        area = [int(va["总面积"]) for va in self.areaDict.values()]
        c1 = (
            Pie()
                .add("",
                [list(z) for z in zip(buildname, area)],
                radius=["30%", "55%"],  # 内环 外环 大小
                center=["25%", "50%"],  # 位置
                rosetype="radius",
                label_opts=opts.LabelOpts(is_show=True),)
                .set_global_opts(title_opts=opts.TitleOpts(title="总面积占比 单元m2", subtitle="", pos_left="20%", pos_top="5%"))
                .set_series_opts(label_opts=opts.LabelOpts(font_size=20, formatter="{b}\n{c}m2"))  # 设置标签字体大小
        )
        # 获取总总价
        colNum = None
        for col in range(len(datas[0])):
            if datas[0][col] == "含税合价":
                colNum = col
                break
        if colNum == None:
            return
        buildnames = [r[0] for r in datas[1:]]  # 二维表的第一列
        total = [round(r[colNum]/10000, 2) for r in datas[1:]]  # 二维表的总价列
        c2 = (
            Pie()
                .add(
                "",
                [list(z) for z in zip(buildnames, total)],
                radius=["30%", "55%"],
                center=["75%", "50%"],
                rosetype="radius",
                label_opts=opts.LabelOpts(is_show=True),)
                .set_global_opts(title_opts=opts.TitleOpts(title="总造价占比，单位：万元", subtitle="", pos_left="70%", pos_top="5%"))
                .set_series_opts(label_opts=opts.LabelOpts(font_size=20, formatter="{b}\n{c}万元"))  # 设置标签字体大小
        )
        grid = (
            Grid(init_opts=opts.InitOpts(width="1400px", height="800px"))
                .add(c1, grid_opts=opts.GridOpts(pos_right="58%"), is_control_axis_index=True)
                .add(c2, grid_opts=opts.GridOpts(pos_left="58%"), is_control_axis_index=True)
                # .render("grid_overlap_multi_xy_axis.html")
        )
        # return grid
        grid.render((self.htmlfolder + '\\' + '总览表.html'))  # '_data'文件夹内保存网页

    # TODO ~~~~~~~~~~~~~~~~~~~~清单分析~~~~~~~~~~~~~~~~~~~~~~~~~~~
    def detailedListAnalysis(self):  # 清单分析 页签
        tab = Tab()
        tab.add(self.unitPriceAnalysis(), "单体单方指标")  # 柱图 总造价单方 分部分项 整理 单项 其他 的单方
        tab.add(self.measuresPrice(), "措施费单价")  # 二维表 措施费单价表
        tab.add(self.buildingMaterialTotalSummary(), "主材汇总表")  # 表
        tab.add(self.buildingMaterialTotalProportion(), "主材含量表")  # 表
        tab.add(self.buildingMaterialTotalProportion2(), "占地面积含量表")  # 表
        tab.add(self.subItemPrice(), "分部分项单价对比")  # 表
        tab.add(self.subItemQuantities(), "分部分项工程量含量对比")  # 表
        tab.render((self.htmlfolder + '\\' + '清单分析.html'))  # '_data'文件夹内保存网页

    # 单体单方指标
    def unitPriceAnalysis(self):
        # print(f"内部传参 路径为：{self.filepath}")
        # 获取单方造价  柱状图
        if not self.quo_sum_datas:
            return
        datas = self.quo_sum_datas  # 调用数据
        # 单方造价 列
        colNum = None
        for col in range(len(datas[0])):
            if datas[0][col] == "单方造价":
                colNum = col
                break
        buildnames = [r[0] for r in datas[1:]]  # 二维表的第一列 取出房号
        unitPrice = []
        if colNum == None:
            unitPrice = [0 for r in datas[1:]]  # 二维表的第一列 取出房号
        else:
            for r in range(1, len(datas)):
                temp = datas[r][colNum]
                if temp:
                    temp = round(temp, 0)
                else:
                    temp = 0
                unitPrice.append(temp)
        # unitPrice = [int(r[colNum]) for r in datas[1:]]  # 二维表的单方造价列
        # 面积 列
        colNum = None
        for col in range(len(datas[0])):
            if datas[0][col] == "总面积":
                colNum = col
                break
        if colNum == None:
            area = [0 for r in datas[1:]]  # 二维表的 总面积 列
        else:
            area = [r[colNum] for r in datas[1:]]  # 二维表的 总面积 列

        # 整体措施费单方 列
        colNum = None
        for col in range(len(datas[0])):
            if datas[0][col] == "整体措施费":
                colNum = col
                break
        wholeCost = []  # 价格/总面积
        if colNum == None:
            wholeCost = [0 for r in datas[1:]]  #
        else:
            subItem = [r[colNum] for r in datas[1:]]  # 二维表的分部分项列
            for r in range(len(subItem)):
                if subItem[r] and area[r]:
                    temp = round((subItem[r] / area[r]), 0)
                else:
                    temp = 0
                wholeCost.append(temp)

        # 其他项目单方 列
        colNum = None
        for col in range(len(datas[0])):
            if datas[0][col] == "其他项目清单":
                colNum = col
                break
        otherCost = []  # 价格/总面积
        if colNum == None:
            otherCost = [0 for r in datas[1:]]
        else:
            subItem = [r[colNum] for r in datas[1:]]  # 二维表的分部分项列
            for r in range(len(subItem)):
                if subItem[r] and area[r]:
                    temp = round((subItem[r] / area[r]), 0)
                else:
                    temp = 0
                otherCost.append(temp)

        # 单项措施费单方 列
        colNum = None
        for col in range(len(datas[0])):
            if datas[0][col] == "单项措施费":
                colNum = col
                break
        singleCost = []  # 价格/总面积
        if colNum == None:
            singleCost = [0 for r in datas[1:]]
        else:
            subItem = [r[colNum] for r in datas[1:]]  # 二维表的分部分项列
            for r in range(len(subItem)):
                if subItem[r] and area[r]:
                    temp = round((subItem[r] / area[r]), 0)
                else:
                    temp = 0
                singleCost.append(temp)

        # 模板费用单方 列
        colNum = None
        for col in range(len(datas[0])):
            if "模板费" in datas[0][col]:   # 关键字 是否在 表头字内 可能模板费 or 费用
                colNum = col
                break
        templateCost = []  # 价格/总面积
        if colNum == None:
            templateCost = [0 for r in datas[1:]]
        else:
            subItem = [r[colNum] for r in datas[1:]]  # 二维表的分部分项列
            for r in range(len(subItem)):
                if subItem[r] and area[r]:
                    temp = round((subItem[r] / area[r]), 0)
                else:
                    temp = 0
                templateCost.append(temp)

        # 垂直运输单方 列
        colNum = None
        for col in range(len(datas[0])):
            if "垂直" in datas[0][col]:   # 关键字 是否在 表头字内 可能模板费 or 费用
                colNum = col
                break
        verticalCost = []  # 价格/总面积
        if colNum == None:
            verticalCost = [0 for r in datas[1:]]
        else:
            subItem = [r[colNum] for r in datas[1:]]  # 二维表的分部分项列
            for r in range(len(subItem)):
                if subItem[r] and area[r]:
                    temp = round((subItem[r] / area[r]), 0)
                else:
                    temp = 0
                verticalCost.append(temp)

        # 脚手架工程单方 列
        colNum = None
        for col in range(len(datas[0])):
            if "脚手架" in datas[0][col]:   # 关键字 是否在 表头字内 可能模板费 or 费用
                colNum = col
                break
        scaffoldingCost = []  # 价格/总面积
        if colNum == None:
            scaffoldingCost = [0 for r in datas[1:]]
        else:
            subItem = [r[colNum] for r in datas[1:]]  # 二维表的分部分项列
            for r in range(len(subItem)):
                if subItem[r] and area[r]:
                    temp = round((subItem[r] / area[r]), 0)
                else:
                    temp = 0
                scaffoldingCost.append(temp)

        # 分部分项单方造价 列
        colNum = None
        for col in range(len(datas[0])):
            if datas[0][col] == "分部分项":
                colNum = col
                break
        subItemPrice = []  # 分部分项总价/总面积
        if colNum == None:
            subItemPrice = [0 for r in datas[1:]]
        else:
            subItem = [r[colNum] for r in datas[1:]]  # 二维表的分部分项列
            for r in range(len(subItem)):
                if subItem[r] and area[r]:
                    temp = round((subItem[r] / area[r]), 0)
                else:
                    temp = 0
                subItemPrice.append(temp)
        # 普通柱状图
        bar = (
            Bar(init_opts=opts.InitOpts(width="1400px", height="800px"))
                .add_xaxis(buildnames)
                .add_yaxis("单方造价", unitPrice)  # gap=0% 柱间距
                .add_yaxis("分部分项", subItemPrice)
                .add_yaxis("整体措施费", wholeCost, is_selected=False)
                .add_yaxis("其他项目清单", otherCost, is_selected=False)
                .add_yaxis("单项措施费", singleCost)
                .add_yaxis("模板费用", templateCost, is_selected=False)
                .add_yaxis("垂直运输", verticalCost, is_selected=False)
                .add_yaxis("脚手架工程", scaffoldingCost, is_selected=False)
                .reversal_axis()  # 翻转XY轴
                .set_global_opts(title_opts=opts.TitleOpts(title="单位：元/m2", subtitle="说明：为对应的金额与总建筑面积的比值！"),
                                 toolbox_opts=opts.ToolboxOpts(),)
                # position="right" 数据标签在右侧
                .set_series_opts(label_opts=opts.LabelOpts(is_show=True, position="right"), markpoint_opts=opts.MarkPointOpts(
                    data=[
                        opts.MarkPointItem(type_="max", name="最大值"),
                        opts.MarkPointItem(type_="min", name="最小值"),
                        # opts.MarkPointItem(type_="average", name="平均值"),
                    ]))
            # 或者直接使用字典参数
            # .set_global_opts(title_opts={"text": "主标题", "subtext": "副标题"})
        )
        return bar
        # bar.render((self.htmlfolder + '\\' + '清单分析.html'))

    # 措施费单价表
    def measuresPrice(self):
        # print('措施费单价表', self.measuresPriceDict)
        if len(self.measuresPriceDict) < 1:  # 如果措施费字典为空 返回一个空的饼图
            table = (Table().add([], []).set_global_opts(
                title_opts=ComponentTitleOpts(title="措施项目清单没有数据", subtitle="无法统计")))
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in self.measuresPriceDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))    # 去重
            # headers.sort()  # 排序
            headers.insert(0, "房号")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)
            for k, v in self.measuresPriceDict.items():  # v {地下_A级：100，地上B：100}
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # 整体措施费：50
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="包含：整体措施费、单项措施费、其他项目清单", subtitle="单位：元/m2"),
                # init_opts=opts.InitOpts(theme="white")
            )
        return table

    # 主材汇总表
    def buildingMaterialTotalSummary(self):
        # 建立一个空的元素
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="分部分项表没有数据,或匹配表无数据", subtitle="无法获取主材表")))
        flag = False
        try:
            path = r'.\data\analysis_data\分析系统匹配表.xlsx'
            with xlrd.open_workbook(path) as f:
                datas = f.sheet_by_name("分部分项主材匹配表")._cell_values  # 获取所有数据
                matchTabledata = datas[4:]  # 去掉表头 的匹配二维表  用于主材匹配表
                datas2 = f.sheet_by_name("分部分项主材匹配表（占地面积）")._cell_values  # 获取所有数据
                matchTabledata2 = datas2[4:]  # 去掉表头 的匹配二维表  用于主材匹配表（占地面积）
            flag = True
        except:   # 如果报错则 返回一个空的元素
            pass
        if flag == False or len(matchTabledata) < 1 or len(matchTabledata2) < 1:
            return table
        buildcounts = len(self.engineeringDict['房号信息'])
        bmtsDict = {}  # 主材总量 {"1#":{"主体砼_地下"：工程量,"二次结构_地下"：工程量,}
        bmtsDict2 = {}  # 主材总量 占地面积指标用 {"1#":{"主体砼_地下"：工程量,"二次结构_地下"：工程量,}
        for buildnum in range(buildcounts):
            buildName = self.engineeringDict['房号信息'][buildnum][0]  # 每一个房号名字
            branchdatas = self.engineeringDict['房号信息'][buildnum][4]  # 分部分项表
            if buildName:
                bmtsDict[buildName] = {}  # 主材汇总表的字典
                bmtsDict2[buildName] = {}  # 主材汇总表的字典 占地面积用
                self.subItemPriceDict[buildName] = {}  # 分部分项单价 字典  {1#:{名称+特征+地上地下+单位：0，...}}
                self.subItemQuantitiesDict[buildName] = {}  # 分部分项工程量 字典  {1#:{名称+特征+地上地下+单位：0，...}}
            # 分部分项表处理
            if branchdatas:  # 有分部分项表
                branchlists = ast.literal_eval(branchdatas)
                for row in range(len(branchlists)):
                    str0 = branchlists[row][0].strip()  # 序号
                    str2 = branchlists[row][2]  # 地上地下
                    str4 = branchlists[row][4].strip()  # 项目名称
                    str5 = branchlists[row][5].strip()  # 项目特征描述
                    str6 = branchlists[row][6]  # 计量\n单位
                    str7 = branchlists[row][7]  # 工程量
                    str9 = branchlists[row][9]  # 单价
                    str10 = branchlists[row][10]  # 合价
                    if str0:  # 分项工程名称
                        if str0[0] in "ABCDEFG":  # 首字符带ABCDEFG的为 分项工程名称
                            subName = str0
                    try:
                        flo7 = round(float(str7), 3)  # 工程量
                    except:  # 如果报错 说明工程量不是数值 则跳至下一条清单
                        flo7 = 0
                    try:  # 试着将单价转为数值型
                        flo9 = round(float(str9), 2)  # 单价
                    except:
                        flo9 = 0
                    try:  # 试着将合价转为数值型
                        flo10 = float(str10)  # 合价
                    except:
                        flo10 = 0
                    # 处理分部分项 单价 含量 表
                    if str4 and flo10:  # 有项目名称 和 合价才取 分部分项子目
                        # 分部分项单价
                        ksywords = str4 + '_' + str5[:10] + '_' + str2 + '_' + str6
                        result = self.subItemPriceDict[buildName].get(ksywords)
                        if result == None:
                            self.subItemPriceDict[buildName][ksywords] = flo9  # 分部分项单价  {1#:{名称+特征+地上地下+单位：0，...}}
                        else:
                            self.subItemPriceDict[buildName][ksywords] += flo9  # 分部分项单价
                        # 分部分项工程量
                        ksywords = str4 + '_' + str5[:10] + '_' + str2 + '_' + str6
                        result = self.subItemQuantitiesDict[buildName].get(ksywords)
                        if result == None:
                            self.subItemQuantitiesDict[buildName][ksywords] = flo7  # 分部分项工程量
                        else:
                            self.subItemQuantitiesDict[buildName][ksywords] += flo7  # 分部分项工程量

                    if not str4 or not flo7:
                        continue  # 如果名称、工程量没有 则跳至下一个清单

                    # 循环匹配表  主材汇总表
                    flagmat1 = False  # 判断是否通过所有匹配表字段
                    for matchrow in matchTabledata:  # matchrow = ['主体砼', '主体结构', '', '混凝土', '', 'm3', '']
                        mat0 = matchrow[0]  # 主材名 写入字典用
                        mat1 = matchrow[1]  # 项目名称 包含
                        mat2 = matchrow[2]  # 项目名称 不含
                        mat3 = matchrow[3]  # 特征描述 包含
                        mat4 = matchrow[4]  # 特征描述 不含
                        mat5 = matchrow[5]  # 计量单位 包含
                        mat6 = matchrow[6]  # 计量单位 不含
                        if not mat0 or not mat1:  # 如果主材名和 项目名称包含关键字 一个为空 则跳过
                            continue
                        # 循环所有 包含 不含 字段 全部通过 flagmat1 = True
                        for x1 in mat1.split(','):  # 项目名称 包含
                            if x1 in str4:  # 项目名称找到关键字
                                flagmat1 = True
                                break
                        if flagmat1 == False:  # 如果项目名称中不含关键字 则跳至下一个 主材名
                            continue
                        if mat2:
                            for x2 in mat2.split(','):  # 项目名称 不含
                                if x2 in str4:
                                    flagmat1 = False
                                    break
                        if flagmat1 == False:  # 如果项目名称不含关键字 中包含 则跳至下一个 主材名
                            continue
                        if mat3:
                            flagmat1 = False
                            for x3 in mat3.split(','):
                                if x3 in str5:
                                    flagmat1 = True
                                    break
                        if flagmat1 == False:
                            continue
                        if mat4:
                            for x4 in mat4.split(','):  # 特征描述 不含
                                if x4 in str5:
                                    flagmat1 = False
                                    break
                        if flagmat1 == False:
                            continue
                        if mat5:
                            flagmat1 = False
                            for x5 in mat5.split(','):  # 计量单位 包含
                                if x5 in str6:
                                    flagmat1 = True
                                    break
                        if flagmat1 == False:
                            continue
                        if mat6:
                            for x6 in mat6.split(','):  # 计量单位 不含
                                if x6 in str6:
                                    flagmat1 = False
                                    break
                        if flagmat1 == False:
                            continue
                        # 所有条件都可以执行 则写入字典
                        else:  # flagmat1 == True
                            if str2 != "地下":
                                str2 = "地上"
                            pingjiestr = mat0 + '_' + str2 + '_' + str6
                            result = bmtsDict[buildName].get(pingjiestr)
                            if result:  # 已有字典
                                bmtsDict[buildName][pingjiestr] += flo7
                            else:  # 创建
                                bmtsDict[buildName][pingjiestr] = flo7
                            break  # 跳出匹配表循环

                    # 循环匹配表  主材汇总表 占地面积用
                    flagmat1 = False  # 判断是否通过所有匹配表字段
                    for matchrow in matchTabledata2:  # matchrow = ['主体砼', '主体结构', '', '混凝土', '', 'm3', '']
                        mat0 = matchrow[0]  # 主材名 写入字典用
                        mat1 = matchrow[1]  # 项目名称 包含
                        mat2 = matchrow[2]  # 项目名称 不含
                        mat3 = matchrow[3]  # 特征描述 包含
                        mat4 = matchrow[4]  # 特征描述 不含
                        mat5 = matchrow[5]  # 计量单位 包含
                        mat6 = matchrow[6]  # 计量单位 不含
                        if not mat0 or not mat1:  # 如果主材名和 项目名称包含关键字 一个为空 则跳过
                            continue
                        # 循环所有 包含 不含 字段 全部通过 flagmat1 = True
                        for x1 in mat1.split(','):  # 项目名称 包含
                            if x1 in str4:  # 项目名称找到关键字
                                flagmat1 = True
                                break
                        if flagmat1 == False:  # 如果项目名称中不含关键字 则跳至下一个 主材名
                            continue
                        if mat2:
                            for x2 in mat2.split(','):  # 项目名称 不含
                                if x2 in str4:
                                    flagmat1 = False
                                    break
                        if flagmat1 == False:  # 如果项目名称不含关键字 中包含 则跳至下一个 主材名
                            continue
                        if mat3:
                            flagmat1 = False
                            for x3 in mat3.split(','):
                                if x3 in str5:
                                    flagmat1 = True
                                    break
                        if flagmat1 == False:
                            continue
                        if mat4:
                            for x4 in mat4.split(','):  # 特征描述 不含
                                if x4 in str5:
                                    flagmat1 = False
                                    break
                        if flagmat1 == False:
                            continue
                        if mat5:
                            flagmat1 = False
                            for x5 in mat5.split(','):  # 计量单位 包含
                                if x5 in str6:
                                    flagmat1 = True
                                    break
                        if flagmat1 == False:
                            continue
                        if mat6:
                            for x6 in mat6.split(','):  # 计量单位 不含
                                if x6 in str6:
                                    flagmat1 = False
                                    break
                        if flagmat1 == False:
                            continue
                        # 所有条件都可以执行 则写入字典
                        else:  # flagmat1 == True
                            if str2 != "地下":
                                str2 = "地上"
                            pingjiestr = mat0 + '_' + str6
                            result = bmtsDict2[buildName].get(pingjiestr)
                            if result:  # 已有字典
                                bmtsDict2[buildName][pingjiestr] += flo7
                            else:  # 创建
                                bmtsDict2[buildName][pingjiestr] = flo7
                            break  # 跳出匹配表循环
        if len(bmtsDict) < 1:  # 没有提取到数据 则返回空表
            return table
        # 获取表头
        templist = []
        temp = [v for k, v in bmtsDict.items()]
        for te in temp:
            for k in te.keys():
                templist.append(k)
        headers = list(set(templist))  # 去重
        headers.sort()  # 排序
        # 获取表头  占地面积用
        templist2 = []
        temp = [v for k, v in bmtsDict2.items()]
        for te in temp:
            for k in te.keys():
                templist2.append(k)
        headers2 = list(set(templist2))  # 去重
        headers2.sort()  # 排序
        # 处理表头
        headers.insert(0, "房号")  # 最前面插入 房号 字段
        datas = []
        datas.append(headers)  # 添加首行表头
        for k, v in bmtsDict.items():
            buildname = k
            datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
            datas[-1][0] = buildname  # 写入房号
            for ka, va in v.items():  # {"主体砼_地下"：工程量,"二次结构_地下"：工程量,}
                va = round(float(va), 2)  # 处理小数位数
                col = datas[0].index(ka)  # 获得列号
                datas[-1][col] = round(va, 3)  # 小数位数处理
        # 处理表头  占地面积用
        headers2.insert(0, "房号")  # 最前面插入 房号 字段
        datas2 = []
        datas2.append(headers2)  # 添加首行表头
        for k, v in bmtsDict2.items():
            buildname = k
            datas2.append(["" for x in range(len(headers2))])  # 最后增加一行空行
            datas2[-1][0] = buildname  # 写入房号
            for ka, va in v.items():  # {"主体砼_地下"：工程量,"二次结构_地下"：工程量,}
                va = round(float(va), 2)  # 处理小数位数
                col = datas2[0].index(ka)  # 获得列号
                datas2[-1][col] = round(va, 3)  # 小数位数处理
        # 最后添加一行 总计
        totalrow = [0 for x in range(len(headers))]  # 最后增加一行空行
        rows = len(datas)
        cols = len(datas[0])
        for co in range(1, cols):  # 跳过房号
            for ro in range(1, rows):  # 跳过表头
                temp = datas[ro][co]
                if temp:  # 跳过空字符
                    totalrow[co] += temp
        # 最后添加一行 总计   占地面积用
        totalrow2 = [0 for x in range(len(headers2))]  # 最后增加一行空行
        rows = len(datas2)
        cols = len(datas2[0])
        for co in range(1, cols):  # 跳过房号
            for ro in range(1, rows):  # 跳过表头
                temp = datas2[ro][co]
                if temp:  # 跳过空字符
                    totalrow2[co] += temp
        # 处理最后一行的总计小数位数
        totalrow = [round(r, 2) for r in totalrow]
        totalrow[0] = "Σ总计"  # 写入房号
        datas.append(totalrow)  # 添加总计行
        # 处理最后一行的总计小数位数   占地面积用
        totalrow2 = [round(r, 2) for r in totalrow2]
        totalrow2[0] = "Σ总计"  # 写入房号
        datas2.append(totalrow2)  # 添加总计行

        if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
            return table
        self.bmtsDatas = datas  # 主材汇总表的二维表 赋值给公共变量
        self.bmtsDatas2 = datas2  # 主材汇总表的二维表 赋值给公共变量     占地面积用
        # print(datas)
        # 创建表格实例
        table = Table()
        headers = datas[0]
        rows = datas[1:]
        table.add(headers, rows)
        table.set_global_opts(
            title_opts=ComponentTitleOpts(title="主要材料总量汇总表",
                                          subtitle="说明：\n"
                                                   r"自定义统计主材表：匹配表在‘data\analysis_data\分析系统匹配表.xlsx’"
                                                   "\n如需模板统计，在分部分项表中点击左上角的'追加模板子目'手动添加！"),)
        return table

    # 主材含量表
    def buildingMaterialTotalProportion(self):
        # print(self.bmtsDatas)
        datas = []  # 存放含量表的二维表
        # 如果主材汇总表为空 返回一个空的表，以防报错
        if self.bmtsDatas == None:
            table = (Table().add([], []).set_global_opts(
                title_opts=ComponentTitleOpts(title="分部分项表没有数据", subtitle="无法获取主材表")))
        else:
            datas.append(self.bmtsDatas[0])
            for row in range(1, len(self.bmtsDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.bmtsDatas[row][0]] + [0 for i in range(len(self.bmtsDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                buildname = datas[row][0]  # 房号名
                for col in range(1, cols):
                    colname = datas[0][col]  # 字段名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.bmtsDatas[row][col] / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.bmtsDatas[row][col] / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="主要材料含量表",
                                              subtitle="说明：\n "
                                                       "含量对应为：地上总量/地上面积，地下总量/地下面积\n"
                                                       "子目根据“主材汇总表”动态增减"
                                              ),)
        return table

    # 占地面积含量表
    def buildingMaterialTotalProportion2(self):
        # print(self.bmtsDatas)
        datas = []  # 存放含量表的二维表
        # 如果主材汇总表为空 返回一个空的表，以防报错
        if self.bmtsDatas2 == None:
            table = (Table().add([], []).set_global_opts(
                title_opts=ComponentTitleOpts(title="分部分项表没有数据", subtitle="无法获取主材表")))
        else:
            datas.append(self.bmtsDatas2[0])
            for row in range(1, len(self.bmtsDatas2) - 1):  # 跳过表头和最后的汇总行
                item = [self.bmtsDatas2[row][0]] + [0 for i in range(len(self.bmtsDatas2[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                buildname = datas[row][0]  # 房号名
                for col in range(1, cols):
                    colname = datas[0][col]  # 字段名
                    try:
                        datas[row][col] = round(self.bmtsDatas2[row][col] *100 / self.areaDict[buildname]["占地面积"], 2)
                    except:
                        datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="占地面积 含量表 扩大100倍",
                                              subtitle="说明：\n "
                                                       "单位：单位总量与占地面积比值\n"
                                                       "占地面积为“建筑面积表”中获取，最大层面积"
                                              ),)
        return table

    # 分部分项单价对比
    def subItemPrice(self):
        if len(self.subItemPriceDict) < 1:  # 如果措施费字典为空 返回一个空的饼图
            table = (Table().add([], []).set_global_opts(
                title_opts=ComponentTitleOpts(title="分部分项单价对比没有数据", subtitle="无法统计")))
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in self.subItemPriceDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))    # 去重
            # headers.sort()  # 排序
            headers.insert(0, "清单子目名称")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)
            for k, v in self.subItemPriceDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])
                datas[-1][0] = buildname
                for ka, va in v.items():
                    col = datas[0].index(ka)
                    datas[-1][col] = va
            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 取表头 取数据
            headers = datas[0]
            rows = datas[1:]
            # 数据二维表 按第一列排序
            rows = sorted(rows, key=(lambda x: x[0]))
            # 创建表格实例
            table = Table()
            # headers = datas[0]
            # rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="分部分项单价对比",
                                              subtitle="单位：元/单位\n子目名为：项目名称_特征描述10字符_地上地下_计量单位"),
                # init_opts=opts.InitOpts(theme="white")
            )
        return table

    # 分部分项工程量含量
    def subItemQuantities(self):
        if len(self.subItemQuantitiesDict) < 1:  # 如果措施费字典为空 返回一个空的饼图
            table = (Table().add([], []).set_global_opts(
                title_opts=ComponentTitleOpts(title="分部分项工程量没有数据或没有面积数据", subtitle="无法统计")))
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in self.subItemQuantitiesDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))    # 去重
            # headers.sort()  # 排序
            headers.insert(0, "清单子目名称")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)
            for k, v in self.subItemQuantitiesDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])
                datas[-1][0] = buildname
                for ka, va in v.items():
                    col = datas[0].index(ka)
                    datas[-1][col] = va
            # 转化为含量
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                buildname = datas[row][0]  # 房号名
                for col in range(1, cols):
                    colname = datas[0][col]  # 字段名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(datas[row][col] * 100 / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(datas[row][col] * 100 / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 取表头 取数据
            headers = datas[0]
            rows = datas[1:]
            # 数据二维表 按第一列排序
            rows = sorted(rows, key=(lambda x: x[0]))
            # 创建表格实例
            table = Table()
            # headers = datas[0]
            # rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="分部分项工程量含量 扩大100倍",
                                              subtitle="说明：对应地上地下面积\n子目名为：项目名称_特征描述10字符_地上地下_计量单位"),
                # init_opts=opts.InitOpts(theme="white")
            )
        return table

    # TODO ~~~~~~~~~~~~~~~~~~~~钢筋预结算单~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # 获取和聚合钢筋预结算单的数据  辅助
    def iron_bill_cal(self):
        # 一些关键参数
        ironDict = {}
        eIronGJMKeystuple = ("楼层框架梁", "屋面框架梁", "框支梁", "框架柱", "暗柱")   # 抗震钢筋  关键字 元组 提高速度
        lowerKeysList = ["基", "地下", "负", "第-"]   # 楼层名称 地下 关键字
        meshKeysList = ["地坪", "楼面", "地面", "网", "刚防", "钢防"]   # 钢筋网片 关键字
        masonryIronKeysList = ["砌体", "拉结筋", "钢防"]    # 判断砌体加筋、板缝筋 关键字
        buildcounts = len(self.engineeringDict['房号信息'])
        for buildnum in range(buildcounts):
            buildName = self.engineeringDict['房号信息'][buildnum][0]  # 每一个房号名字
            filedatas = self.engineeringDict['房号信息'][buildnum][8]  # 钢筋工程量文件表  # 可能有多个 也可能没有
            if not filedatas:  # 没有钢筋工程量文件就跳过
                continue
            filedatas = ast.literal_eval(filedatas)
            ironDict[buildName] = {}  # 每一个房号 创建一个字典
            for filedata in filedatas:
                datas = ast.literal_eval(filedata[3])  # 每一张钢筋的原始报表 二维表
                rows = len(datas)
                cols = len(datas[0])
                for row in range(1, rows):  # 跳过第一行表头
                    # 以下提取每一行数据
                    floorName = datas[row][0]  # 楼层名称
                    partsBig = datas[row][1]  # 构件大类
                    partsSmall = datas[row][2]  # 构件小类
                    partsName = datas[row][3]  # 构件名称
                    ironLevel = datas[row][4]  # 钢筋等级
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
                    ironD = str(ironD) + "圆"
                    ironDInt = float(ironD[:-1])  # 判断直径大小用
                    kg = round(float(kg), 3)
                    hoopKg = round(float(hoopKg), 3)
                    connectNum = int(connectNum)
                    # 各种样式的汇总数据 总量 接头 按级别 数量
                    # ~~~~~~开始分类~~~~~~
                    # 级别总重
                    # if meshflag == False:  # 非网片钢筋
                    if kg:
                        qdm =ironLevel + "_" + lowUppstr  # qdm 清单名称 一级子目
                        result = ironDict[buildName].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            ironDict[buildName][qdm] = kg
                        else:  # 说明级别已存在  工程总量做累加
                            ironDict[buildName][qdm] += kg  # 结果real累加
                        # 总重
                        qdm = "总重"  # qdm 清单名称 一级子目
                        result = ironDict[buildName].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            ironDict[buildName][qdm] = kg
                        else:  # 说明级别已存在  工程总量做累加
                            ironDict[buildName][qdm] += kg  # 结果real累加

                    # 网片钢筋（总量已含）
                    if meshflag == True:
                        qdm = "网片_" + lowUppstr + ironLevel  # qdm 清单名称 一级子目
                        result = ironDict[buildName].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            ironDict[buildName][qdm] = kg
                        else:  # 说明级别已存在  工程总量做累加
                            ironDict[buildName][qdm] += kg  # 结果real累加

                    # 抗震钢筋（总量以含）
                    if partsBig in eIronGJMKeystuple:  # 在抗震钢筋备选构件名列表中
                        if not hoopKg:  # 箍筋为0  才计算 纵筋重量
                            qdm = "带E钢筋"  # qdm 清单名称 一级子目
                            result = ironDict[buildName].get(qdm)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                ironDict[buildName][qdm] = kg
                            else:  # 说明级别已存在  工程总量做累加
                                ironDict[buildName][qdm] += kg  # 结果real累加

                    # 接头个数
                    if connectNum != 0:  # 有接头数量才执行
                        qdm = connectType + "_" + lowUppstr  # qdm 清单名称 一级子目
                        result = ironDict[buildName].get(qdm)  # 找一级清单子目
                        if result == None:  # 说明key未创建
                            ironDict[buildName][qdm] = connectNum
                        else:  # 说明级别已存在  工程总量做累加
                            ironDict[buildName][qdm] += connectNum  # 结果real累加
        # print(ironDict)
        return ironDict

    # 钢筋预结算单
    def iron_bill(self):
        ironDict = self.iron_bill_cal()  # 获取和聚合钢筋预结算单的数据
        # buildNum = len(ironDict)
        # 获取表头
        templist = []
        temp = [v for k, v in ironDict.items()]
        for te in temp:
            for k in te.keys():
                templist.append(k)
        headers = list(set(templist))    # 去重
        headers.sort()  # 排序
        headers.insert(0, "房号")  # 最前面插入 房号 字段
        # 创建一个包含表头的 二维表
        datas = []

        datas.append(headers)
        for k, v in ironDict.items():  # v {地下_A级：100，地上B：100}
            buildname = k
            datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
            datas[-1][0] = buildname  # 写入房号
            for ka, va in v.items():  # 地下_A级：100
                col = datas[0].index(ka)  # 获得列号
                datas[-1][col] = int(va)
        # 创建表格实例
        table = Table()
        headers = datas[0]
        rows = datas[1:]
        table.add(headers, rows)
        table.set_global_opts(
            title_opts=ComponentTitleOpts(title="钢筋专业预结算表单-2020.09版", subtitle="重量单位kg,其中网片以含在总量中。"),
            # init_opts=opts.InitOpts(theme="white")
        )
        # table.render("table_base.html")
        # return table
        table.render((self.htmlfolder + '\\' + '钢筋预结算单.html'))  # '_data'文件夹内保存网页

    # TODO ~~~~~~~~~~~~~~~~~~~~土建清单报价汇总表~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # 土建预结算单
    def civilengineering_bill(self):
        datas = self.quo_sum_datas  # 调用数据
        # 去掉"单项措施费" 一列
        rows = len(datas)
        columns = len(datas[0])
        colNum = -1
        for col in range(columns):
            if datas[0][col] == "单项措施费":
                colNum = col
                break
        if colNum != -1:
            datas = [[row[i] for i in range(len(datas[0])) if i != colNum] for row in datas]

        # 创建表格实例
        table = Table()
        headers = datas[0]
        rows = datas[1:]
        table.add(headers, rows)
        table.set_global_opts(
            title_opts=ComponentTitleOpts(title="清单报价汇总表", subtitle="")
        )
        # table.render("table_base.html")
        # return table
        table.render((self.htmlfolder + '\\' + '土建清单报价汇总表.html'))  # '_data'文件夹内保存网页

    # TODO ~~~~~~~~~~~~~~~~~~~~土建模型分析~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # 获取土建基础数据
    # def soilWoodDataGet(self):
    #     noCalKeysList = ['不提', '不计', '不算', '辅助']   # 不提量 关键字
    #     lowerKeysList = ["基", "地下", "负", "第-"]   # 楼层名称 地下 关键字
    #
    #     buildcounts = len(self.engineeringDict['房号信息'])
    #     # connectDict = {}  # 接头字典{1#:{电渣焊_地上:0,机械连接_地下:0},}
    #     # levelDWeightDict = {}  # 级别直径重量字典{1#:{A6_地上:0,C8_地下:0},}
    #     # levelWeightDict = {}  # 级别重量字典{1#:{A_地上:0,C_地下:0},}
    #     for buildnum in range(buildcounts):  # 房号循环
    #         buildName = self.engineeringDict['房号信息'][buildnum][0]  # 每一个房号名字
    #         soilWoodFileDatas = self.engineeringDict['房号信息'][buildnum][10]  # 绘图工程量文件表
    #         if buildName:
    #             # connectDict[buildName] = {}  # 新建一个房号key
    #             # levelDWeightDict[buildName] = {}  # 新建一个房号key
    #             # levelWeightDict[buildName] = {}  # 新建一个房号key
    #         # 绘图工程量文件表 处理
    #         if ironFileDatas:
    #             ironFileDatas = ast.literal_eval(ironFileDatas)
    #             for filestr in ironFileDatas:  # 可能有多张 钢筋表的情况下
    #                 datas = ast.literal_eval(filestr[3])  #得到钢筋二维列表
    #                 # print(type(fileData), len(fileData))
    #                 for row in range(1, len(datas)):  # 跳过表头
    #                     # 以此提取每一列数据
    #                     floorName = datas[row][0]  # 楼层名称
    #                     partsBig = datas[row][1]  # 构件大类
    #                     partsSmall = datas[row][2]  # 构件小类
    #                     partsName = datas[row][3]  # 构件名称
    #                     ironLevel = datas[row][4]  # 钢筋等级
    #                     ironD = datas[row][5]  # 钢筋直径
    #                     connectType = datas[row][6]  # 接头类型
    #                     kg = datas[row][7]  # 总重kg
    #                     hoopKg = datas[row][8]  # 箍筋kg
    #                     connectNum = datas[row][9]  # 接头个数
    #                     # 判断地上、下  多轴网 多施工段 需考虑
    #                     floorName = re.sub(r"\([^\(]*\)", '', floorName)  # 多轴网文件 (轴网) 删除
    #                     lowUppstr = "地上"  # 默认都为地上
    #                     if not floorName:
    #                         continue
    #                     # 如果楼层名称在 地下楼层关键字列表内 则判断为 地下
    #                     for lowkey in lowerKeysList:
    #                         if lowkey in floorName:
    #                             lowUppstr = "地下"
    #                             break
    #                     # 判断砌体加筋、板缝筋
    #                     masonryIron = False
    #                     for makey in masonryIronKeysList:
    #                         if partsSmall in makey:
    #                             masonryIron = True  # 是砌体加筋
    #                     # 判断是否冷拔丝
    #                     # 判断是否为网片
    #                     meshflag = False  # 判断是否为网片
    #                     drawingIron = False  # 判断是否冷拔丝
    #                     if int(float(ironD)) < 6:  # 小于6圆都属于网片
    #                         meshflag = True
    #                         drawingIron = True
    #                     else:  # 大于等于6圆 公司计算口径也可能认定为网片
    #                         for meshkey in meshKeysList:
    #                             if meshkey in partsName and int(float(ironD)) < 9:  # 如果名字带有关键字 且直径 < 9 圆 判定为网片
    #                                 meshflag = True
    #                                 break
    #                     # 处理 重量 直径
    #                     ironD = str(ironD)
    #                     # ironDInt = float(ironD[:-1])  # 判断直径大小用
    #                     # kg = round(float(kg), 3)
    #                     kg = float(kg)
    #                     hoopKg = round(float(hoopKg), 3)
    #                     connectNum = int(connectNum)
    #                     # 各种样式的汇总数据 总量 接头 按级别 数量
    #                     # ~~~~~~~~~~~~~~~~~~~~~~~~~开始分类~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #                     # 接头个数
    #                     if connectNum != 0 and "绑扎" not in connectType:  # 有接头数量 并且不等于绑扎 才执行
    #                         ksywords = connectType + '_' + lowUppstr  # ksywords 接头名 + 地上|地下
    #                         result = connectDict[buildName].get(ksywords)  # 找一级清单子目
    #                         if result == None:  # 说明key未创建
    #                             connectDict[buildName][ksywords] = connectNum
    #                         else:  # 说明级别已存在  工程总量做累加
    #                             connectDict[buildName][ksywords] += connectNum  # 结果real累加
    #                     # 级别重量
    #                     if kg:  # 有重量 才执行
    #                         ksywords = lowUppstr + '_' + ironLevel   # ksywords 接头名 + 地上|地下
    #                         result = levelWeightDict[buildName].get(ksywords)  # 找一级清单子目
    #                         if result == None:  # 说明key未创建
    #                             levelWeightDict[buildName][ksywords] = kg
    #                         else:  # 说明级别已存在  工程总量做累加
    #                             levelWeightDict[buildName][ksywords] += kg  # 结果real累加
    #                     # 级别直径重量
    #                     if kg:  # 有重量 才执行
    #                         ksywords = lowUppstr + '_' + ironLevel + ironD   # ksywords 接头名 + 地上|地下
    #                         result = levelDWeightDict[buildName].get(ksywords)  # 找一级清单子目
    #                         if result == None:  # 说明key未创建
    #                             levelDWeightDict[buildName][ksywords] = kg
    #                         else:  # 说明级别已存在  工程总量做累加
    #                             levelDWeightDict[buildName][ksywords] += kg  # 结果real累加
    #     self.connectDict = connectDict  # 接头表字典 赋值给公共变量
    #     self.levelWeightDict = levelWeightDict  # 级别直径重量字典 赋值给公共变量
    #     self.levelDWeightDict = levelDWeightDict  # 级别直径重量字典 赋值给公共变量


    # 土建模型分析页签元素
    def soilWoodModelAnalysis(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="研发中", subtitle="")))
        # 获取土建基础数据
        # self.soilWoodDataGet()
        # 生成页签组件
        tab = Tab()
        # tab.add(self.weightSummary(), "研发中")  # 表
        tab.add(table, "研发中")  # 表
        tab.render((self.htmlfolder + '\\' + '土建模型分析.html'))  # '_data'文件夹内保存网页

    # TODO ~~~~~~~~~~~~~~~~~~~~钢筋模型分析~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # 获取钢筋基础数据
    def ironDataGet(self):
        noCalKeysList = ['不提', '不计', '不算', '辅助']   # 不提量 关键字
        lowerKeysList = ["基", "地下", "负", "第-"]   # 楼层名称 地下 关键字
        meshKeysList = ["地坪", "楼面", "地面", "网", "刚防", "钢防"]    # 钢筋网片 关键字
        twoStructKeysList = ["构造柱", "门框", "窗框", "抱框", "MK", "CK",
                             "圈梁", "腰梁", "系梁", "过梁", "拉结筋", "砌体"]    # 二次结构 关键字
        # masonryIronKeysList = ["砌体", "拉结筋", "钢防"]    # 判断砌体加筋、板缝筋 关键字

        buildcounts = len(self.engineeringDict['房号信息'])
        connectDict = {}  # 接头字典{1#:{电渣焊_地上:0,机械连接_地下:0},}
        connectDDict = {}  # 接头直径字典{1#:{电渣焊_地上:0,机械连接_地下:0},}
        eTypeConnectDDict = {}  # 电渣焊 构件直径
        mTypeConnectDDict = {}  # # 机械连接 构件直径
        levelDWeightDict = {}  # 级别直径重量字典{1#:{A6_地上:0,C8_地下:0},}
        levelWeightDict = {}  # 级别重量字典{1#:{A_地上:0,C_地下:0},}
        typeWeightDict = {}  # 类别重量字典{1#:{A_地上:0,C_地下:0},}
        weightDict = {}  # 重量字典{1#:{地上:0,地下:0},}
        oneTwoWeightDict = {}  # 一二次结构重字典{1#:{地上:0,地下:0},}
        for buildnum in range(buildcounts):  # 房号循环
            buildName = self.engineeringDict['房号信息'][buildnum][0]  # 每一个房号名字
            ironFileDatas = self.engineeringDict['房号信息'][buildnum][8]  # 钢筋文件表
            if buildName:
                connectDict[buildName] = {}  # 新建一个房号key
                connectDDict[buildName] = {}  # 新建一个房号key
                eTypeConnectDDict[buildName] = {}
                mTypeConnectDDict[buildName] = {}
                levelDWeightDict[buildName] = {}  # 新建一个房号key
                levelWeightDict[buildName] = {}  # 新建一个房号key
                typeWeightDict[buildName] = {}  # 新建一个房号key
                weightDict[buildName] = {}  # 新建一个房号key
                oneTwoWeightDict[buildName] = {}  # 新建一个房号key
            # 钢筋文件表 处理
            if ironFileDatas:
                ironFileDatas = ast.literal_eval(ironFileDatas)
                for filestr in ironFileDatas:  # 可能有多张 钢筋表的情况下
                    datas = ast.literal_eval(filestr[3])  #得到钢筋二维列表
                    # print(type(fileData), len(fileData))
                    for row in range(1, len(datas)):  # 跳过表头
                        # 以此提取每一列数据
                        floorName = datas[row][0]  # 楼层名称
                        partsBig = datas[row][1]  # 构件大类
                        partsSmall = datas[row][2]  # 构件小类
                        partsName = datas[row][3]  # 构件名称
                        ironLevel = datas[row][4]  # 钢筋等级
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
                        # 判断是否为网片
                        meshflag = False  # 判断是否为网片
                        for meshkey in meshKeysList:  # 在 构件名称中
                            if meshkey in partsName and int(float(ironD)) < 9:  # 如果名字带有关键字 且直径 < 9 圆 判定为网片
                                meshflag = True
                                break
                        # 判断是否为 二次结构
                        twoStructFlag = False
                        for twoS in twoStructKeysList:
                            if twoS in partsBig:  # 如果名字带有关键字 且直径 < 9 圆 判定为网片
                                twoStructFlag = True
                                break
                        # 处理 重量 直径
                        ironD = str(ironD)
                        # ironDInt = float(ironD[:-1])  # 判断直径大小用
                        # kg = round(float(kg), 3)
                        kg = float(kg)
                        hoopKg = round(float(hoopKg), 3)
                        connectNum = int(connectNum)
                        # 各种样式的汇总数据 总量 接头 按级别 数量
                        # ~~~~~~~~~~~~~~~~~~~~~~~~~开始分类~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        # 接头个数
                        if connectNum != 0 and "绑扎" not in connectType:  # 有接头数量 并且不等于绑扎 才执行
                            ksywords = lowUppstr + '_' + connectType  # ksywords 接头名 + 地上|地下
                            result = connectDict[buildName].get(ksywords)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                connectDict[buildName][ksywords] = connectNum
                            else:  # 说明级别已存在  工程总量做累加
                                connectDict[buildName][ksywords] += connectNum  # 结果real累加
                        # 接头直径个数
                        if connectNum != 0 and "绑扎" not in connectType:  # 有接头数量 并且不等于绑扎 才执行
                            ksywords = lowUppstr + '_' + connectType + '_' + ironLevel + ironD  # ksywords 地上|地下 + 接头名 + A20
                            result = connectDDict[buildName].get(ksywords)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                connectDDict[buildName][ksywords] = connectNum
                            else:  # 说明级别已存在  工程总量做累加
                                connectDDict[buildName][ksywords] += connectNum  # 结果real累加
                        # 电渣焊 构件直径接头个数
                        if connectNum != 0 and "电渣" in connectType:  # 有电渣接头数量 并且不等于绑扎 才执行
                            ksywords = lowUppstr + '_' + partsBig + '_' + ironLevel + ironD  # ksywords 地上|地下 + 接头名 + A20
                            result = eTypeConnectDDict[buildName].get(ksywords)  # 找一级清单子目
                            if result == None:
                                eTypeConnectDDict[buildName][ksywords] = connectNum
                            else:
                                eTypeConnectDDict[buildName][ksywords] += connectNum  # 结果real累加
                        # 机械接头 构件直径接头个数
                        if connectNum != 0:  # 有接头数量 并且不等于绑扎 才执行
                            if "螺纹" in connectType or "套管" in connectType:  # 有电渣接头数量 并且不等于绑扎 才执行
                                ksywords = lowUppstr + '_' + partsBig + '_' + ironLevel + ironD  # ksywords 地上|地下 + 接头名 + A20
                                result = mTypeConnectDDict[buildName].get(ksywords)  # 找一级清单子目
                                if result == None:
                                    mTypeConnectDDict[buildName][ksywords] = connectNum
                                else:
                                    mTypeConnectDDict[buildName][ksywords] += connectNum  # 结果real累加
                        # 总重量
                        if kg:  # 有重量 才执行
                            ksywords = lowUppstr   # ksywords 接头名 + 地上|地下
                            result = weightDict[buildName].get(ksywords)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                weightDict[buildName][ksywords] = kg
                            else:  # 说明级别已存在  工程总量做累加
                                weightDict[buildName][ksywords] += kg  # 结果real累加

                        # 一次二次结构、网片重量
                        if kg:  # 有重量 才执行
                            if meshflag:  #  是否为网片
                                ksywords = lowUppstr + '_网片'   # ksywords 接头名 + 地上|地下
                                result = oneTwoWeightDict[buildName].get(ksywords)  # 找一级清单子目
                                if result == None:  # 说明key未创建
                                    oneTwoWeightDict[buildName][ksywords] = kg
                                else:  # 说明级别已存在  工程总量做累加
                                    oneTwoWeightDict[buildName][ksywords] += kg  # 结果real累加
                            elif twoStructFlag:  #  是否为二次结构
                                ksywords = lowUppstr + '_二次结构'   # ksywords 接头名 + 地上|地下
                                result = oneTwoWeightDict[buildName].get(ksywords)  # 找一级清单子目
                                if result == None:  # 说明key未创建
                                    oneTwoWeightDict[buildName][ksywords] = kg
                                else:  # 说明级别已存在  工程总量做累加
                                    oneTwoWeightDict[buildName][ksywords] += kg  # 结果real累加
                            else:  #  是否为一次次结构
                                ksywords = lowUppstr + '_一次结构'   # ksywords 接头名 + 地上|地下
                                result = oneTwoWeightDict[buildName].get(ksywords)  # 找一级清单子目
                                if result == None:  # 说明key未创建
                                    oneTwoWeightDict[buildName][ksywords] = kg
                                else:  # 说明级别已存在  工程总量做累加
                                    oneTwoWeightDict[buildName][ksywords] += kg  # 结果real累加
                        # 构件大类重量
                        if kg:  # 有重量 才执行
                            ksywords = lowUppstr + '_' + partsBig  # ksywords 大类 + 地上|地下
                            result = typeWeightDict[buildName].get(ksywords)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                typeWeightDict[buildName][ksywords] = kg
                            else:  # 说明级别已存在  工程总量做累加
                                typeWeightDict[buildName][ksywords] += kg  # 结果real累加
                        # 级别重量
                        if kg:  # 有重量 才执行
                            ksywords = lowUppstr + '_' + ironLevel   # ksywords 接头名 + 地上|地下
                            result = levelWeightDict[buildName].get(ksywords)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                levelWeightDict[buildName][ksywords] = kg
                            else:  # 说明级别已存在  工程总量做累加
                                levelWeightDict[buildName][ksywords] += kg  # 结果real累加
                        # 级别直径重量
                        if kg:  # 有重量 才执行
                            ksywords = lowUppstr + '_' + ironLevel + ironD   # ksywords 接头名 + 地上|地下
                            result = levelDWeightDict[buildName].get(ksywords)  # 找一级清单子目
                            if result == None:  # 说明key未创建
                                levelDWeightDict[buildName][ksywords] = kg
                            else:  # 说明级别已存在  工程总量做累加
                                levelDWeightDict[buildName][ksywords] += kg  # 结果real累加
        self.connectDict = connectDict  # 接头表字典 赋值给公共变量
        self.connectDDict = connectDDict  # 接头直径表字典 赋值给公共变量
        self.eTypeConnectDDict = eTypeConnectDDict  # 电渣焊 构件直径接头直径表字典 赋值给公共变量
        self.mTypeConnectDDict = mTypeConnectDDict  # 机械连接 构件直径接头直径表字典 赋值给公共变量
        self.levelWeightDict = levelWeightDict  # 级别直径重量字典 赋值给公共变量
        self.typeWeightDict = typeWeightDict  # 类别重量字典 赋值给公共变量
        self.levelDWeightDict = levelDWeightDict  # 级别直径重量字典 赋值给公共变量
        self.weightDict = weightDict  # 重量字典 赋值给公共变量
        self.oneTwoWeightDict = oneTwoWeightDict  # 一二次结构重字典 赋值给公共变量

    # 钢筋模型分析页签元素
    def ironModelAnalysis(self):
        # 获取钢筋基础数据
        self.ironDataGet()
        # 生成页签组件
        tab = Tab()
        tab.add(self.weightSummary(), "总重")  # 表
        tab.add(self.weightProportion(), "总含量")  # 表

        tab.add(self.oneTwoWeightSummary(), "一二次结构重")  # 表
        tab.add(self.oneTwoWeightProportion(), "一二次结构含量")  # 表

        tab.add(self.levelWeightSummary(), "级别重")  # 表
        tab.add(self.levelWeightProportion(), "级别含量")  # 表

        tab.add(self.levelDWeightSummary(), "级别直径重")  # 表
        tab.add(self.levelDWeightProportion(), "级别直径含量")  # 表

        tab.add(self.typeWeightSummary(), "构件重")  # 表
        tab.add(self.typeWeightProportion(), "构件含量")  # 表

        tab.add(self.connectSummary(), "接头数量")  # 接头分页表
        tab.add(self.connectProportion(), "接头含量")  # 接头分页表

        tab.add(self.connectDSummary(), "接头直径数量")  # 接头分页表
        tab.add(self.connectDProportion(), "接头直径含量")  # 表

        tab.add(self.eTypeConnectDSummary(), "电渣压力焊统计")  # 接头分页表
        tab.add(self.eTypeConnectDProportion(), "电渣压力焊含量")  # 表

        tab.add(self.mTypeConnectDSummary(), "机械接头统计")  # 接头分页表
        tab.add(self.mTypeConnectDProportion(), "机械接头含量")  # 表

        tab.render((self.htmlfolder + '\\' + '钢筋模型分析.html'))  # '_data'文件夹内保存网页

    # 重量汇总表
    def weightSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.weightDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "房号")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va  # int()
            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp

            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.weightDatas = datas    # 接头表字典 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋重量汇总表",
                                              subtitle="说明：\n单位：kg\n区分地上地下"),)
        return table

    # 重量含量表
    def weightProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.weightDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.weightDatas[0])
            for row in range(1, len(self.weightDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.weightDatas[row][0]] + [0 for i in range(len(self.weightDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.weightDatas[row][col] / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.weightDatas[row][col] / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋含量汇总表",
                                              subtitle="说明：\n单位：kg/m2\n区分地上地下"),)
        return table

    # 一二次结构重
    def oneTwoWeightSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.oneTwoWeightDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "大类")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va  # int()

            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.oneTwoWeightDatas = datas    # 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋一次二次结构重量汇总表",
                                              subtitle="说明：\n单位：kg\n二次结构、网片、不包含在一次结构内"),)
        return table

    # 一二次结构含量
    def oneTwoWeightProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.oneTwoWeightDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.oneTwoWeightDatas[0])
            for row in range(1, len(self.oneTwoWeightDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.oneTwoWeightDatas[row][0]] + [0 for i in range(len(self.oneTwoWeightDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.oneTwoWeightDatas[row][col] / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.oneTwoWeightDatas[row][col] / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="一次二次结构重量含量汇总表",
                                              subtitle="说明：\n单位：kg/m2\n二次结构、网片、不包含在一次结构内"),)
        return table

    # 级别重量汇总表
    def levelWeightSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.levelWeightDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "房号")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va  # int()
            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.levelWeightDatas = datas    # 接头表字典 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋级别重量汇总表",
                                              subtitle="说明：\n单位：kg\n区分地上地下"),)
        return table

    # 级别重量含量表
    def levelWeightProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.levelWeightDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.levelWeightDatas[0])
            for row in range(1, len(self.levelWeightDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.levelWeightDatas[row][0]] + [0 for i in range(len(self.levelWeightDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.levelWeightDatas[row][col] / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.levelWeightDatas[row][col] / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋级别含量汇总表",
                                              subtitle="说明：\n单位：kg/m2\n区分地上地下"),)
        return table

    # 级别直径汇总表
    def levelDWeightSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.levelDWeightDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "房号")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va
            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.levelDWeightDatas = datas    # 接头表字典 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋直径重量汇总表",
                                              subtitle="说明：\n单位：kg\n区分地上地下"),)
        return table

    # 级别直径重量含量表
    def levelDWeightProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.levelDWeightDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.levelDWeightDatas[0])
            for row in range(1, len(self.levelDWeightDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.levelDWeightDatas[row][0]] + [0 for i in range(len(self.levelDWeightDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.levelDWeightDatas[row][col] / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.levelDWeightDatas[row][col] / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋级别直径含量汇总表",
                                              subtitle="说明：\n单位：kg/m2\n区分地上地下"),)
        return table

    # 构件重量汇总表
    def typeWeightSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.typeWeightDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "构件大类")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va  # int()

            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.typeWeightDatas = datas    # 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋构件大类重量汇总表",
                                              subtitle="说明：\n单位：kg\n区分地上地下"),)
        return table

    # 构件重量含量表
    def typeWeightProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.typeWeightDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.typeWeightDatas[0])
            for row in range(1, len(self.typeWeightDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.typeWeightDatas[row][0]] + [0 for i in range(len(self.typeWeightDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.typeWeightDatas[row][col] / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.typeWeightDatas[row][col] / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋级别含量汇总表",
                                              subtitle="说明：\n单位：kg/m2\n区分地上地下"),)
        return table

    # 接头数量汇总表
    def connectSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.connectDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "接头类型")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va

            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.connectDatas = datas    # 接头表字典 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋接头汇总表",
                                              subtitle="说明：\n区分地上地下"),)
        return table

    # 接头含量表
    def connectProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.connectDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.connectDatas[0])
            for row in range(1, len(self.connectDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.connectDatas[row][0]] + [0 for i in range(len(self.connectDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.connectDatas[row][col] / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.connectDatas[row][col] / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="钢筋接头含量表",
                                              subtitle="说明：\n含量对应为：地上总量/地上面积，地下总量/地下面积"),)
        return table

    # 接头直径汇总表
    def connectDSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.connectDDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "接头级别直径")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va  # int()

            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.connectDDatas = datas    # 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="接头直径汇总表",
                                              subtitle="说明：\n单位：个\n区分地上地下、级别、直径"),)
        return table

    # 接头直径含量表
    def connectDProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.connectDDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.connectDDatas[0])
            for row in range(1, len(self.connectDDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.connectDDatas[row][0]] + [0 for i in range(len(self.connectDDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.connectDDatas[row][col] *10 / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.connectDDatas[row][col] *10 / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="接头直径含量表 *10倍 ",
                                              subtitle="说明：\n单位：个/m2*10\n区分地上地下"),)
        return table

    # 电渣焊 构件直径接头汇总表
    def eTypeConnectDSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.eTypeConnectDDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "接头级别直径")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va  # int()

            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.eTypeConnectDDatas = datas    # 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="电渣焊接头 各种类构件直径汇总表",
                                              subtitle="说明：\n单位：个\n区分地上地下、级别、直径"),)
        return table

    # 电渣焊 构件直径接头 含量表
    def eTypeConnectDProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.eTypeConnectDDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.eTypeConnectDDatas[0])
            for row in range(1, len(self.eTypeConnectDDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.eTypeConnectDDatas[row][0]] + [0 for i in range(len(self.eTypeConnectDDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.eTypeConnectDDatas[row][col] * 10 / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.eTypeConnectDDatas[row][col] * 10 / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="电渣焊接头直径含量表 *10倍 ",
                                              subtitle="说明：\n单位：个/m2*10\n区分地上地下"),)
        return table

    # 机械连接 构件直径接头汇总表
    def mTypeConnectDSummary(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        summaryDict = self.mTypeConnectDDict
        # 如果字典为空 返回一个空的表，以防报错
        if len(summaryDict) < 1:
            return table
        else:
            # 获取表头
            templist = []
            temp = [v for k, v in summaryDict.items()]
            for te in temp:
                for k in te.keys():
                    templist.append(k)
            headers = list(set(templist))  # 去重
            headers.sort()  # 排序
            headers.insert(0, "接头级别直径")  # 最前面插入 房号 字段
            # 创建一个包含表头的 二维表
            datas = []
            datas.append(headers)  # 添加首行表头
            for k, v in summaryDict.items():
                buildname = k
                datas.append(["" for x in range(len(headers))])  # 最后增加一行空行
                datas[-1][0] = buildname  # 写入房号
                for ka, va in v.items():  # {电渣焊:0,机械连接:0}
                    col = datas[0].index(ka)  # 获得列号
                    datas[-1][col] = va  # int()

            # 二维表转置
            datas = list(map(list, zip(*datas)))
            # 最后添加一行 总计 取整数
            totalrow = [0 for x in range(len(datas[0]))]  # 最后增加一行空行
            rows = len(datas)
            cols = len(datas[0])
            for co in range(1, cols):  # 跳过房号
                for ro in range(1, rows):  # 跳过表头
                    temp = datas[ro][co]
                    if temp:  # 跳过空字符
                        datas[ro][co] = int(round(temp, 0))
                        totalrow[co] += temp
            # 处理最后一行的总计小数位数
            totalrow = [int(round(r, 0)) for r in totalrow]
            totalrow[0] = "Σ总计"  # 写入房号
            datas.append(totalrow)  # 添加总计行
            # 如果没有钢筋表
            if len(datas[1:]) < 1:  # 没有提取到数据 则返回空表
                return table
            self.mTypeConnectDDatas = datas    # 赋值给公共变量
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="机械接头 各种类构件直径汇总表",
                                              subtitle="说明：\n单位：个\n区分地上地下、级别、直径 "
                                                       "\n包括：直螺纹、锥螺纹、套管等"),)
        return table

    # 机械连接 构件直径接头 含量表
    def mTypeConnectDProportion(self):
        table = (Table().add([], []).set_global_opts(
            title_opts=ComponentTitleOpts(title="工程量表中未导入钢筋表", subtitle="无法获取")))
        if self.mTypeConnectDDatas == None:
            return table
        else:
            datas = []  # 存放含量表的二维表
            datas.append(self.mTypeConnectDDatas[0])
            for row in range(1, len(self.mTypeConnectDDatas) - 1):  # 跳过表头和最后的汇总行
                item = [self.mTypeConnectDDatas[row][0]] + [0 for i in range(len(self.mTypeConnectDDatas[0]) - 1)]  # 拼接一条 ["1#",0,0,0...]
                datas.append(item)
            rows = len(datas)
            cols = len(datas[0])
            for row in range(1, rows):  # 跳过表头
                colname = datas[row][0]  # 字段名
                # # 如果没有相应的面积表
                # try:
                #     a = self.areaDict[buildname].get("地下面积")
                #     a = self.areaDict[buildname].get("地上面积")
                # except:
                #     return table
                for col in range(1, cols):
                    buildname = datas[0][col]  # 房号名
                    if "地下" in colname:
                        try:
                            datas[row][col] = round(self.mTypeConnectDDatas[row][col] * 10 / self.areaDict[buildname]["地下面积"], 2)
                        except:
                            datas[row][col] = ''
                    else:
                        try:
                            datas[row][col] = round(self.mTypeConnectDDatas[row][col] * 10 / self.areaDict[buildname]["地上面积"], 2)
                        except:
                            datas[row][col] = ''
            # 创建表格实例
            table = Table()
            headers = datas[0]
            rows = datas[1:]
            table.add(headers, rows)
            table.set_global_opts(
                title_opts=ComponentTitleOpts(title="机械接头直径含量表 *10倍 ",
                                              subtitle="说明：\n单位：个/m2*10\n区分地上地下"),)
        return table


if __name__ == '__main__':
    print('程序内执行')
    # path = r'C:\Users\wang\Desktop\123.cqd'
    # a = DataCal(path)