# -*- coding: utf-8 -*-
from function import read, sunny_lunar
import xlrd
import xlwt
from xlutils.copy import copy
import datetime
import easygui as eg
import os


class Data:
    def __init__(self, date, time, sl, tag, id):
        """

        :param date: 日期
        :param time: 时间
        :param sl: 潮位
        :param tag: 标识符
        :param id: 点号
        """
        self.date = date
        self.time = time
        self.sl = float(sl)
        self.tag = tag
        self.id = id

        self.td = 0
        self.sld = 0
        self.td_s = 0

        self.day_time = "{0} {1}".format(self.date, self.time).replace("/", "-")
        self.dateTime_p = datetime.datetime.strptime(self.day_time, '%Y-%m-%d %H:%M:%S')

        # 农历计算 0、获取ydm 1、获取农历 2、转换格式mm/dd
        self.ymd = self.date.split("/")  # 拆分日期数据
        self.y = int(self.ymd[0])
        self.m = int(self.ymd[1])
        self.d = int(self.ymd[2])
        self.lm, self.ld = sunny_lunar.sunny_lunar(self.y, self.m, self.d)
        self.lunar = "{0}/{1}".format(self.lm, self.ld)


def shujuluru(file_path):
    '''
    数据读取录入
    :param file_path: 路径名
    :return: 返回分割成一个个数据的列表
    '''
    file = file_path
    data = read.read_file(file)
    # print(data)
    data_list = []
    for each in data:
        data_list.append(each.strip().split(","))
        # 删除回车，以逗号分隔成一个个数列
    # print(data_list)
    dt = []
    for each in data_list:
        a = Data(each[0], each[1], each[2], "n", data_list.index(each))
        dt.append(a)
    return dt


def shujushaixuan(dt):
    '''
    整点和高低数据挑选
    :param dt: 数据录入后拆分得到的dt数列
    :return: 返回整点和高低数据
    '''
    gd_dt = []
    for each in dt:
        if dt.index(each) < len(dt) - 1:
            a = each
            b = dt[dt.index(each) + 1]
            a_list = a.time.split(":")
            m = a_list[1]
            if m != "00" or a.time == b.time:
                gd_dt.append(dt.pop(dt.index(each)))
        # 当数据的时间不是正点或者时间相同时，弹出数据加入到高低数据中,剩下的数据就是整点数据
        else:
            a = each
            b = dt[dt.index(each) - 1]
            a_list = a.time.split(":")
            m = a_list[1]
            if m != "00" or a.time == b.time:
                gd_dt.append(dt.pop(dt.index(each)))

    for i in range(0, len(gd_dt) - 1):
        a = gd_dt[i]
        b = gd_dt[i + 1]
        td = b.dateTime_p - a.dateTime_p
        sld = b.sl - a.sl
        b.td = td
        b.sld = sld
        h_m_s = str(td).split(":")
        b.td_s = int(h_m_s[0]) * 3600 + int(h_m_s[1]) * 60 + int(h_m_s[2])
    # dayingceshi(gd_dt)

    # 选择高低数据的前两个数据的潮位比较大小,1>2时1为g,2为d,之后循环dg
    i = gd_dt[0].sl
    j = gd_dt[1].sl
    if i > j:
        gd_dt[0].tag = "g"
    else:
        gd_dt[0].tag = "d"
    for each in gd_dt:
        if gd_dt.index(each) == 0:
            continue
        else:
            if gd_dt[gd_dt.index(each) - 1].tag == "g":
                each.tag = "d"
            else:
                each.tag = "g"
    # 经过挑选数据得到整点数据dt和高低数据gd_dt
    # dayingceshi(gd_dt)
    return dt, gd_dt


def gdfenlei(dt_list, tag):
    """
    将高低水位分类
    :param dt_list: 高水位或低水位实例化对象的list
    :param tag: 修改的分类标签 只能是"g" 或者 "d" 不然会出错
    :return: 分类后的实例化对象列表
    """
    for each in dt_list:
        if dt_list.index(each) < len(dt_list) - 1:
            date1 = each.date
            time1 = each.time
            d1 = int(date1.split("/")[2])
            hour1 = int(time1.split(":")[0])

            next_each = dt_list[dt_list.index(each) + 1]
            date2 = next_each.date
            time2 = next_each.time
            d2 = int(date2.split("/")[2])
            hour2 = int(time2.split(":")[0])

            if d1 == d2:
                if hour1 < hour2:
                    each.tag = tag + "1"
                else:
                    each.tag = tag + "2"
            else:
                if hour1 < 12:
                    each.tag = tag + "1"
                else:
                    each.tag = tag + "2"
        else:
            date1 = each.date
            time1 = each.time
            d1 = int(date1.split("/")[2])
            hour1 = int(time1.split(":")[0])

            next_each = dt_list[dt_list.index(each) - 1]
            date2 = next_each.date
            time2 = next_each.time
            d2 = int(date2.split("/")[2])
            hour2 = int(time2.split(":")[0])

            if d1 == d2:
                if hour1 < hour2:
                    each.tag = tag + "1"
                else:
                    each.tag = tag + "2"
            else:
                if hour1 < 12:
                    each.tag = tag + "1"
                else:
                    each.tag = tag + "2"


def dayingceshi(file):
    """
    打印测试
    :param file:
    :return:
    """
    for each in file:
        print("日期为{0},时间为{1},潮位为{2},标识符为{3},ID为{4},农历为{5},时间戳为{6},时间差为{7},超差为{8}".format(each.date, each.time, each.sl,
                                                                                         each.tag, each.id,
                                                                                         each.lunar, each.dateTime_p,
                                                                                         each.td, each.sld))


def shujushuchu(zd_dt, gd_dt, xls_file, area="请输入区域"):
    """
    数据输出为excel
    :param zd_dt: zd_dt = [整点数据数据的实例化对象,...]
    :param gd_dt: gd_dt = [高低潮位数据的实例化对象,...]
    :return:
    """
    # ***************查询日期定为的字典***************开始
    day_row = {1: 5, 2: 6, 3: 7, 4: 8, 5: 9, 6: 10, 7: 11, 8: 12, 9: 13, 10: 14, 11: 15, 12: 16, 13: 17, 14: 18, 15: 19,
               16: 20, 17: 21, 18: 22, 19: 23, 20: 24, 21: 25, 22: 26, 23: 27, 24: 28, 25: 29, 26: 30, 27: 31, 28: 32,
               29: 33, 30: 34, 31: 35}
    hour_col = {0: 3, 1: 4, 2: 5, 3: 6, 4: 7, 5: 8, 6: 9, 7: 10, 8: 11, 9: 12, 10: 13, 11: 14, 12: 15, 13: 16, 14: 17,
                15: 18, 16: 19, 17: 20, 18: 21, 19: 22, 20: 23, 21: 24, 22: 25, 23: 26}
    # ***************查询日期定为的字典***************结束
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet1 = workbook.add_sheet(area)
    worksheet2 = workbook.add_sheet("每月统计")
    workbook.save(xls_file)

    oldWb = xlrd.open_workbook(xls_file, formatting_info=True)
    # 先打开已存在的表
    newWb = copy(oldWb)
    # 复制
    newWs = newWb.get_sheet(0)
    # 取sheet表

    # ***************写入过程***************开始
    # newWs.write(2, 4, "pass") 写入的范例
    # 1、输出农历 1,2两列
    for each in gd_dt:
        date = each.date
        dd = int(date.split("/")[2])
        row = day_row[dd]
        lunar = each.lunar
        lunar_m = int(lunar.split("/")[0])
        lunar_d = int(lunar.split("/")[1])
        newWs.write(row, 1, lunar_m)
        newWs.write(row, 2, lunar_d)
    # 2、输出整点
    for each in zd_dt:
        date = each.date
        time = each.time
        dd = int(date.split("/")[2])
        hh = int(time.split(":")[0])
        row = day_row[dd]
        col = hour_col[hh]
        sl = each.sl * 100
        newWs.write(row, col, sl)

    # 3、输出高低潮位
    # 高潮位和低潮位筛选
    g_dt = []
    d_dt = []
    for each in gd_dt:
        if each.tag == "g":
            g_dt.append(each)
        elif each.tag == "d":
            d_dt.append(each)
        else:
            pass
    print("前***********************************************************************前")
    dayingceshi(g_dt)
    print("前############################################################################钱")
    dayingceshi(d_dt)
    # todo：算法更改
    gdfenlei(g_dt, "g")
    gdfenlei(d_dt, "d")
    print("***********************************************************************")
    dayingceshi(g_dt)
    print("############################################################################")
    dayingceshi(d_dt)
    for each in g_dt:
        date = each.date
        day = int(date.split("/")[2])
        sl = each.sl * 100
        tag = each.tag
        row = day_row[day]
        dateTime_p = each.dateTime_p
        if tag == "g1":
            newWs.write(row, 29, dateTime_p)
            newWs.write(row, 30, sl)
        if tag == "g2":
            newWs.write(row, 31, dateTime_p)
            newWs.write(row, 32, sl)

    for each in d_dt:
        date = each.date
        day = int(date.split("/")[2])
        sl = each.sl * 100
        tag = each.tag
        row = day_row[day]
        dateTime_p = each.dateTime_p
        if tag == "d1":
            newWs.write(row, 33, dateTime_p)
            newWs.write(row, 34, sl)
        if tag == "d2":
            newWs.write(row, 35, dateTime_p)
            newWs.write(row, 36, sl)

    # 特征值统计，数列按从小到大排序后生成一个新数列
    g_dt_tx = sorted(g_dt, key=lambda a: a.sl)
    d_dt_tx = sorted(d_dt, key=lambda a: a.sl)

    g_dt_sld = sorted(g_dt, key=lambda a: a.sld)
    d_dt_sld = sorted(d_dt, key=lambda a: a.sld)

    g_dt_td = sorted(g_dt, key=lambda a: a.td_s)
    d_dt_td = sorted(d_dt, key=lambda a: a.td_s)

    if g_dt_td[0].td_s == 0:
        g_dt_td.pop(0)
    elif d_dt_td[0].td_s == 0:
        d_dt_td.pop(0)
    else:
        pass

    # 高低潮特征值摘取
    gc_sl_max, gc_sl_min, gc_sl_ave = tongji(g_dt_tx, "sl")
    dc_sl_max, dc_sl_min, dc_sl_ave = tongji(d_dt_tx, "sl")

    # 涨落潮潮差特征值摘取
    gc_sld_max, gc_sld_min, gc_sld_ave = tongji(g_dt_sld, "sld")
    dc_sld_max, dc_sld_min, dc_sld_ave = tongji(d_dt_sld, "sld")

    # 涨落潮潮时特征值摘取
    gc_td_max, gc_td_min, gc_td_ave = tongji(g_dt_td, "td_s")
    dc_td_max, dc_td_min, dc_td_ave = tongji(d_dt_td, "td_s")

    # ***************统计值输出至表格*************** 开始
    newWs = newWb.get_sheet(1)
    list1_tongji = [[gc_sl_max, gc_sl_min], [dc_sl_max, dc_sl_min], [gc_sld_max, gc_sld_min],
                    [dc_sld_max, dc_sld_min], [gc_td_max, gc_td_min], [dc_td_max, dc_td_min]]
    list2_tongji = [gc_sl_ave, dc_sl_ave, gc_sld_ave, dc_sld_ave, gc_td_ave, dc_td_ave]
    plus = 0
    for i in list1_tongji:
        if list1_tongji.index(i) == 0 or list1_tongji.index(i) == 1:
            for j in i:
                tuple_str = shuchu_str(j)
                # print(tuple_str)
                newWs.write(0 + plus, 0, tuple_str[0])
                newWs.write(1 + plus, 0, tuple_str[1])
                newWs.write(2 + plus, 0, tuple_str[2])
                newWs.write(3 + plus, 0, tuple_str[3])
                newWs.write(4 + plus, 0, tuple_str[3])
                plus += 5
            newWs.write(plus, 0, list2_tongji[list1_tongji.index(i)])
            plus += 1

        elif list1_tongji.index(i) == 2 or list1_tongji.index(i) == 3:
            for j in i:
                tuple_str = shuchu_str(j)
                # print(tuple_str)
                newWs.write(0 + plus, 0, tuple_str[0])
                newWs.write(1 + plus, 0, tuple_str[1])
                newWs.write(2 + plus, 0, tuple_str[3])
                newWs.write(3 + plus, 0, tuple_str[3])
                plus += 4
            newWs.write(plus, 0, list2_tongji[list1_tongji.index(i)])
            plus += 1

        elif list1_tongji.index(i) == 4 or list1_tongji.index(i) == 5:
            for j in i:
                tuple_str = shuchu_str(j)
                # print(tuple_str)
                newWs.write(0 + plus, 0, tuple_str[4])
                newWs.write(1 + plus, 0, tuple_str[1])
                newWs.write(2 + plus, 0, tuple_str[3])
                newWs.write(3 + plus, 0, tuple_str[3])
                plus += 4
            newWs.write(plus, 0, s_hms(list2_tongji[list1_tongji.index(i)]))
            plus += 1
    newWb.save(xls_file)
    # 保存至result路径

    dayingceshi(gc_td_min)


def s_hms(str_s):
    s = (int(str_s) % 3600) % 60
    m = (int(str_s) % 3600) // 60
    h = int(str_s) // 3600
    if s < 10:
        s = "0{0}".format(s)
    if m < 10:
        m = "0{0}".format(m)
    if h < 10:
        h = "0{0}".format(h)
    hms = "{0}:{1}:{2}".format(h, m, s)
    return hms


def shuchu_str(max_min):
    str_sl = ""
    str_date = ""
    str_time = ""
    str_lunar = ""
    str_td = ""
    str_sld = ""
    for each in max_min:
        str_sl = str_sl + " {0}".format(str(each.sl))
        str_date = str_date + " {0}".format(str(each.date))
        str_time = str_time + " {0}".format(str(each.time))
        str_lunar = str_lunar + " {0}".format(str(each.lunar))
        str_td = str_td + " {0}".format(str(each.td))
        str_sld = str_sld + " {0}".format(str(each.sld))
    return str_sl, str_date, str_time, str_lunar, str_td, str_sld


def tongji(list, x):
    """
    采取从小到大排序号的数列当中的最大最小和平均值，
    :param list: list为从小到大排序号的数列
    :param x: “x”为数列中需要进行比较的参数，如sl、td
    :return:
    """
    max = []
    min = []
    for i in range(0, len(list)):
        exec("if list[i].{0}==list[0].{0}:\n    min.append(list[i])\nelse:\n    pass".format(x))

    for i in range(0, len(list)):
        exec("if list[i].{0}==list[len(list)-1].{0}:\n    max.append(list[i])\nelse:\n    pass".format(x))

    n = []
    for i in range(0, len(list)):
        exec("n.append(list[i].{0})".format(x))

    sum = 0
    for each in n:
        sum += each
    ave = sum / len(list)
    return max, min, ave


def run():
    """
    单文件处理
    :return:
    """
    try:
        csv_path = eg.fileopenbox("请选择输入的潮位数据", default='*.csv')
        dt = shujuluru(csv_path)
        zd_dt, gd_dt = shujushaixuan(dt)
        xls_path = eg.filesavebox("请选择保存的位置", default='*.xls')
        area = csv_path.split("\\").pop().split(".")[0]
        shujushuchu(zd_dt, gd_dt, xls_path, area=area)
    except OSError as reason:
        eg.msgbox("运行错误！", "警告！")
        print(str(reason))


def all_run():
    """
    多文件处理
    :return:
    """
    try:
        file_path = eg.diropenbox("请选择输入的潮位数据集")
        csv_list = os.listdir(file_path)
        print(csv_list)
        path_list = []
        for each in csv_list:
            if each.split(".")[1] == "csv":
                path_list.append("test/" + each)
        print(path_list)
        for each in path_list:
            dt = shujuluru(each)
            zd_dt, gd_dt = shujushaixuan(dt)
            xls_path = each.split(".")[0] + ".xls"
            area = each.split(".")[0].split("/")[1]
            shujushuchu(zd_dt, gd_dt, xls_path, area=area)
    except OSError as reason:
        eg.msgbox("运行错误！", "警告！")
        print(str(reason))


if __name__ == "__main__":
    choice = eg.choicebox("请选择处理模式", choices=["单文件处理", "多文件处理"])
    if choice == "单文件处理":
        run()
    if choice == "多文件处理":
        all_run()
    eg.msgbox("处理完成！")
