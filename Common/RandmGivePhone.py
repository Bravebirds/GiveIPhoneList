#!/usr/bin/env python 3.7
# -*- coding:utf-8 -*-
'''
# FileName： RandmGivePhone.py
# Author : v_yanqyu
# Desc: PyCharm
# Date： 2020/5/15 12:55
'''
__author__ = 'v_yanqyu'
# 1.创建输出文件夹/自动引入所需第三方库
# 2.随机平均分配手机
# 3.手机/人员从Execl中获取数据
# 4.输出保存至Execl表中/设置表格式
# 导入所需要的第三方库
try:
    import os, sys, xlrd, xlwt, random, datetime
    from xlutils.copy import copy

    print("所需模块导入成功 Required modules imported successfully！")
except Exception as ImportError:
    print("本地环境没有所需第三方库 Check whether the third-party library of the local environment is downloaded")


# 定义输出路径并创建文件夹
class Mkdir():
    def makedir(filename=''):
        try:
            if os.path.exists(filename):  # 检查文件夹是否存在
                message = '文件夹已存在 Sorry Folder Already  {} Exists！'.format(filename)
                print(message)
            else:
                os.makedirs(filename)  # 越级创建文件夹
                message = '文件夹创建成功 OK Storage Path  {} Created Successfully ！'.format(filename)
                print(message)
        except Exception as NotFoundFile:  # 创建失败或其它事件则抛出异常
            print("暂时抛出异常,未知错误{}".format(NotFoundFile))


class ReadExecl():
    # 读取手机配置和人员数据源
    def read_phone(self, execlfile=''):
        global ad_senior, ad_ordinary, ios_senior, ios_ordinary, user_names,android_row_num,ios_row_num
        try:
            # 打开配置源的Execl
            data = xlrd.open_workbook(execlfile, encoding_override='utf-8')
            # 遍历所有的Sheet名字、总长度
            sheet_name = data.sheet_names()
            sheet_num = len(sheet_name)
            print("检测当前有{}个Sheet".format(sheet_num), sheet_name, "开始清理数据！")
            ad_senior = []  # 定义Android高级手机存储容器
            ad_ordinary = []  # 定义Android普通手机存储容器
            ios_senior = []  # 定义IOS高级手机存储容器
            ios_ordinary = []  # 定义IOS普通手机存储容器
            user_names = []  # 定义存储用户的容器
            error = []
            # 定义需要清洗的Sheet_name
            android = data.sheet_by_index(0)
            ios = data.sheet_by_index(1)
            user = data.sheet_by_index(2)

            # 获取Android设备信息
            for rowNum in range(android.nrows):  # 扫描Android所有的列
                ad_data = android.row_values(rowNum)
                result = ad_data[5]  # 列出特殊等级列
                # 筛选所需要的列是否符合要求 1：高级  0：普通
                if result == "等级":
                    continue  # 跳过 "标题列"
                elif result == "1":
                    ad_senior.append(android.row_values(rowNum))  # 检索高级手机添加至ad_senior列表
                elif result == "0":
                    ad_ordinary.append(android.row_values(rowNum))  # 检索普通手机添加至ad_ordinary列表
                else:
                    error.append(rowNum)
                    print("{}工作表等级队列第{}行内容有误，请先修改！！！".format(sheet_name[0], error))

            # 获取IOS设备信息
            for rowNum in range(ios.nrows):  # 扫描所有的列
                ios_data = ios.row_values(rowNum)
                result = ios_data[5]  # 列出特殊等级列
                # 筛选所需要的列是否符合要求 1：高级  0：普通
                if result == "等级":
                    continue  # 跳过"标题列"
                elif result == "1":
                    ios_senior.append(ios.row_values(rowNum))  # 检索高级手机添加至ios_senior列表
                elif result == "0":
                    ios_ordinary.append(ios.row_values(rowNum))  # 检索普通手机添加至ios_ordinary列表
                else:
                    error.append(rowNum)
                    print("{}工作表等级队列第{}行内容有误，请先修改！！！".format(sheet_name[1], error))

            # 获取用户信息
            for rowNum in range(user.nrows):
                user_data = user.row_values(rowNum)
                result = user_data[0]
                # 筛选所需要的列，去掉“姓名”标题
                if result == "姓名":
                    continue
                else:
                    user_names.append(result)  # 检索的内容append到 user_names列表中
            random.shuffle(user_names)  # 控制输出时每次排列不会重复定位
            print("已索引表格全部子目录的数据 Data source cleaned ！")
            commont = len(ad_senior), len(ad_ordinary), len(ios_senior),  len(ios_ordinary)

            # 等级平均分配异常处理
            if len(ad_senior)!=len(ad_ordinary) or len(ios_senior)!=len(ios_ordinary):
                print("请先检查对应的等级队列是否平均分配再重试！！！",commont)
                sys.exit()
            else:
                print(commont)
                return   ad_senior,ad_ordinary,ad_senior,ad_ordinary
        except Exception as FileNotFoundError:
            print("暂时抛出异常，Execl未找到 ！", FileNotFoundError)


class AssginPhone():
    def random_phone():
        global ad_senior_random, ad_ordinary_random, ios_senior_random, ios_ordinary_random, Android, Ios
        try:
            execlfile = r'../Execl/设备及人员分布.xlsx'
            ReadExecl().read_phone(execlfile=execlfile)
        except Exception as FileNotFoundError:
            print("暂时抛出异常，Execl未找到 ！", FileNotFoundError)
        Android = []  # 定义Android手机存储容器
        Ios = []  # 定义IOS手机存储容器
        disad_senior = 0
        disad_ordinary = 0
        disios_senior = 0
        disios_ordinary = 0
        # 分配Android机型
        for user_name in user_names:  # 外循环控制所需分配的用户
            # 高端机型产生
            for ad_senior_num in range(len(ad_senior)):  # 内循环控制随机产生及添加姓名字段
                if ad_senior != []:  # 数组判空处理
                    ad_senior_random = random.sample(ad_senior, 1)  # 随机生成机型、1：代表随机多少个
                    # 修改ad_senior_random插入 user_name字段
                    senior_random = ad_senior_random[0].insert(1, user_names[disad_senior])
                    Android.append(ad_senior_random)
                else:
                    continue
                disad_senior = disad_senior + 1  # 自增为user_names对应下标

                # 从原始表中移除已随机出来的机型
                for i in ad_senior:
                    for j in ad_senior_random:
                        if j in ad_senior:
                            ad_senior.remove(j)
                            # 垃圾机型产生
                if ad_ordinary != []:  # 数组判空处理
                    ad_ordinary_random = random.sample(ad_ordinary, 1)  # 随机生成机型、1：代表随机多少个
                    # 修改ad_ordinary_random插入 user_name字段
                    ad_ordinary_random[0].insert(1, user_names[disad_ordinary])
                    Android.append(ad_ordinary_random)
                else:
                    continue
                disad_ordinary = disad_ordinary + 1

                # 从原始表中移除已随机出来的机型
                for i in ad_ordinary:
                    for j in ad_ordinary_random:
                        if j in ad_ordinary:
                            # print(j,"已参与分配从队列中移除！！！")
                            remove = ad_ordinary.remove(j)

        # 随机分配IOS机型
        for user_name in user_names:
            for ios_senior_num in range(len(ios_senior)):
                # 数组判空处理
                if ios_senior != []:
                    ios_senior_random = random.sample(ios_senior, 1)  # 随机生成个数
                    ios_senior_random[0].insert(1, user_names[disios_senior])
                    Ios.append(ios_senior_random)
                else:
                    continue
                disios_senior = disios_senior + 1
                for i in ios_senior:
                    for j in ios_senior_random:
                        if j in ios_senior:
                            ios_senior.remove(j)
                # 数组判空处理
                if ios_ordinary != []:
                    # print(ios_ordinary)
                    ios_ordinary_random = random.sample(ios_ordinary, 1)  # 随机生成个数
                    ios_ordinary_random[0].insert(1, user_names[disios_ordinary])
                    Ios.append(ios_ordinary_random)
                else:
                    continue
                disios_ordinary = disios_ordinary + 1
                for i in ios_ordinary:
                    for j in ios_ordinary_random:
                        if j in ios_ordinary:
                            remove = ios_ordinary.remove(j)

    def write_execl():
        nowtime = datetime.date.today()
        style = xlwt.XFStyle()  # 格式信息
        style1 = xlwt.XFStyle()  # 格式信息
        style2 = xlwt.XFStyle()  # 格式信息
        font = xlwt.Font()  # 字体基本设置
        font.name = u'微软雅黑'
        font.color = 'black'
        font.height = 230  # 字体大小，220就是11号字体，大概就是11*20得来的吧
        style.font = font
        style1.font = font
        style2.font = font
        alignment = xlwt.Alignment()  # 设置字体在单元格的位置
        alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
        alignment.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
        style.alignment = alignment
        style1.alignment = alignment
        style2.alignment = alignment
        border = xlwt.Borders()  # 给单元格加框线
        border.left = xlwt.Borders.THIN  # 左
        border.top = xlwt.Borders.THIN  # 上
        border.right = xlwt.Borders.THIN  # 右
        border.bottom = xlwt.Borders.THIN  # 下
        border.left_colour = 0x40  # 设置框线颜色，0x40是黑色
        border.right_colour = 0x40
        border.top_colour = 0x40
        border.bottom_colour = 0x40
        # 设置背景颜色
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['sky_blue']
        style.pattern = pattern

        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['lime']
        style1.pattern = pattern

        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
        style2.pattern = pattern

        style.borders = border
        style1.borders = border
        style2.borders = border
        try:
            AssginPhone.random_phone()
            write_list = []
            print(len(user_names), "人参与随机", len(Android), "台Android设备、", len(Ios), "台IOS设备均已分配完成，即将写入Execl表中！")
            if len(Android)!=len(Ios):
                print("分发机型不足跳过运行 [Android：{}台、IOS：{}台]".format( len(Android),len(Ios)))
                sys.exit(0)
            else:
                wb = xlwt.Workbook()
                ws = wb.add_sheet(u'{}手机分配'.format(nowtime), cell_overwrite_ok=True)
                # 创建sheet将数据写入第 i 行，第 j 列
                title = ["Android手机编号", "姓名", "手机机型", "系统版本", "Android资产编号","分辨率","等级", "IOS手机编号", "姓名", "手机机型", "系统版本", "IOS资产编号","分辨率","等级"]
                # 写入首行自定义标题
                for i in range(len(title)):
                    ws.write(0, i, title[i], style)  # 外加样式
                # 追加分配机型写入
                i = 0
                for list in range(len(Android)):
                    Android[list][0].extend(Ios[list][0])
                    write_list = Android[list][0]
                    # print(write_list)
                    for j in range(len(write_list)):
                        if i in (0, 1, 4, 5, 8, 9, 12, 13, 16, 17, 20, 21,24,25,28,29,32,33,36,37,40,41):
                            ws.write(i + 1, j, write_list[j], style1)
                        else:
                            ws.write(i + 1, j, write_list[j], style2)
                        # 设置单元格宽度 控制列表自增长
                        if i in (0,2, 4, 5, 7, 8,9, 10,11,12):
                            ws.col(i).width = 5100
                        else:
                            ws.col(i).width = 2600
                    print(write_list)
                    i = i + 1
                wb.save("../Result/{}手机分配.xls".format(nowtime))  # 保存文件
                print("已写入Execl！！！详情见附件{}手机分配.xls".format(nowtime))
                return write_list
        except Exception as IndexError:
            print("文件写入失败Error：",IndexError,len(Android),len(Ios)) # 报错类型：list index out of range 这里需要看看源表是否正常


if __name__ == '__main__':
    Mkdir.makedir(filename="../Result")
    AssginPhone.write_execl()

