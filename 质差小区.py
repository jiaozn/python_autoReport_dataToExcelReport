import os
import shutil
import xlwt
import xlrd
from xlutils.copy import copy
from xlutils.margins import number_of_good_rows
from xlutils.margins import number_of_good_cols
print("输入考核次数：5/2/2/3")
count_num = int(input())
dic_dishi={"德州":[],"东营":[],"菏泽":[],"济南":[],"济宁":[],"聊城":[],"临沂":[],\
"日照":[],"泰安":[],"枣庄":[],"莱芜":[]}


def get_txt():
    items = os.listdir(".")
    list_txt = []
    for item in items:
        if item.startswith('历史性能') and item.endswith('.txt'):
            list_txt.append(item)
    if list_txt == []:
        print('没有“历史性能XXXXXXXX.txt”，输入任意键退出程序')
        input()
        os._exit()
    return list_txt


def update_result(d, d_list):
    for i in d_list:
        if d[9] == i[9]:
            if d[1] > i[1]:
                temp = i[-1]
                i = d
                i.append(temp + 1)
            else:
                i[-1] += 1
            break
    else:
        d.append(1)
        d_list.append(d)


def check_count(d_list, c):
    for it in d_list[:]:
        if int(it[-1]) < c:
            d_list.remove(it)


def ptof(str1):
    return float(str1.strip('%')) / 100


def ftop(f):
    return str(ptof(f) * 100) + "%"


yonghushouxian = []
wuxianjietonglv = []
yewudiaohualv = []
tongpinqiehuanlv = []
csfb60 = []
csfb40 = []
prb = []
low10 = []
for tf in get_txt():
    print("开始处理：" + tf)
    with open(tf, 'rU') as file:
        start_flag = 0  # 是否检测到表头，标准为行以“序号”起始
        for line in file:
            # 1.S检测是否到表头了
            if not start_flag:
                if line.startswith("序号"):  # 默认不是标头，要执行检测，这一行是不是表头
                    start_flag = 1
            # 1.E检测是否到表头了
            # 2.S开始处理数据
            else:
                one_data = line.strip('\n').split('\t')
                # S判断是不是电信站小区末尾-T
                if one_data[9].endswith('-T'):
                    continue
                # E判断是不是电信站小区末尾-T
                # S用户数受限次数
                if float(one_data[13]) >= 100:
                    update_result(one_data, yonghushouxian)
                # E用户数受限次数
                # S无线接通率
                if ptof(one_data[14]) <= 0.8 and int(one_data[15]) >= 300:
                    update_result(one_data, wuxianjietonglv)
                # E无线接通率

                # E业务掉话率
                if ptof(one_data[16]) >= 0.1 and float(one_data[17]) >= 10:
                    update_result(one_data, yewudiaohualv)
                # S业务掉话率

                # S同频切换成功率
                if ptof(one_data[18]) <= 0.8 and float(one_data[19]) >= 300:
                    update_result(one_data, tongpinqiehuanlv)
                # E同频切换成功率

                # Scsfb60
                if ptof(one_data[20]) <= 0.8 and float(one_data[21]) > 60:
                    update_result(one_data, csfb60)
                # Ecsfb60

                # Scsfb40
                if ptof(one_data[20]) == 0 and (float(one_data[21]) > 40
                                                and float(one_data[21]) <= 60):
                    update_result(one_data, csfb40)
                # Ecsfb40

                # Sprb
                if ptof(one_data[22]) >= 0.9 and (float(one_data[24]) < 10 and
                                                  float(one_data[23]) > 150):
                    update_result(one_data, prb)
                # Eprb

                # Slow10
                if float(one_data[26]) > 400 and ptof(one_data[25]) > 90:
                    update_result(one_data, low10)
                # Elow10
            # 2.E开始处理数据
        print("    数值判断完毕!")
print("-----------\n数据匹配质差列表完毕！")
#S去掉不够个数要求的记录5223
check_count(yonghushouxian, count_num)
check_count(wuxianjietonglv, count_num)
check_count(yewudiaohualv, count_num)
check_count(tongpinqiehuanlv, count_num)
check_count(csfb60, count_num)
check_count(csfb40, count_num)
check_count(prb, count_num)
check_count(low10, count_num)
#E去掉不够个数要求的记录5223
print("质差列表次数筛查完毕！")
#S按地市做表
# distri_yonghushouxian()
for yonghushouxian_item in yonghushouxian:
    dic_dishi[yonghushouxian_item[5][:2]].append(
        yonghushouxian_item[9] + "RRC用户数受限次数" + yonghushouxian_item[13])
    # distri_wuxianjietonglv()
for wuxianjietonglv_item in wuxianjietonglv:
    dic_dishi[wuxianjietonglv_item[5][:2]].append(wuxianjietonglv_item[9]+"无线接通率"+ftop(wuxianjietonglv_item[14])+\
    ",RRC请求次数"+wuxianjietonglv_item[15])
    # distri_yewudiaohualv()
for yewudiaohualv_item in yewudiaohualv:
    dic_dishi[yewudiaohualv_item[5][:2]].append(yewudiaohualv_item[9]+"，业务掉话率"+ftop(yewudiaohualv_item[16])+\
    "，掉话次数"+yewudiaohualv_item[17])
    # distri_tongpinqiehuanlv()
for tongpinqiehuanlv_item in tongpinqiehuanlv:
    dic_dishi[tongpinqiehuanlv_item[5][:2]].append(tongpinqiehuanlv_item[9]+"，同频切换成功率"+ftop(tongpinqiehuanlv_item[18])+\
    "，切换请求次数"+tongpinqiehuanlv_item[19])
    # distri_csfb60()
for csfb60_item in csfb60:
    dic_dishi[csfb60_item[5][:2]].append(csfb60_item[9]+"，CSFB成功率"+ftop(csfb60_item[20])+\
    "，CSFB请求次数"+csfb60_item[21])
    # distri_csfb40()
for csfb40_item in csfb40:
    dic_dishi[csfb40_item[5][:2]].append(csfb40_item[9]+"，CSFB成功率"+ftop(csfb40_item[20])+\
    "，CSFB请求次数"+csfb40_item[21])
    # distri_prb()
for prb_item in prb:
    dic_dishi[prb_item[5][:2]].append(prb_item[9]+"，PRB利用率"+ftop(prb_item[22])+\
    "，单用户速率"+prb_item[24]+"M，RRC最大用户数"+prb_item[23])
    # distri_low10()
for low10_item in low10:
    dic_dishi[low10_item[5][:2]].append(low10_item[9]+"，下行用户低于10M占比"+ftop(low10_item[25])+\
    "，业务量"+low10_item[26]+"M")
    #E按地市做表
print("地市分配完毕！开始excel操作")
goal_file = r"质差通报_update.xls"
model_src = r"质差通报.xls"
try:
    if not os.path.exists(goal_file) and os.path.exists(model_src):
        shutil.copy(model_src, goal_file)
        # print('复制模版成功')
except Exception as e:
    print("模版文件复制失败")
rb = xlrd.open_workbook(goal_file, formatting_info=True)
rs = rb.sheet_by_index(2)
nrows = rs.nrows
wb = copy(rb)
ws = wb.get_sheet(2)
ws.write(nrows + 1, 0, "中兴区域网络质量问题小区简报（2月8日 10:00-13:00）")
ws.write(nrows + 2, 0, "地市")
ws.write(nrows + 2, 1, "4G网络")
ws.write(nrows + 2, 2, "3G网络")
city_list = ["德州", "东营", "菏泽", "济南", "济宁", "聊城", "临沂", "日照", "泰安", "枣庄", "莱芜"]
for i in range(len(city_list)):
    ws.write(nrows + 3 + i, 0, city_list[i])
    ws.write(nrows + 3 + i, 1, ";\r\n".join(dic_dishi[city_list[i]]))
    print("写地市..." + str(i) + city_list[i])
wb.save(goal_file)
print("完成!文件名：" + goal_file + "\n Enter键退出")
input()
