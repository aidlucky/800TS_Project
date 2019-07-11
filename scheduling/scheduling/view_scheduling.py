from django.shortcuts import render
import os
from django.http import FileResponse
from django.utils.http import urlquote
from django.http import HttpResponseRedirect
import xlrd
import xlsxwriter
import logging
from datetime import *
import random
import math
import time
import copy

REST_HOURS = 14  # 班次间隔
CONTINUE_WORK_DAYS = 5  # 连续上班天数
AVERAGE_OFF = False  # 平均分配员工休息时间
AVERAGE_SHIFT = False  # 平均分配员工不同班次
SAME_GROUP = False  # 同一组别班次相同


# 主视图函数
def sehceduing(request):
    if request.method == 'GET':
        return render(request, 'scheduling.html')

    if request.method == 'POST':
        # 第一部分 获取前端数据
        global REST_HOURS
        global CONTINUE_WORK_DAYS
        global AVERAGE_OFF
        global AVERAGE_SHIFT
        global SAME_GROUP

        REST_HOURS = int(request.POST.get('rest_hours', ''))
        CONTINUE_WORK_DAYS = int(request.POST.get('continue_work_days', ''))
        check_box_list = request.POST.getlist("checkbox_list", '')
        if 'isOFF' in check_box_list:
            AVERAGE_OFF = True
        if 'isPerson' in check_box_list:
            AVERAGE_SHIFT = True
        if 'isGroup' in check_box_list:
            SAME_GROUP = True

        excel = request.FILES['excel']
        se = scheduling_excel(excel)

        # 第二部分 校验前端传过来的Excel表格中的数据
        if se.check_and_parse(excel):
            return render(request, 'scheduling.html', {'error_message': se.error_message})

        # 第三部分 生成排班结果数据
        se.generate_dataframe()

        # test code
        # print('-' * 50 + 'people_info 人员信息' + '-' * 200)
        # print(se.people_info)
        #
        # print('-' * 50 + 'shift_info 班次信息' + '-' * 200)
        # print(se.shift_info)
        #
        # print('-' * 50 + 'sheduling_info 排班周期' + '-' * 200)
        # print(se.sheduling_info)
        #
        # print('-' * 50 + 'date_list 排班日期' + '-' * 200)
        # print(se.date_list)

        # print('-' * 50 + 'dataframe 排班结果' + '-' * 200)
        # print(se.dataframe)

        # for var in se.dataframe:
        #     print(str(var)+',')

        # for var in se.people_info:
        #     print(str(var)+':'+str(se.people_info[var]))

        # print(se.every_date_off_num)

        # 第四部分 生成排班Excel
        se.generate_excel()

        # 第五部分 将生成的Excel表格响应给前端
        if request.method == 'POST':
            file_name = "排班.xlsx"
            file = open(os.getcwd() + '\\' + file_name, 'rb')
            response = FileResponse(file)
            response['Content-Type'] = 'application/octet-stream'
            response['Content-Disposition'] = 'attachment;filename={0}'.format(urlquote(file_name))
            return response

# 第六部分 下载模板
def get_template(request):
    if request.method == 'GET':
        return render(request, 'scheduling.html')

    if request.method == 'POST':
        file_name = "排班模板.xlsx"
        file = open(os.getcwd() + '\\' + file_name, 'rb')

        response = FileResponse(file)
        response['Content-Type'] = 'application/octet-stream'
        response['Content-Disposition'] = 'attachment;filename={0}'.format(urlquote(file_name))
        return response


# 第七部分 功能函数和类
# excel列名转换
def column_to_name(colnum):
    if type(colnum) is not int:
        return colnum

    str = ''

    while (not (colnum // 26 == 0 and colnum % 26 == 0)):

        temp = 25

        if (colnum % 26 == 0):
            str += chr(temp + 65)
        else:
            str += chr(colnum % 26 - 1 + 65)

        colnum //= 26
        # print(str)
    # 倒序输出拼写的字符串
    return str[::-1]


def get_week_day(date):
    week_day_dict = {
        0: 'Mon',
        1: 'Tue',
        2: 'Wed',
        3: 'Thu',
        4: 'Fri',
        5: 'Sat',
        6: 'Sun',
    }
    day = date.weekday()
    return week_day_dict[day]


# 定义处理数据类
class scheduling_excel():
    def __init__(self, excel):
        # 数据校验
        # 变量声明
        self.people_info = {}  # sheet 人员
        self.shift_info = {}  # sheet 班次
        self.sheduling_info = {}  # sheet 排班周期
        self.error_message = []  # 传给 前端模板，显示到网页
        self.date_list = []  # 整个周期内的所有时间 比如从5月1日到6月1日
        self.every_date_off_num = {}  # 每一天off的数量
        self.dataframe = []  # 排班结果数据

    # 1.校验和解析excel表格中的数据
    def check_and_parse(self, excel):
        # 校验
        # 文件sheet名称校验，
        wb = xlrd.open_workbook(filename=excel.name, file_contents=excel.read())

        # 检查4个指定的sheet是否存在
        for var in ['人员', '班次', '排班周期', '指定休假']:
            if var not in wb.sheet_names():
                self.error_message.append('缺少sheet，或sheet名称错误，必须包含：人员，班次，排班周期, 指定休假')

        if self.error_message:
            return self.error_message

        # 数据校验

        # sheet2 班次数据校验
        t2 = wb.sheet_by_name('班次')
        shift_in_shift = []  # sheet2中所有的班次 类似['A1','B1','B2','C1','C5','E1','E3','通宵']

        for row in range(2, t2.nrows):
            shift_in_shift.append(t2.cell(row, 0).value)
            if not all([t2.cell(row, 0).value, t2.cell(row, 1).value, t2.cell(row, 2).value, t2.cell(row, 3).value]):
                self.error_message.append('班次表第' + str(row + 1) + '行，数据不完整!,班次,上班时间,下班时间，出勤工时，班次属性分值 为必填！')
            else:
                try:
                    xlrd.xldate.xldate_as_datetime(t2.cell(row, 1).value, 0)
                except:
                    self.error_message.append('班次表第' + str(row + 1) + '行，上班时间  格式不正确！')
                try:
                    xlrd.xldate.xldate_as_datetime(t2.cell(row, 2).value, 0)
                except:
                    self.error_message.append('班次表第' + str(row + 1) + '行，下班时间 格式不正确！')
                try:
                    float((t2.cell(row, 3).value))
                except:
                    self.error_message.append('人员表第' + str(row + 1) + '行，出勤工时必须是一个数字！')

                # 2019/6/14 add 更新check的逻辑
                try:
                    if int(t2.cell(row, 4).value) not in [0, 10, 20, 30, 40, 50]:
                        self.error_message.append('班次表第%d行，班次属性分值 %s  不是列表[0,10,20,30,40,50]中的一个数字！' % (
                            row + 1, t2.cell(row, 4).value))
                except Exception as e:
                    self.error_message.append('班次表第%d行，班次属性分值 %s  不是列表[0,10,20,30,40,50]中的一个数字！' % (
                        row + 1, t2.cell(row, 4).value))

                    # 如果 是否轮班后休息 这一列有值，它必须是 Y，也可以没有值
                if t2.cell(row, 5).value and (t2.cell(row, 5).value != "Y"):
                    self.error_message.append('班车表第' + str(row + 1) + '行，是否轮班后休息,格式不正确！')

        if self.error_message:
            return self.error_message

        # sheet1 人员数据校验
        t1 = wb.sheet_by_name('人员')

        # 验证表头部分字符串是否都存在
        row = 1
        if not all([t1.cell(row, 0).value, t1.cell(row, 1).value, t1.cell(row, 2).value
                       , t1.cell(row, 3).value, t1.cell(row, 4).value, t1.cell(row, 5).value
                       , t1.cell(row, 6).value, t1.cell(row, 7).value, t1.cell(row, 8).value
                       , t1.cell(row, 9).value, t1.cell(row, 10).value
                    ]):
            self.error_message.append('人员表第' + str(row + 1) + '行，数据不完整!,组别,姓名,组长等等 为必填！')

        group = []  # 用于校验组，只允许有一个组别
        people = []  # 用于校验人名，不允许姓名重复

        for row in range(2, t1.nrows):
            group.append(t1.cell(row, 0).value)  # 例如 语音
            people.append(t1.cell(row, 1).value)  # 人名 董丽红

            # 校验日期格式,日期是可以不填写的，所以要判断是否填写了内容
            if not len(t1.cell(row, 2).value):
                self.error_message.append('人员表第' + str(row + 1) + '行，组长名称 需要有值！')

            if t1.cell(row, 3).value:  # 有值，则转换一下试试看
                try:
                    float(t1.cell(row, 3).value)
                except:
                    self.error_message.append('人员表第' + str(row + 1) + '行，人员属性分值 请写入一个数字！')
            else:
                self.error_message.append('人员表第' + str(row + 1) + '行，人员属性分值 需要有值！')

            if t1.cell(row, 4).value:
                try:
                    xlrd.xldate.xldate_as_datetime(t1.cell(row, 4).value, 0)
                except:
                    self.error_message.append('人员表第' + str(row + 1) + '行，周期前最后工作（下班）日期时间  格式不正确！')

            if t1.cell(row, 5).value:
                try:
                    xlrd.xldate.xldate_as_datetime(t1.cell(row, 5).value, 0)
                except:
                    self.error_message.append('人员表第' + str(row + 1) + '行，加入日期  格式不正确！')

            # 离职日期/最后上班日期  可有可无，如果有值，必须是可以转成日期的字符串
            # 如果为空，则什么都不做
            if t1.cell(row, 6).value:
                try:
                    xlrd.xldate.xldate_as_datetime(t1.cell(row, 6).value, 0)
                except:
                    self.error_message.append('人员表第' + str(row + 1) + '行，离职日期/最后上班日期  格式不正确！')

            # 是否行政班
            # 这个位置的值，如果有必须是 'Y'，如果为空则什么都不做
            if t1.cell(row, 7).value:
                if t1.cell(row, 7).value == "Y":
                    pass
                else:
                    self.error_message.append('人员表第' + str(row + 1) + '行，行政班，如果填的话，必须是大写字母 Y')
            else:
                pass

            # 班次偏好1 班次偏好2 排班性质 待定，应该先把t2 赋值并校验，在来做t1的事情
            if t1.cell(row, 8).value:
                if t1.cell(row, 8).value in shift_in_shift:
                    pass
                else:
                    self.error_message.append('人员表第' + str(row + 1) + '行，班次偏好1  没有出现在sheet班次中')
            else:
                pass

            if t1.cell(row, 9).value:
                if t1.cell(row, 9).value in shift_in_shift:
                    pass
                else:
                    self.error_message.append('人员表第' + str(row + 1) + '行，班次偏好2  没有出现在sheet班次中')
            else:
                pass

            # 排班性质  5月30日  只要中间有减号，两边有数字，就是正常,干10天休10天，干100天休100天，不必两边相加等于7
            if t1.cell(row, 10).value:
                if t1.cell(row, 10).value.strip() not in shift_in_shift:
                    self.error_message.append('人员表第' + str(row + 1) + '行，指定排班，没有存在于班次表中' + t1.cell(row, 10).value)
                #     try:
                #         float(t1.cell(row, 10).value.split('-')[0])
                #         float(t1.cell(row, 10).value.split('-')[1])
                #     except:
                #         self.error_message.append('人员表第' + str(row+1) + '行，排班性质 减号两边必须是一个数字！')
                # else:
                # self.error_message.append('人员表第' + str(row+1) + '行，排班性质 格式不正确')

        if self.error_message:
            return self.error_message

        # 校验组别数量，应该只有一个
        if len(list(set(group))) > 1:
            self.error_message.append('人员表出现多个组别，当前版本只允许有一个组别！')

        if self.error_message:
            return self.error_message

        # 校验是否有重复姓名
        # 计算人数
        people_count = len(people)
        if len(list(set(people))) < people_count:
            self.error_message.append('人员表中出现重复姓名，请检查')

        # 排班周期数据校验
        t3 = wb.sheet_by_name('排班周期')
        # 校验日期格式
        for col in range(1, t3.ncols):  # 从第0行开始
            try:
                self.date_list.append(xlrd.xldate.xldate_as_datetime(t3.cell(0, col).value, 0))
            except:
                self.error_message.append('排班周期表第' + str(col + 1) + '列 日期格式不正确！')

        shift_in_schedule = []  # 班次
        # 2019/6/21 add  修改了 t3的验证数据部分
        for row in range(1, t3.nrows):  # 从第1行开始
            for col in range(1, t3.ncols):  # 从第1列开始
                if len(str(t3.cell(row, col).value)) == 0:
                    self.error_message.append('排班周期表第' + str(row + 1) + '行,第' + str(col) + '列数据缺失！')
                else:
                    try:
                        int(t3.cell(row, col).value)
                    except:
                        self.error_message.append('排班周期表第' + str(row + 1) + '行,第' + str(col) + '列填写错误，必须为数字！')

            shift_in_schedule.append(t3.cell(row, 0).value)

        if self.error_message:
            return self.error_message

        # 校验班次与排班班次是否一致 5-30 change
        if not sorted(shift_in_schedule) == sorted(shift_in_shift):
            self.error_message.append('班次不同：班次sheet中的排班 %s ' % sorted(shift_in_shift))
            self.error_message.append('班次不同：排班周期sheet中的排班 %s' % sorted(shift_in_schedule))

        # 校验指定休假数据
        t4 = wb.sheet_by_name('指定休假')
        for row in range(1, t4.nrows):
            # 判断 sheet 指定休假 中的 人名，是否在 sheet人员中
            if t4.cell(row, 1).value not in people:
                self.error_message.append('指定休假表中：' + str(t4.cell(row, 1).value) + '不存在')

        for col in range(3, t4.ncols - 1):
            try:
                xlrd.xldate.xldate_as_datetime(t4.cell(0, col).value, 0)
            except:
                self.error_message.append('指定休假表第' + str(col + 1) + '列 日期格式不正确！')

        if str(t4.cell(0, t4.ncols - 1).value).upper() != 'TOTAL':
            self.error_message.append('指定休假表中找不到Total列！')

        # 人员信息 people_info 赋值
        for row in range(2, t1.nrows):
            name = t1.cell(row, 1).value
            try:
                self.people_info[name] = {}  # 姓名
                self.people_info[name]['group'] = t1.cell(row, 0).value  # 组别
                self.people_info[name]['leader'] = t1.cell(row, 2).value  # 组长
                self.people_info[name]['score'] = t1.cell(row, 3).value  # 人员属性分值
                # 周期前最后工作（下班）日期时间
                self.people_info[name]['before_schedule_off_date'] = xlrd.xldate.xldate_as_datetime(
                    t1.cell(row, 4).value, 0) if t1.cell(row, 4).value else (self.date_list[0] + timedelta(days=-2))
                # 加入日期
                self.people_info[name]['join_date'] = xlrd.xldate.xldate_as_datetime(t1.cell(row, 5).value,
                                                                                     0) if t1.cell(row, 5).value else ''
                # 离职日期/最后上班日期
                self.people_info[name]['quit_date'] = xlrd.xldate.xldate_as_datetime(t1.cell(row, 6).value,
                                                                                     0) if t1.cell(row, 6).value else ''
                # 是否行政班
                self.people_info[name]['week_rest'] = t1.cell(row, 7).value if t1.cell(row, 7).value else ''
                # 班次偏好1
                self.people_info[name]['banci_pianhao1'] = t1.cell(row, 8).value if t1.cell(row, 8).value else ''
                # 班次偏好2
                self.people_info[name]['banci_pianhao2'] = t1.cell(row, 9).value if t1.cell(row, 9).value else ''
                # 排班性质 的左侧 和 右侧
                self.people_info[name]['continue_work_days'] = CONTINUE_WORK_DAYS
                self.people_info[name]['continue_rest_days'] = 1
                self.people_info[name]['appoint_shift'] = str(t1.cell(row, 10).value).split('-')[0] if t1.cell(row,
                                                                                                               10).value else ''

                """
                2019/6/11 add
                前台新增"连续上班不得超过天数"
                作为填空项进行填写，默认5天；
                传过来的变量  CONTINUE_WORK_DAYS
                2019/6/12 add
                    如果excel中有指定，以excel中指定的天数为主
                    如果excel中没有指定：以前端传来的连续工作天数为主
                    （可以指定工作html_con_work_day后休息，但是会出现OFF过多的问题，需要解决)

                if not self.people_info[name]['continue_work_days']:
                    self.people_info[name]['continue_work_days'] = CONTINUE_WORK_DAYS
                if not self.people_info[name]['continue_rest_days']:
                    self.people_info[name]['continue_rest_days'] = 1
                """
            except Exception as e:
                self.error_message.append('人员表取值时报错，请检查第%d行,名字: %s' % (row + 1, name))

            for var in range(1, t4.nrows):
                if t4.cell(var, 1).value == name:  # 通过人名找到固定的行的值，得到 var
                    self.people_info[name]['total_rest'] = t4.cell(var, t4.ncols - 1).value  # 指定休息的总天数，列的最后一个位置 -1
                    self.people_info[name]['appoint_rest'] = []
                    for col in range(3, t4.ncols):  # 找到固定的行号var之后，从列3开始循环和 append
                        if str(t4.cell(var, col).value).upper() == 'Y':  # 指定休息,如果不是Y则不管了
                            # 符合条件，只等于 Y 的日期假如到列表  t4.cell(0,col)  ,第 0 行，从3开始的某一列
                            self.people_info[name]['appoint_rest'].append(
                                xlrd.xldate.xldate_as_datetime(t4.cell(0, col).value, 0))

        # 班次 赋值
        for row in range(2, t2.nrows):
            shift = t2.cell(row, 0).value
            self.shift_info[shift] = {}
            self.shift_info[shift]['start_work_time'] = xlrd.xldate.xldate_as_datetime(t2.cell(row, 1).value, 0)
            self.shift_info[shift]['off_work_time'] = xlrd.xldate.xldate_as_datetime(t2.cell(row, 2).value, 0)
            self.shift_info[shift]['attendance_hours'] = float(t2.cell(row, 3).value)
            self.shift_info[shift]['banci_score'] = float(t2.cell(row, 4).value)
            self.shift_info[shift]['mustrest'] = t2.cell(row, 5).value

        # 排班周期 赋值
        for row in range(1, t3.nrows):
            shift = t3.cell(row, 0).value
            self.sheduling_info[shift] = {}
            for col in range(1, t3.ncols):
                self.sheduling_info[shift][xlrd.xldate.xldate_as_datetime(t3.cell(0, col).value, 0)] = t3.cell(row,
                                                                                                               col).value

        # 如果实际人数小于排班周期中实际需要的人数，报错到前端  2019/6/21 add
        sum1 = 0
        for d in sorted(self.date_list):
            for s in self.sheduling_info:
                sum1 += self.sheduling_info[s][d]

            if sum1 > len(self.people_info):
                self.error_message.append(d.strftime('%Y-%m-%d') + ' 这一天人力不能满足实际需要，不能生成排班表,请检查表格。 ')
                break
            sum1 = 0

        # 计算每一天off的数量
        if not self.error_message:
            for var1 in self.date_list:
                total_shift = 0
                for var2 in self.sheduling_info:  # 遍历 sheet排班周期 大字典的每一行  var 是 排班代号
                    for var3 in self.sheduling_info[var2]:
                        if var1 == var3:
                            total_shift = total_shift + self.sheduling_info[var2][var3]

                total_people = 0
                for p in self.people_info:
                    # 加入日期，离职日期,班次为空
                    if not ((self.people_info[p]['join_date'] and self.people_info[p]['join_date'] > var1) or (
                            self.people_info[p]['quit_date'] and self.people_info[p]['quit_date'] < var1)):
                        total_people = total_people + 1
                self.every_date_off_num[var1] = total_people - total_shift

        # 计算每个人应该分配的休息天数
        # for name in self.people_info:
        #     rest_days, work_days = self.get_people_off_and_work_days(name)
        #     self.people_info[name]['rest_days'] = rest_days
        #     # off 与 work的比重
        #     self.people_info[name]['off_work_rate'] = float((rest_days / work_days))
        #     # print(self.people_info[name]['off_work_rate'])
        return self.error_message

    # 2.生成排班结果数据
    def generate_dataframe(self):
        # 生成指定休息日
        for d in self.date_list:  # 循环 排班周期 中的每一天，比如 从5月1日到 5月30日
            for p in self.people_info:  # 循环大字典中的每一个人
                # 规则1，加入日期，离职日期,班次为空
                gz1 = (self.people_info[p]['join_date'] and self.people_info[p]['join_date'] > d) \
                      or (self.people_info[p]['quit_date'] and self.people_info[p]['quit_date'] < d)

                if gz1:
                    self.dataframe.append({
                        'date': d,
                        'name': p,
                        'shift': '',
                        'appoint': True,
                    })

                else:
                    self.dataframe.append({
                        'date': d,
                        'name': p,
                        'shift': '',
                        'appoint': False,
                    })

        # 当给某一天排班时，随机从现有人员中选择一人，进行排班，
        # 这样可以有效的避免班次和OFF集中在某个区域的情况
        people_list = []
        admin_shift = [] # 存放需要上行政班人员的姓名
        middle_list = []  # 存储每一天所对应的dataframe中的数据
        OFF_dic = {}  # 暂时存放连续工作的天数大于指定天数的人员

        for dat in self.date_list:
            for var in self.people_info:
                if str(self.people_info[var]['week_rest']).upper() == 'Y':
                    admin_shift.append(var)
                else:
                    people_list.append(var)

            for var1 in self.dataframe:
                if var1['date'] == dat:
                    middle_list.append(var1)

            '''
            从人员表中随机取出一个人进行当前的排班，并判断当前人员连续工作的天数是否大于限定的工作天数
            如果当前人员连续工作的天数大于限定的工作天数则暂时不给该员工进行排班，同时将该员工的姓名作为字
            典的键，连续工作天数作为字典的值存入指定的字典中。
            '''

            while True:
                if not admin_shift:
                    break
                Flag = 'ADMIN'
                admin_person = random.choice(admin_shift)

                # 判断当前员工是否已经到岗或离职
                if (self.people_info[admin_person]['join_date'] and self.people_info[admin_person]['join_date'] > dat) \
                        or (self.people_info[admin_person]['quit_date'] and self.people_info[admin_person]['quit_date'] < dat):
                    admin_shift.remove(admin_person)
                    continue

                # 判断员工某一天是否指定要休息
                if dat in self.people_info[admin_person]['appoint_rest']:
                    OFF_dic[admin_person] = 1000
                    admin_shift.remove(admin_person)
                    continue

                for dic in middle_list:
                    if not dic['appoint'] and dic['date'] == dat and dic['name'] == admin_person:
                        dic['shift'] = self.get_fit_shift(dic['date'], dic['name'], Flag)
                        admin_shift.remove(admin_person)

            while True:
                if not people_list:
                    break
                Flag = 'T'
                person = random.choice(people_list)

                # 判断当前员工是否已经到岗或离职
                if (self.people_info[person]['join_date'] and self.people_info[person]['join_date'] > dat) \
                    or (self.people_info[person]['quit_date'] and self.people_info[person]['quit_date'] < dat):
                    people_list.remove(person)
                    continue

                # 判断员工某一天是否指定要休息
                if dat in self.people_info[person]['appoint_rest']:
                    OFF_dic[person] = 1000
                    people_list.remove(person)
                    continue

                for var2 in middle_list:
                    if not var2['appoint'] and var2['date'] == dat and var2['name'] == person:
                        var2['shift'] = self.get_fit_shift(var2['date'], var2['name'],Flag)
                        if type(var2['shift']) is int:
                            OFF_dic[person] = var2['shift']
                        people_list.remove(person)
            # print(OFF_dic)

            '''
            从OFF_dic字典中优先取出连续工作天数最短的人员，进行排班，因为当人员需求没有被满足时，
            让这些连续工作天数较短的员工去满足需求。因此连续工作时间最长的员工留在最后排，被排到OFF的概率更大一些。
            '''
            while True:
                if not OFF_dic:
                    break
                Flag = 'F'
                personal = min(OFF_dic, key=OFF_dic.get)

                for var3 in middle_list:
                    if not var3['appoint'] and var3['date'] == dat and var3['name'] == personal:
                        var3['shift'] = self.get_fit_shift(var3['date'], var3['name'],Flag)
                        del OFF_dic[personal]
            # print(dat)

            middle_list.clear()

    def get_fit_shift(self, date, people,Flag):
        dic = {}  # 用于打分
        for s in self.shift_info:
            dic.update({s: 0})  # 初始化时，所有的班次都是 0 分 OFF是1分
        dic.update({'OFF': 1})

        # 加权处理
        for s in self.shift_info:
            # OK 规则1 判断是否满足需求人数
            gz1 = self.sheduling_info[s][date] > self.get_arranged(s, date)
            if gz1:
                dic[s] = dic[s] + 2

            # OK 规则2 保证休息时间足够，距离上次下班时间超过sleep_hour小时
            # 判断是否是排班周期第一天
            if date == sorted(self.date_list)[0]:  # 判断是否是第一天
                # 获取周期前最后下班时间
                if self.people_info[people]['before_schedule_off_date']:
                    gz2 = ((self.get_shift_start_work_date_time(s, date) - self.people_info[people][
                        'before_schedule_off_date']).total_seconds() >= (REST_HOURS * 3600)) and gz1
                    if gz2 and gz1:
                        dic[s] = dic[s] + 1
            else:
                # 获取人员前一天下班时间
                yesterday_off_work_time = self.get_people_off_work_time((date + timedelta(days=-1)), people)
                # 获取班次上班时间，计算休息时长
                gz2 = ((self.get_shift_start_work_date_time(s, date) - yesterday_off_work_time).total_seconds() >= (
                        REST_HOURS * 3600)) and gz1
                if gz2 and  gz1:
                    dic[s] = dic[s] + 1


            # OFF平均化,判断当前连续工作天数是否大于或等于指定的天数，
            # 大于的话则执行以下代码，并返回该员工连续工作天数
            continue_WD_web = self.people_info[people]['continue_work_days']
            continue_WD = self.get_people_continue_work_days(date, people)
            if continue_WD >= continue_WD_web and Flag == 'T':
                return int(continue_WD)

        if Flag == 'ADMIN':
            if date.weekday() >= 5:
                # 判断给当前日期和当前人员排OFF时，当前剩余OFF数量是否足够，如果OFF数量为0，则不能给当前人员强制排OFF
                if self.current_OFF_num(date):
                    dic['OFF'] += 10000
            elif self.sheduling_info['C1'][date] > self.get_arranged('C1', date):
                dic['C1'] += 10000
            else:
                dic['B2'] += 10000

            # 获取平均休息天数和平均班次数量
            # average_off_num,average_shift_num = self.get_people_off_and_work_days()

            # Thomas code

            # OK 规则3,连续上班天数
            # gz3 = self.get_people_continue_work_days(date, people) < float(
            #     self.people_info[people]['continue_work_days'])
            # if gz3 and gz2 and gz1:
            #     dic[s] += 1

            # # 规则4 班次平均化
            # # 获取当前某个人的某个班次已排数量
            # arranged_shift_num = 0
            # for var in self.dataframe:
            #     if people == var['name'] and var['shift'] == s:
            #         arranged_shift_num += 1
            # gz4 = arranged_shift_num < average_shift_num
            # if gz1 and gz4:
            #     dic[s] += 1

            # # 规则5 OFF 平均化
            # arranged_OFF_num = 0
            # for var in self.dataframe:
            #     if people == var['name'] and var['shift'] == "OFF":
            #         arranged_OFF_num += 1
            # gz5 = arranged_OFF_num < average_off_num
            # if gz5 and gz1:
            #     dic[s] += 1

            # 规则4，连续休息天数
            # rest_days = self.get_continue_rest_days(date, people)
            # gz4 = rest_days >= float(self.people_info[people]['continue_rest_days'])
            # if gz4 and gz1:
            #     dic[s] = dic[s] + 1

            # # 规则5，同一个组长的组员班次
            # gz5 = False
            # if SAME_GROUP:
            #     gz5 = (s == self.get_same_group_shift(date, people))
            #     if gz5 and gz1 and gz2:
            #         dic[s] = dic[s] + 1

            # # 把OFF的情况，平均分配到各个组中
            #
            # # 考虑到班次偏好的情况  2019/6/18 add
            # gz6 = (s in [self.people_info[people]['banci_pianhao1'], self.people_info[people]['banci_pianhao2']])
            # if gz6 and gz1:
            #     dic[s] = dic[s] + 1
            #
            # # 排班模板 sheet4 指定休假 的 total列，需要考虑进去  2019/6/20 add
            # gz7 = self.people_info[people]['total_rest'] > self.get_arranged_shift_by_people(people, 'OFF')
            # if gz7 and gz1 and gz2:
            #     dic[s] = dic[s] + 1
            #
            # 行政班 为OFF的情况
            # gz12 = (str(self.people_info[people]['week_rest']).upper() == 'Y')
            # if gz12 and gz1 and gz2:
            #     dic[s] = dic[s] + 1

            # 班次分数平均化
            # if gz1 and gz2:
            #     # 班次平均分，
            #     average = self.cal_shift_average_score()
            #     # 1.获取当前人员已排班次的平均分
            #     # 2.如果当前人员已排班次的平均分数小于班次平均分并且当前班次的班次分值大于班次平均分数时，给该班次加分
            #     # 3. 或者当前人员已排班次的平均分数大于班次平均分并且当前班次的班次分值小于班次平均分数是，给该班次加分
            #     people_avg = self.get_people_arranged_avg_score(people)
            #     if ((people_avg < average) and (self.shift_info[s]['banci_score'] > average)) or (
            #             (people_avg > average) and self.shift_info[s]['banci_score'] < average):
            #         dic[s] = dic[s] + 10

            # 平均化OFF
            # gz15 = False
            # if (self.people_info[people]['off_work_rate']) > self.people_info[people]['continue_rest_days'] / \
            #         self.people_info[people]['continue_work_days']:
            #     s_num = 0  # 排班数量
            #
            #     for s1 in self.shift_info:
            #         s_num = s_num + self.get_arranged_shift_by_people(people, s1, date)
            #
            #     if s_num != 0 and (off_days_num > self.get_arranged('OFF', date)):
            #         if (self.people_info[people]['off_work_rate'] > self.get_arranged_shift_by_people(people, 'OFF',
            #                                                                                           date) / s_num) or (
            #                 self.get_arranged_shift_by_people(people, 'OFF', date) == 0):
            #             dic['OFF'] = dic['OFF'] + 1
            #             gz15 = True
            #
            # #  (off_days > self.get_arranged('OFF',date))  实际dataFrame ,OFF未排满
            # if (not gz3) and (off_days_num > self.get_arranged('OFF', date)) and gz15 and (not gz12):
            #     dic['OFF'] = dic['OFF'] + 1
            #
            # # 针对前一天是休息的情况。决定是否连续休息

            # if (not gz4) and (rest_days >= 1) and (off_days_num > self.get_arranged('OFF', date)) and gz15 and (
            #         not gz12):
            #     dic['OFF'] = dic['OFF'] + 1

        # 返回权重最高的班次,就是dic字典中value值最大的那个key
        return max(dic, key=dic.get)

    # 计算当前排班日期剩余的OFF数量
    def current_OFF_num(self,date):
        current_date_arranged_off_num = 0  # 当前日期已排OFF数量
        current_date_total_require_num = 0 # 当前日期总需求数量
        current_date_total_people_num = len(self.people_info) # 当前日期总人力数量
        not_join_or_quit_num = 0 # 当前日期未加入人员和已离职人员的数量

        for var in self.dataframe:
            if var['date'] == date and var['shift'] == 'OFF':
                current_date_arranged_off_num += 1

        for _,dic in self.sheduling_info.items():
            for k, v in dic.items():
                if k == date:
                    current_date_total_require_num += v

        for _,value in self.people_info.items():
            if (value['join_date'] and value['join_date'] > date) or (value['quit_date'] and value['quit_date'] < date):
                not_join_or_quit_num += 1

        # 计算当前日期总OFF数量
        current_date_total_off_num = (current_date_total_people_num - not_join_or_quit_num) - current_date_total_require_num

        # 计算当前日期剩余OFF数量
        current_date_remain_OFF_num = current_date_total_off_num - current_date_arranged_off_num

        # print('日期：',date,'当前日期总OFF数量：',current_date_total_off_num,'当前日期剩余OFF数量：',current_date_remain_OFF_num)

        if current_date_remain_OFF_num == 0:
            return False
        else:
            return True

    # 获取当前人员已排班次的平均分
    def get_people_arranged_avg_score(self, people):
        count = 0
        total_score = 0
        for var in self.dataframe:
            if var["name"] == people and (var['shift'] in self.shift_info):
                count = count + 1
                total_score = total_score + self.shift_info[var['shift']]['banci_score']
        if count == 0:
            return 0
        return total_score / count

    # 计算班次平均分
    def cal_shift_average_score(self):
        total_score_temp = 0
        total_shift_temp = 0
        for d in self.date_list:
            for s in self.sheduling_info:  # 遍历 sheet排班周期 大字典的每一行  var 是 排班代号
                for d2 in self.sheduling_info[s]:
                    if d == d2:
                        total_shift_temp = total_shift_temp + self.sheduling_info[s][d2]
                        total_score_temp = total_score_temp + (
                                    self.shift_info[s]['banci_score'] * self.sheduling_info[s][d2])

        # print(total_score_temp) # 32720.0
        # print(total_shift_temp) # 968.0
        return int(total_score_temp / total_shift_temp)

    def get_people_off_and_work_days(self):
        # 总人力
        total_people_temp = 0
        # 总班次
        total_shift_temp = 0
        # off总数量
        off_num = 0
        for d in self.date_list:
            for p in self.people_info:
                # 加入日期，离职日期,班次为空
                if not ((self.people_info[p]['join_date'] and self.people_info[p]['join_date'] > d) or (
                        self.people_info[p]['quit_date'] and self.people_info[p]['quit_date'] < d)):
                    total_people_temp = total_people_temp + 1

            for s in self.sheduling_info:  # 遍历 sheet排班周期 大字典的每一行  var 是 排班代号
                for d2 in self.sheduling_info[s]:
                    if d == d2:
                        total_shift_temp = total_shift_temp + self.sheduling_info[s][d2]

            for var in self.every_date_off_num:
                if var == d:
                    off_num = off_num + self.every_date_off_num[d]

        # 计算这个人的工作天数
        # not_work_days = 0
        # for d in self.date_list:
        #     for p in self.people_info:
        #         # 加入日期，离职日期,班次为空
        #         if p == name:
        #             if ((self.people_info[p]['join_date'] and self.people_info[p]['join_date'] > d) or (
        #                     self.people_info[p]['quit_date'] and self.people_info[p]['quit_date'] < d)):
        #                 not_work_days = not_work_days + 1
        # work_days = len(self.date_list) - not_work_days

        # test code
        # print('工作天数：',work_days) # 工作天数： 31
        # print('没有工作的天数',not_work_days) # 没有工作的天数 0
        # print('总人力：',total_people_temp) # 总人力： 1364
        # print('总班次：',total_shift_temp) # 总班次： 968.0
        # print('总OFF数量：',off_num) # 总OFF数量： 396.0

        average_off_num = round(off_num / len(self.people_info))

        average_shift_num = round((len(self.date_list) - average_off_num) / len(self.shift_info))


        return average_off_num, average_shift_num

    """
    看看这个人排了多少个这种班
    """

    def get_arranged_shift_by_people(self, people, shift, date=''):
        sumOFF = 0

        for var in self.dataframe:
            if date:
                if var['shift'] == shift and var["name"] == people and var['date'] <= date:
                    sumOFF += 1
            else:
                if var['shift'] == shift and var["name"] == people:
                    sumOFF += 1
        return sumOFF

    # 获取当天该班次已排的人数
    def get_arranged(self, shift, date):
        n = 0
        for var in self.dataframe:
            if var['date'] == date and var['shift'] == shift:
                n = n + 1
        return n

    # 获取该班次当天上班时间
    def get_shift_start_work_date_time(self, shift, date):
        if shift == 'OFF' or shift == '':
            return date + timedelta(days=-1)

        start_work_time = self.shift_info[shift]['start_work_time']  # 上班时间
        # 取日期前10 ,加上班时间后面字符，生成新的日期
        start_work_date_time = datetime.strptime(str(date)[:10] + str(start_work_time)[10:], "%Y-%m-%d %H:%M:%S")
        return start_work_date_time

    # 获取该班次当天下班时间
    def get_shift_off_work_date_time(self,shift,date):
        if shift == 'OFF':
            return date + timedelta(days=-1)
        off_work_time = self.shift_info[shift]['off_work_time']

        if str(off_work_time)[11:] < str(self.shift_info[shift]['start_work_time'])[11:]:   # 下班时间小于上班时间，说明班次跨天了
            off_work_date_time = datetime.strptime(str(date + timedelta(days=1))[:10]+str(off_work_time)[10:], "%Y-%m-%d %H:%M:%S")
        else:
            off_work_date_time = datetime.strptime(str(date)[:10]+str(off_work_time)[10:], "%Y-%m-%d %H:%M:%S")

        return off_work_date_time

    # 获取人员当天下班时间
    def get_people_off_work_time(self, date, people):
        for var in self.dataframe:
            if var['name'] == people and var['date'] == date:
                if var['shift'] == 'OFF' or not var['shift']:  # 为OFF或者 空字符串的情况
                    return date + timedelta(days=-1)
                return self.get_shift_off_work_date_time(var['shift'], date)  # shift不为OFF或空的情况

    # 获取持续工作天数
    def get_people_continue_work_days(self, date, people):
        n = 0
        d = date
        while 1:
            d = d + timedelta(days=-1)
            if self.get_arranged_shift(d, people) != 'OFF' and self.get_arranged_shift(d, people) != '' and str(
                    d) >= str(sorted(self.date_list)[0]):
                n = n + 1
            else:
                return n

    # 获取连续休息天数
    def get_continue_rest_days(self, date, people):
        n = 0
        d = date
        while 1:
            d = d + timedelta(days=-1)
            if self.get_arranged_shift(d, people) == 'OFF' and str(d) >= str(sorted(self.date_list)[0]):
                n = n + 1
            else:
                break

        return n

    # 获取已排班次
    def get_arranged_shift(self, date, people):
        for var in self.dataframe:
            if var['date'] == date and var['name'] == people:
                return var['shift']

    def get_hardest_shift(self):
        temp_dict = {}

        for row in self.shift_info:
            temp_dict.update({row: 0})
        for shift1 in self.shift_info:
            for shift2 in self.shift_info:
                if shift1 != shift2:
                    date = datetime.now()  # 随便定义一个日期
                    s1 = self.get_shift_start_work_date_time(shift1, date + timedelta(days=1))
                    s2 = self.get_shift_off_work_date_time(shift2, date)
                    if (s1 - s2).total_seconds() < (REST_HOURS * 3600):
                        temp_dict[shift1] = temp_dict[shift1] + 1
        return max(temp_dict, key=temp_dict.get)

    # 获取同一组人员同一天班次
    def get_same_group_shift(self, date, people):
        colleague = ''
        if not self.people_info[people]['leader']:
            return ''
        for name in self.people_info:
            if (self.people_info[people]['leader'] == self.people_info[name]['leader']) and (
                    people != name):  # 同一leader,不同组员
                colleague = name
                break

        if not colleague:
            return ''

        for var in self.dataframe:
            if var['name'] == colleague and var['date'] == date:
                return var['shift']
        return ''

    # 计算 某人加入日期  和 离职日期带来的强制排OFF的天数  2019/6/24 add
    def calculate_off_days_join_quit(self, people):
        a, b = 0, 0
        if self.people_info[people]['join_date']:
            a = int(str((self.people_info[people]['join_date'] - self.date_list[0])).split(' ')[0])

        if self.people_info[people]['quit_date']:
            b = int(str((self.date_list[-1] - self.people_info[people]['quit_date'])).split(' ')[0])
        # print(people,a,b)
        return a + b

    # 计算休息天数
    def calculate_people_sleep_days(self, people):
        n = 0
        for var in self.dataframe:
            if var['name'] == people and var['shift'] == 'OFF':
                n = n + 1
        return n
        # return n - self.calculate_off_days_join_quit(people)

    # 统计每人每月实际班次总分
    def calculate_people_shift_score(self, people):
        sumScore = 0
        for var in self.dataframe:
            if people == var["name"]:  # 定位到某个人
                # print(var["shift"])
                if var["shift"] != "OFF" and var["shift"] != "":
                    sumScore += self.shift_info[var["shift"]]['banci_score']
        return sumScore

    def generate_excel(self):
        workbook = xlsxwriter.Workbook('排班.xlsx')
        worksheet = workbook.add_worksheet("排班")
        shift_in_shift = []
        for s in self.shift_info:
            shift_in_shift.append(s)
        shift_in_shift.append('OFF')

        worksheet.freeze_panes(len(shift_in_shift) + 2, 1)
        worksheet.set_column(str(column_to_name(1)) + ':' + str(column_to_name(2)), 4.37)
        worksheet.set_column(str(column_to_name(3)) + ':' + str(column_to_name((len(self.date_list) + 2))), 3.87)
        format = workbook.add_format({'num_format': 'mm/dd'})
        format.set_border()
        format.set_center_across()
        format.set_font_size(9)
        format.set_font_name('Arial')
        cell_format = workbook.add_format()
        cell_format.set_center_across()
        cell_format.set_border()
        cell_format.set_font_size(9)
        cell_format.set_font_name('Arial')
        off_format = workbook.add_format()
        off_format.set_center_across()
        off_format.set_fg_color('#FFB5B5')
        off_format.set_font_color('red')
        off_format.set_border()
        off_format.set_font_size(9)
        off_format.set_font_name('Arial')
        warning_cell_format = workbook.add_format()
        warning_cell_format.set_center_across()
        warning_cell_format.set_border()
        warning_cell_format.set_font_size(9)
        warning_cell_format.set_font_name('Arial')
        warning_cell_format.set_fg_color('#FFD700')
        warning_cell_format.set_font_color('red')

        # 2019/6/21 add
        warning_cell_list = list()

        # 2019/6/21 add 给self.dataframe打补丁
        # 找到打补丁的位置
        quit_cell_format = workbook.add_format()
        quit_cell_format.set_center_across()
        quit_cell_format.set_border()
        quit_cell_format.set_font_size(9)
        quit_cell_format.set_font_name('Arial')
        quit_cell_format.set_fg_color('#CDC9C9')

        for n, d in enumerate(self.date_list):
            worksheet.write_datetime(0, n + 2, d, format)
            worksheet.write(1, n + 2, get_week_day(d), cell_format)

            for nn, s in enumerate(shift_in_shift):
                worksheet.write(nn + 2, n + 2, self.get_arranged(s, d), cell_format)

        worksheet.write(len(shift_in_shift) + 1, len(self.date_list) + 2, '休息天数', cell_format)
        worksheet.write(len(shift_in_shift) + 1, len(self.date_list) + 3, '班次总分', cell_format)  # 2019/6/14 add

        # 下面2行，仅仅是 班次表的代号
        for n, s in enumerate(shift_in_shift):
            worksheet.write(n + 2, 1, s, cell_format)

        # 下面三行，在各个列中，写入人名 和 组长名称
        for n, p in enumerate(self.people_info):
            worksheet.write(n + len(shift_in_shift) + 2, 0, p, cell_format)
            worksheet.write(n + len(shift_in_shift) + 2, 1, self.people_info[p]['leader'], cell_format)

        # 大循环，写入班次代号
        for np, p in enumerate(self.people_info):
            for nd, d in enumerate(self.date_list):
                for var in self.dataframe:
                    if var['date'] == d and var['name'] == p:
                        # if var['shift'] == '':
                        #    var['shift'] = '-' # 不能为空
                        if var['shift'] == 'OFF':  # OFF不同的格式
                            worksheet.write(np + 2 + len(shift_in_shift), nd + 2, var['shift'], off_format)
                        else:
                            # 判断是否早接晚，
                            if d == sorted(self.date_list)[0]:  # 判断是否是第一天
                                # 获取周期前最后下班时间
                                if self.people_info[p]['before_schedule_off_date']:
                                    # print(self.get_shift_start_work_date_time(var['shift'],d),self.people_info[p]['before_schedule_off_date'],d,p)
                                    if ((self.get_shift_start_work_date_time(var['shift'], d) - self.people_info[p][
                                        'before_schedule_off_date']).total_seconds() < (REST_HOURS * 3600)):
                                        worksheet.write(np + 2 + len(shift_in_shift), nd + 2, var['shift'], warning_cell_format)  # warning_cell_format
                                        warning_cell_list.append((var['shift'], p, d))  # 2019/6/21 add
                                    else:
                                        worksheet.write(np + 2 + len(shift_in_shift), nd + 2, var['shift'], cell_format)
                                else:
                                    worksheet.write(np + 2 + len(shift_in_shift), nd + 2, var['shift'], cell_format)
                            else:
                                # 获取人员前一天下班时间
                                yesterday_off_work_time = self.get_people_off_work_time((d + timedelta(days=-1)), p)
                                # 获取班次上班时间，计算休息时长

                                if ((self.get_shift_start_work_date_time(var['shift'], d) - yesterday_off_work_time).total_seconds() >= (
                                        REST_HOURS * 3600)):
                                    worksheet.write(np + 2 + len(shift_in_shift), nd + 2, var['shift'], cell_format)
                                else:
                                    worksheet.write(np + 2 + len(shift_in_shift), nd + 2, var['shift'],
                                                    warning_cell_format)  # warning_cell_format
                                    if not self.people_info[p]['quit_date']:#为修复 sheet 异常排班提示信息加入 2019/7/11 add
                                        warning_cell_list.append((var['shift'], p, d))  # 2019/6/21 add

            worksheet.write(np + 2 + len(shift_in_shift), len(self.date_list) + 2, self.calculate_people_sleep_days(p),
                            cell_format)
            # 统计班次属性总分
            worksheet.write(np + 2 + len(shift_in_shift), len(self.date_list) + 3, self.calculate_people_shift_score(p),
                            cell_format)

        # 2019/6/21 add  给上面的sheet打补丁
        for np, p in enumerate(self.people_info):
            if self.people_info[p]["join_date"]:
                # print(self.people_info[p]["join_date"],p)
                for nd, d in enumerate(self.date_list):
                    if d < self.people_info[p]["join_date"]:  # 确定某些天
                        # print(d)
                        for var in self.dataframe:
                            if var["name"] == p:
                                worksheet.write(np + 2 + len(shift_in_shift), nd + 2, " ", quit_cell_format)
            if self.people_info[p]["quit_date"]:
                # print(self.people_info[p]["quit_date"],p)
                for nd, d in enumerate(self.date_list):
                    if d > self.people_info[p]["quit_date"]:  # 确定某些天
                        # print(d)
                        for var in self.dataframe:
                            if var["name"] == p:
                                worksheet.write(np + 2 + len(shift_in_shift), nd + 2, " ", quit_cell_format)

        # 2019/6/21 add  新开sheet
        worksheetAlert = workbook.add_worksheet("异常排班提示信息")
        worksheetAlert.write(0, 0, "班次", format)  # 班次
        worksheetAlert.write(0, 1, "人名", format)  # 某人
        worksheetAlert.write(0, 2, "日期", format)  # 某天
        worksheetAlert.merge_range(0, 3, 0, 6, "提示信息", format)  # merge 从0行第3列，到0行第6列

        for n, d in enumerate(warning_cell_list):
            worksheetAlert.write(n + 1, 0, d[0], format)
            worksheetAlert.write(n + 1, 1, d[1], format)
            worksheetAlert.write(n + 1, 2, d[2], format)
            # worksheetAlert.write(n+1, 3, "两个班次之间的休息时间小于%d小时" % REST_HOURS )
            worksheetAlert.merge_range(n + 1, 3, n + 1, 6, "两个班次之间的休息时间小于%d小时" % REST_HOURS, format)


        workbook.close()





