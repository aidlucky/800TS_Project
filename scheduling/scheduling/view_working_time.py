from django.shortcuts import render
import os
from django.http import FileResponse 
from django.utils.http import urlquote
from django.http import HttpResponseRedirect
import xlrd
import xlsxwriter
import logging
from datetime import *
import re


"""
5月30日测试总结
1 假期需要体现在 假期表中，包含工时  比如A+2,XOT2 在假期表中是  OT2，多个工时用逗号分开
2 A+2,XOT2的情况合理 在拦截部分  和 计算工时部分都需要修改

        计算总工时total时只考虑第一个逗号左边的+ 和 - 后面的数字，
        不考虑第一个逗号右边的所有数字
待拦截的情况： 
    A,XOT2,XEO.2   最后一个 O.2 不是 0.2   done
    以英文逗号结尾 O,XOT2, 但是要排除 A, O,的情况  done

    计算工时部分需要在看一遍

"""


def get_working_time_template(request):
    if request.method == 'GET':
        return render(request,'scheduling.html')

    if request.method == 'POST':
        file_name = "工时模板.xlsx"
        file = open(os.getcwd()+'\\'+file_name, 'rb')

        response = FileResponse(file)
        response['Content-Type'] = 'application/octet-stream'
        response['Content-Disposition'] = 'attachment;filename={0}'.format(urlquote(file_name))
        return response

def working_time(request):
    if request.method == 'GET':
        return render(request,'working_time.html',{'attribute_list':''})

    if request.method == 'POST':
        excel = request.FILES['excel']
        #初始化对象 对象的使用，避免了反复传达参数的麻烦
        wte = working_time_excel(excel)

        #如果有错误信息则打印
        if wte.check_and_parse():
            return render(request,'working_time.html',{'error_message':wte.error_message})
        wte.generate_excel()

        # 返回生成的Excel
        file_name = "工时.xlsx"
        file = open(os.getcwd()+'\\'+file_name, 'rb')
        response = FileResponse(file)
        response['Content-Type'] = 'application/octet-stream'
        response['Content-Disposition'] = 'attachment;filename={0}'.format(urlquote(file_name))
        return response

# 根据日期获取星期的缩写
def get_week_day(date):
    week_day_dict = {
    0 : 'Mon',
    1 : 'Tue',
    2 : 'Wed',
    3 : 'Thu',
    4 : 'Fri',
    5 : 'Sat',
    6 : 'Sun',
    }
    day = date.weekday()
    return week_day_dict[day]

class working_time_excel:
    def __init__(self,excel):

        self.error_message = []  # 错误信息
        self.shift_info = {}     # 班次信息  代码和工时的对应
        self.month_attend_working_time = 0   # 当月出勤工时
        self.schedule_data = {}  # 排班数据
        self.excel = excel

        self.rest_data = []#假期数据 元组组成的列表。 20190528 15:59  暂存有逗号左半部分和有半部分，左半边包含-号的

    def check_and_parse(self):
        wb = xlrd.open_workbook(filename=self.excel.name, file_contents=self.excel.read())
        # sheet名称 数据校验
        for var in ['班次&假期说明','实际出勤表']:
            if var not in wb.sheet_names():
                self.error_message.append('缺少sheet，或sheet名称错误，必须包含：班次&假期说明，实际出勤表')
        
        for var in wb.sheet_names()[1:3]:
            if len(var) != len(var.strip()):
              self.error_message.append("请查看excel表格的sheet名称>> %s <<的两边是否有空格" % var)



        if self.error_message:
            return self.error_message

        # 班次数据校验 
        t1 = wb.sheet_by_name('班次&假期说明')
        if t1.cell(0, 0).value != '代码':
            self.error_message.append('班次代码不存在！')
        if t1.cell(0, 2).value != '工时':
            self.error_message.append('班次工时不存在！')
        for row in range(1,t1.nrows):
            try:
                # 班次信息  给班次shift_info表赋值
                self.shift_info[t1.cell(row,0).value] = float(t1.cell(row,2).value)
                # print(t1.cell(row,0).value ,t1.cell(row,2).value)
            except Exception as e:
                continue#如果在试图转换某一行时遇到了错误，则忽略这行  2019/6/13 add
                self.error_message.append('工时数据格式不正确！,请检查第 %d 行,错误信息 %s' % (row,str(e)))


        # 检查内容是否存在 以及 格式
        if t1.cell(0, 7).value != '当月出勤工時':
            self.error_message.append('当月出勤工時不存在！')

        try:
            self.month_attend_working_time = float(t1.cell(0,8).value)
        except:
            self.error_message.append('當月出勤工時格式不正确！')

        if self.error_message:
            return self.error_message

        # 排班数据校验 ，开始检查排班表
        t2 = wb.sheet_by_name('实际出勤表')
        # 校验数据是否完成 第11列排班数据之前，指定的几个位置是有值
        for row in range(3,t2.nrows):
            if not all([t2.cell(row, 6).value, t2.cell(row, 7).value, t2.cell(row, 8).value, t2.cell(row, 9).value, t2.cell(row, 10).value]):
                self.error_message.append('排班表第' + str(row+1) + '行数据不完整!,姓名,全/兼职,职级,岗位,平台 为必填！')

        """
        # 判断第4行，实际有值的地方，找出截止那一列,因为后面的都没有意义
        改变t2.ncols的值
        2019-6-6 add

        2019/6/14 add
        平台不同，组不同，工作时间长短不同
        """   
        # num = 0
        # for col in range(11,t2.ncols):
        #     if not t2.cell(3, col).value:#找到了为空值的，则赋值，没找到num的值一直不变
        #         num = col#在这个数字之前都是有值的
        #         break
        # #如果不加这几行，正常情况反而处理不了了   2019/6/13 add  修复BUG
        # if num == 0:
        #     pass
        # else:
        #     t2.ncols = num

        # 处理多个平台，上班周期有长有短，混杂在一起的情况  2019/6/17 add
        num = 0
        for row in range(3,t2.nrows):#从第3行开始
            for col in range(11,t2.ncols):#从第11列开始
                if len(t2.cell(row, col).value) > 0 and (col > num):
                    num = col
        t2.ncols = num + 1






        # 校验日期格式，从11列开始的每一列  2019/5/1  如果是这个格式，则正常，否则要修改
        for col in range(11,t2.ncols):
            try:
                xlrd.xldate.xldate_as_datetime(t2.cell(1, col).value,0)
            except:
                self.error_message.append('排班表第 %d 行，第 % d列 日期格式不正确 %s！'%(2,col,(t2.cell(1, col).value,0)))#  5-30 change
        #上面的错误不返回，会导致后面的给大字典赋值时，导致后台崩掉
        if self.error_message:
            return self.error_message


        #self.error_message.append("目前所有的班次:   "+' | '.join(self.shift_info.keys()))
        # 校验 实际出勤表中的 出勤代码 没有出现在 班次&假期说明 中的情况
        for row in range(3,t2.nrows):#处理序号为3的行
            for col in range(11,t2.ncols):#从第11列开始   

                """
                2019-5-28 15:30 各种拦截异常的情况，完成
                待定：Q-4的情况，这4个小时，需要时间来补的
                有几种情况
                    1 为空 或者  OFF
                    2 S
                    3 另外：直接以大写字母X开头的，比如XA 已拦截
                    4 这个字符串在  班次&假期说明  中

                        包含英文逗号的情况
                        3 字符串 'S-4,XA2,XE2' 
                            如果有中文逗号的情况，要实现转成英文
                            拆分成列表，第一个元素参考4的方式解决

                        不包含英文逗号的情况
                        4 S-4 或者 Q+4
                        5 ojefafa 最后一个else会处理
                            排班都是大写字母，小写字母会拦截  已拦截
                    SR2-8.XM8.0
                    需要处理的：
                        fwafaf  已拦截
                        S2, 
                    另外 格式虽然正确但是 不符合实际的情况，需要拦截：
                        R+3,XA1,XK1   当日有加班，但是又用各种假期来填充，合理 ，不需要拦截
                        R-3,XA5,XK1   当日请假，但是不能用太多的时间来填补    | 拦截成功

                    w,XOT2  报错 需要放在上面的 拦截函数中  5-30 change  已拦截
                    O,XOT2,XE2  不合理 5-30 change 已拦截

                """
                #假如中文逗号出现在 字符串中，则替换成英文逗号
                cellString = str(t2.cell(row, col).value).replace("，",",")
                
                # 1 如果是空或者OFF，则什么都不做
                if (cellString == "OFF") or (not len(cellString)):
                    pass
                elif cellString.startswith("X"):
                    self.error_message.append(',-第%d行，第%d列排班代码 %s 以假期代码X开头，不合理，请修改' % (row+1,col,t2.cell(row, col).value))

                #如果cellString 已经在班表中了，则不报error到前台
                elif cellString in self.shift_info:
                    pass

                elif "," in cellString:#字符串中有逗号的情况

                    self.rest_data.append((row, col, cellString))

                    listOfCellString = cellString.split(',',1)#从第一个逗号处分隔
                    leftOfString = listOfCellString[0]#例如：S-4 或者  S+4, 还有一种情况S, O
                    rightOfString = listOfCellString[1]#第一个逗号右半边字符串
                    if "-" in leftOfString:
                            
                        temp1 = leftOfString.split("-")[0]

                        if temp1 not in self.shift_info:
                            self.error_message.append(',-第%d行，第%d列排班代码  %s  没有出现在班次中' % (row+1,col,t2.cell(row, col).value))
                            
                        #有逗号，有-号，并且R也在self.shift_info的情况
                        #R-3,XA5,XK1   当日请假，但是不能用太多的时间来填补
                        else:
                            try:
                                float(leftOfString.split("-")[1])#R-O.5,XA1,XK1  #O.5不能转换成数字的情况
                            except:
                                self.error_message.append(',-第%d行，第%d列排班代码 %s 第一个-号的右边，第一个逗号的左边，子字符串不能转换成数字，请修改' % (row+1,col,t2.cell(row, col).value))
                  

                    elif "+" in leftOfString:
                        temp1 = leftOfString.split("+")[0]
                        if temp1 not in self.shift_info:
                            self.error_message.append(',+第%d行，第%d列排班代码  %s  没有出现在班次中' % (row+1,col,t2.cell(row, col).value))
                            
                        #这种情况不需要拦截  5-30 change
                        else:#R+3,XA1,XK1  当日有加班，但是又用各种假期来填充，不合理  
                            try:
                                float(leftOfString.split("+")[1])#R+O.5,XA1,XK1  #O.5不能转换成数字的情况
                            except:
                                self.error_message.append(',+第%d行，第%d列排班代码 %s 第一个+号的右边，第一个逗号的左边，子字符串不能转换成数字，请修改' % (row+1,col,t2.cell(row, col).value))
                            
                    else:#FFF,   fff,  O,XA8   需要判断逗号的左边有没有在
                        if leftOfString in self.shift_info:
                            #O,XOT2,XE2  不合理 5-30 change
                            #A,XOT2,XE2  不合理 5-30 change  这种情况放过 pass  
                            if cellString.startswith("O") and len(cellString.split(',')) >2:# O,XOT2,XE1
                                self.error_message.append('有逗号没有+-号的特殊情况：第%d行，第%d列排班代码  %s 是一种不合理的O班' % (row+1,col,t2.cell(row, col).value))
                            else:#O,XB8 请病假的情况 这是一种正常情况，放过
                                pass
                        else:
                            # OFF,XQ8  和  OFF,XBM8    2019/6/13 add  需要处理
                            if cellString.startswith("OFF"):
                                continue
                            else:
                                self.error_message.append('有逗号也没有+-号的特殊情况：第%d行，第%d列排班代码  %s  没有出现在班次中' % (row+1,col,t2.cell(row, col).value))

                else:#字符串中没有逗号的情况 A+3 a+3 A-3  A，另外一种情况SR2-8.XM8.0

                    #SR2-8.XM8.0  需要考虑的
                    if "-" in cellString:
                        temp2 = cellString.split("-")[0]
                        if temp2 not in self.shift_info:
                            self.error_message.append('-第%d行，第%d列排班代码  %s  没有出现在班次中' % 
                                (row+1,col,t2.cell(row, col).value))

                        """
                        2019/6/19 add
                        SR2-8.XM8.0  右半边 8.XM8.0 中
                        re.search('[a-z]', str)
                        字符串中包含英文则返回一个re对象 ，
                            不包含英文则返回None
                            包含英文，返回true，报错返回到前端
                        """
                        if re.search('[a-zA-Z]', cellString.split("-")[1]):
                            self.error_message.append('-第%d行，第%d列排班代码 %s ，请检查是否符合命名规则' % 
                                (row+1,col,t2.cell(row, col).value))
                            

                    #SR2+8.XM8.0  需要考虑的2019/6/19 add
                    elif "+" in cellString:
                        temp2 = cellString.split("+")[0]
                        if temp2 not in self.shift_info:
                            self.error_message.append('+第%d行，第%d列排班代码  %s  没有出现在班次中' % 
                                (row+1,col,t2.cell(row, col).value))


                        #SR2+8.XM8.0  右半边 8.XM8.0 中
                        if re.search('[a-zA-Z]', cellString.split("+")[1]):
                            self.error_message.append('-第%d行，第%d列排班代码 %s ，请检查是否符合命名规则' % 
                                (row+1,col,t2.cell(row, col).value))

                    else:#例如  XXX A T O   ，之前在班次列表中的情况都拦截了                        
                        self.error_message.append('****** 第%d行，第%d列排班代码  %s  没有出现在班次中' % 
                            (row+1,col,t2.cell(row, col).value))



        # 数据部分从 表格 到 大字典 schedule_data
        for row in range(3,t2.nrows):
            # row =3 col = 6,7,8,9,10  姓名   属性  职级  岗位  平台

            name = t2.cell(row,6).value

            # 2019/6/17 add 解决同一个名字在表中出现了2次的情况的问题
            if name in  self.schedule_data:
                # print(name,row)
                name = name + str(t2.cell(row,10).value)

            self.schedule_data[name] = {}
            self.schedule_data[name]['attribute'] = t2.cell(row,7).value
            self.schedule_data[name]['rank'] = t2.cell(row,8).value
            self.schedule_data[name]['position'] = t2.cell(row,9).value
            self.schedule_data[name]['platform'] = t2.cell(row,10).value
            self.schedule_data[name]['schedule'] = {}

            #之后在函数 def transform(self, temp):中给total改变value的值
            self.schedule_data[name]['total'] = 0
            for col in range(11,t2.ncols):
                #用for loop 给 日期工时 小字典赋值
                #日期和工时代号组成的元组
                #{datetime.datetime(2019, 5, 18, 0, 0): 'OFF',  这是其中的一个元素}

                # 5-30 change  从这里返回到之前，查BUG
                self.schedule_data[name]['schedule'][xlrd.xldate.xldate_as_datetime(t2.cell(1, col).value,0)] = t2.cell(row, col).value


                


        return self.error_message


    """
    因为之前的各种格式不正确的情况，已经在 函数 def check_and_parse(self):中拦截
    def transform(self, temp)  把字符串转换成值
    字符串可能的形式
    
    OFF 和  长度为0的情况

    
    
    在正常的排班 代码 中的情况
    A
    S1
    #不在排班代码中的情况
    if 有逗号分隔的情况:        
        异常情况   试试中间为中文逗号的情况   已拦截
        异常情况   R+3,XA+1,XK+1        已拦截
        A-3,XA+3
    else:
        S1-1
    

    """

    #如果 transform 没有返回值，那么调用它的代码附近会报错
    def transform(self, temp):
        #假如中文逗号出现在 字符串中，则替换成英文逗号
        temp = temp.replace("，",",")
        #str类型
        #print(type(temp))

        # 班次为空或OFF，工时为零
        if not len(str(temp)) or temp == 'OFF':
            return 0
        elif temp in self.shift_info:#temp在排班代码中的情况，正常情况,返回了对应的value
            return self.shift_info[temp]
        else:

            #有逗号的情况 A-3,XA+3  R-3,XA+1   R-3,XA+1,XK+1 ，R+3,XA+1,XK+1 
            #还有一种情况O,XOT2 和A,XOT2，当天没有排班，来加了2个小时
            if ',' in temp:#从第一个逗号出分隔,是想要的效果 R-3 和 XA+1,XK+1,即2个字符串
                # ----------需要考虑的，异常情况   R+3,XA+1,XK+1
                if '+' in temp.split(',',1)[0]:#第一个逗号左半边字符串 R+3

                    # R+3,XA+1,XK+1 不需拦截，需要计算工时，需要有返回值  5-30 change
                    temp0 = temp.split(',',1)[0].split("+")[0]# ['R','3']  "R"
                    temp1 = temp.split(',',1)[0].split("+")[1]# ['R','3']   "3"
                    return float(self.shift_info[temp0]) + float(temp1)# 正常工时和加班，都要计算到总工时里

                # ----------需要考虑的，异常情况   R-3,XA+1,XK+1
                elif '-' in temp.split(',',1)[0]:#第一个逗号左半边字符串 R-3
                    # R+3,XA+1,XK+1 不需拦截，需要计算工时，需要有返回值 5-30 change
                    temp0 = temp.split(',',1)[0].split("-")[0]# ['R','3']  "R"
                    temp1 = temp.split(',',1)[0].split("-")[1]# ['R','3']   "3"
                    return float(self.shift_info[temp0]) - float(temp1)# 正常工时和请假，都要计算到总工时里

                # 有逗号的情况，没有+ - 号， 取逗号左侧的值
                # 如果是 O 开头 ，需要去逗号右侧的数字，这就是O班的特殊性
                else:
                    #O,XOT2   O, 
                    if temp.startswith("O") and (not temp.split(",")[-1]):
                        return float(self.shift_info[temp.split(',')[0]]) + float(temp[-1])#O,XOT2

                    #2019/6/13 add
                    elif temp.startswith("OFF"):#OFF,XQ8  和 OFF,XBM8  特殊情况
                        return 0
                    else:#O,   O班是0小时
                        return float(self.shift_info[temp.split(',')[0]])#SR2,XA2 以及O，
                """
                5-30 change
                    O,  A2, 合理
                    O,XOT2  A,XOT2  合理
                    
                    A,XOT2,XE2  合理

                    w,XOT2  报错 需要放在上面的 拦截函数中  -----------------已拦截
                    O,XOT2,XE2  不合理-------------------  已拦截

                    假如第一个逗号的左边没有存在于 self.shift_info 中
                        
                    假如存在：

                """

            #没有逗号的情况A-3
            else:
                if '-' in temp:
                    return self.shift_info[temp.split('-')[0]] - float(temp.split('-')[1])
                if '+' in temp:
                    return self.shift_info[temp.split('+')[0]] + float(temp.split('+')[1])


    def generate_excel(self):
        workbook = xlsxwriter.Workbook('工时.xlsx')
        format_date = workbook.add_format({'num_format': 'mm/dd','align': 'center'})
        format = workbook.add_format({'align': 'center'})
        # Sheet1
        sheet1 = workbook.add_worksheet('Output_Rawdata')
        for row,people in enumerate(self.schedule_data):
            # 表头信息
            if not row:#在第0行写入，其余行就不写了
                sheet1.merge_range(0, 0, 1, 0, '序号',format)
                sheet1.merge_range(0, 1, 1, 1, '姓名',format)
                sheet1.merge_range(0, 2, 1, 2, '属性',format)
                sheet1.merge_range(0, 3, 1, 3, '职级',format)
                sheet1.merge_range(0, 4, 1, 4, '岗位',format)
                sheet1.merge_range(0, 5, 1, 5, '平台',format)
                sheet1.merge_range(0,len(self.schedule_data[people]['schedule'])+6,1,len(self.schedule_data[people]['schedule'])+6,'Total',format)

            # 数据
            #sheet1.write(某行，某列 ，值， 格式format) row+2出现过多次
            sheet1.write(row+2,0,row+1,format)
            sheet1.write(row+2,1,people,format)
            sheet1.write(row+2,2,self.schedule_data[people]['attribute'],format)
            sheet1.write(row+2,3,self.schedule_data[people]['rank'],format)
            sheet1.write(row+2,4,self.schedule_data[people]['position'],format)
            sheet1.write(row+2,5,self.schedule_data[people]['platform'],format)

            for col,var in enumerate(self.schedule_data[people]['schedule']):
                #下面2行在循环过程中被反复写入，但是都是同样的内容
                sheet1.write_datetime(0,col+6,var,format_date)#固定的行，写入时间  2019/5/1 var
                sheet1.write(1,col+6,get_week_day(var),format)#固定的行，写入时间 例如 Wed, Sun

                #print(col,var,self.schedule_data[people]['schedule'][var])
                # 0 2019-05-01 00:00:00
                # 1 2019-05-02 00:00:00 OFF
                # col    key是var时间      value = self.schedule_data[people]['schedule'][var] 是OFF


                #在类内部调用 类函数的方式 | 准备temp的值 把字符串转换成值
                temp = self.transform(str(self.schedule_data[people]['schedule'][var]))
                sheet1.write(row+2,col+6,temp,format)


                #每循环一次累加一次  注意：如果transform 该返回结果的地方没有返回，则下行报错
                self.schedule_data[people]['total'] = self.schedule_data[people]['total'] + temp

            # 上面这个循环结束之后，把total的值写入末尾
            sheet1.write(row+2,len(self.schedule_data[people]['schedule'])+6,self.schedule_data[people]['total'],format)




        # Sheet2，Function
        sheet2 = workbook.add_worksheet('Output_Function')
        sheet2.write(0,0,'Role',format)
        sheet2.write(0,1,'Function',format)
        sheet2.write(0,2,'Total HC',format)
        sheet2.write(0,3,'working time',format)

        #遍历字典的没一个元素，每个元素都是字典，用'rank' 整个key来取 value，并去重之后取得列表
        # rank_list = ["agent","mgt"]  类似这种
        rank_list = list(set([self.schedule_data[var]['rank'] for var in self.schedule_data]))

        function_dict = {}
        # 获取function
        for rank in rank_list:
            function_dict[rank] = {}#{}
            for var in self.schedule_data:
                if self.schedule_data[var]['platform'] not in function_dict[rank] and self.schedule_data[var]['rank'] == rank:
                    function_dict[rank][self.schedule_data[var]['platform']] = 0

        # print(function_dict)
        # {'Mgt': {'OS': 0}, 'Agent': {'JD': 0}}


        # 计算total
        for rank in function_dict:
            for function in function_dict[rank]:
                for var in self.schedule_data:
                    if self.schedule_data[var]['rank'] == rank and self.schedule_data[var]['platform'] == function:
                        function_dict[rank][function] = function_dict[rank][function] + self.schedule_data[var]['total']

        # print(function_dict)
        # {'Mgt': {'OS': 32.0}, 'Agent': {'JD': 24.0}}   
        # 写入Excel
        temp_row = 1
        for rank in function_dict:
            if len(function_dict[rank]) == 1:
                sheet2.write(temp_row, 0,rank,format)
            else:
                sheet2.merge_range(temp_row, 0, len(function_dict[rank])+temp_row-1, 0, rank,format)

            for n,function in enumerate(function_dict[rank]):
                sheet2.write(n+temp_row,1,function,format)
                sheet2.write(n+temp_row,2,
                    round(float(function_dict[rank][function])/float(self.month_attend_working_time), 1), 
                    format)
                sheet2.write(n+temp_row,3,function_dict[rank][function],format)
            temp_row = temp_row + len(function_dict[rank])





        # Sheet3，Level
        sheet3 = workbook.add_worksheet('Output_Position')
        sheet3.write(0,0,'Role',format)
        sheet3.write(0,1,'Level',format)
        sheet3.write(0,2,'Total HC',format)
        sheet3.write(0,3,'working time',format)

        level_dict = {}
        # 获取level
        for rank in rank_list:
            level_dict[rank] = {}
            for var in self.schedule_data:
                if self.schedule_data[var]['position'] not in level_dict[rank] and self.schedule_data[var]['rank'] == rank:
                    level_dict[rank][self.schedule_data[var]['position']] = 0

        # 计算 total
        for rank in level_dict:
            for level in level_dict[rank]:
                for var in self.schedule_data:
                    if self.schedule_data[var]['rank'] == rank and self.schedule_data[var]['position'] == level:
                        level_dict[rank][level] = level_dict[rank][level] + self.schedule_data[var]['total']
                
        # 写入Excel
        temp_row = 1
        for rank in level_dict:
            if len(level_dict[rank]) == 1:
                sheet3.write(temp_row, 0,rank,format)
            else:
                sheet3.merge_range(temp_row, 0, len(level_dict[rank])+temp_row-1, 0, rank,format)

            for n,level in enumerate(level_dict[rank]):
                sheet3.write(n+temp_row,1,level,format)
                sheet3.write(n+temp_row,2,
                    round(float(level_dict[rank][level])/float(self.month_attend_working_time), 1),
                    format)
                sheet3.write(n+temp_row,3,level_dict[rank][level],format)
            temp_row = temp_row + len(level_dict[rank])


        # Sheet4
        sheet4 = workbook.add_worksheet('Output_Restdata')

        # 写入表头部分
        for row,people in enumerate(self.schedule_data):
            if not row:#在第0行写入，其余行就不写了
                sheet4.merge_range(0, 0, 1, 0, '序号',format)
                sheet4.merge_range(0, 1, 1, 1, '姓名',format)
                sheet4.merge_range(0, 2, 1, 2, '属性',format)
                sheet4.merge_range(0, 3, 1, 3, '职级',format)
                sheet4.merge_range(0, 4, 1, 4, '岗位',format)
                sheet4.merge_range(0, 5, 1, 5, '平台',format)
                sheet4.merge_range(0,len(self.schedule_data[people]['schedule'])+6,1,len(self.schedule_data[people]['schedule'])+6,'Total',format)

            sheet4.write(row+2,0,row+1,format)
            sheet4.write(row+2,1,people,format)
            sheet4.write(row+2,2,self.schedule_data[people]['attribute'],format)
            sheet4.write(row+2,3,self.schedule_data[people]['rank'],format)
            sheet4.write(row+2,4,self.schedule_data[people]['position'],format)
            sheet4.write(row+2,5,self.schedule_data[people]['platform'],format)

            for col,var in enumerate(self.schedule_data[people]['schedule']):
                #下面2行在循环过程中被反复写入，但是都是同样的内容
                sheet4.write_datetime(0,col+6,var,format_date)#写入时间  2019/5/1
                sheet4.write(1,col+6,get_week_day(var),format)#写入时间 例如 Wed, Sun

        #假期 写入表格
        for rest in self.rest_data:


            #R-7,XA5,XK1  ->   XA,XK    ->     A,K
            restString = rest[2].split(',',1)[1]#用逗号拆分，XA5,XK1

            # restList = re.findall("\D+", restString)#['XA+', ',XK+']
            # rList = [s.strip(',+').strip("X") for s in restList]#['A','K'] (5-29 change)

            restList = restString.split(",")# ['XA5', 'XK1'] A5 K1必须体现在假期表格中
            rList = [s.strip("X") for s in restList]#['A5','K1'] (5-30 change) done
            #print(rest[0]-1, rest[1]-5, "".join(rList))
            sheet4.write(rest[0]-1, rest[1]-5, "".join(rList), format)


        workbook.close()