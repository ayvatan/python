# coding=UTF-8
import datetime
import xlrd, xlwt, xlutils
import time
import sys
#from func import *



logs = "log.txt"
def log_wr(info):
    info =_get_log_()+'   '+info+'\n'
    f = open(logs,'a')
    f.write(info)
    f.close()
def  _get_log_():
    import datetime
    now = datetime.datetime.now()
    log_date = now.strftime('%Y-%m-%d %H:%M:%S')
    return log_date

######################################
##数据校验###
from progress.bar import Bar
def times1(str):
   type = sys.getfilesystemencoding()
   Mystring = str
   print Mystring.decode('utf-8').encode(type)
   bar = Bar('', max=100, fill='#', suffix='%(percent)d%%')
   for i in range(100):
       time.sleep(0.1)
       bar.next()
   bar.finish()


def times():
    type = sys.getfilesystemencoding()
    Mystring = '数据校验，请稍等！！！'
    print Mystring.decode('utf-8').encode(type)
    time.sleep(10)
######################################


# coding=UTF-8
import datetime
import xlrd, xlwt, xlutils
# def  title_def():
#     #函数处理excel列头，确定计算需要的列
#      if
#          pass
#      elif
#          pass
#      return title_num[]
def invest(amount, rate, time):
#    print('princical amount: {}'.format(amount))
    rate1 = rate/360
    for t in range(1, time + 1):
      amount = amount * (rate1 + 1)
  #  print('year {}: {}'.format(t, amount))
    return amount


def pay_process(flag,lists,f_date):
    #此函数计算正常利率以及逾期利率等，主要适配利息计算方式
    a = int(f_date[0])
    b = int(f_date[1])
    c = int(f_date[2])
    d = int(f_date[3])
    s_date = datetime.datetime.strptime(str(lists[a]), '%Y%m%d')
    e_date = datetime.datetime.strptime(str(lists[b]), '%Y%m%d')
    if flag == 1:
       s_date = datetime.datetime.strptime(str(lists[a]),'%Y%m%d')
       e_date = datetime.datetime.strptime(str(lists[b]),'%Y%m%d')
       paydate = (s_date - e_date).days
       rate = float(lists[c])*float(lists[d])/360*paydate
       return rate
    elif flag == 2:
       s_date = datetime.datetime.strptime(str(lists[a]), '%Y%m%d')
       e_date = datetime.datetime.strptime(str(lists[b]), '%Y%m%d')
       t_date = datetime.datetime.now()
       t_d = t_date.strftime('%Y%m%d')
       paydate = (t_date - e_date).days
       percent = float(lists[d])+float(lists[d])*0.4
       rate = lists[c] * percent/360*paydate
       return rate
    elif flag == 3:
        s_date = datetime.datetime.strptime(str(lists[a]), '%Y%m%d')
        e_date = datetime.datetime.strptime(str(lists[b]), '%Y%m%d')
        t_date = datetime.datetime.now()
        paydate = (t_date - e_date).days
        return invest(lists[c],float(lists[d]),paydate)
    elif flag == 4:
        s_date = datetime.datetime.strptime(str(lists[a]), '%Y%m%d')
        e_date = datetime.datetime.strptime(str(lists[b]), '%Y%m%d')
        t_date = datetime.datetime.now()
        paydate = (t_date - e_date).days
        return invest(float(f_date[4]), float(lists[d]), paydate)
    elif flag == 5:
        s_date = datetime.datetime.strptime(str(lists[a]), '%Y%m%d')
        e_date = datetime.datetime.strptime(str(lists[b]), '%Y%m%d')
        t_date = datetime.datetime.now()
        return (t_date - e_date).days


def sheet_headinfor(sheet_num,ncol):
    i = 0
    head_name_list1 = []
    for i in range(ncol):
        head_name_list1.append(sheet_num.cell_value(0, i))
    return head_name_list1


def data_process(nrows,ncol,sheetname,f_date,sheet_nums):
    module_list1 = []
    se = int(f_date[sheet_nums][3])
    for d_s in range(nrows):
        zh_data1 = []
        for d_1 in range(ncol):
            if sheetname.cell_type(d_s, d_1) == 2 and d_1 <> se:
                zh_data1.append(long(sheetname.cell_value(d_s, d_1)))
            elif sheetname.cell_type(d_s, d_1) == 3 and d_1 <> se:
                zh_data1.append(long(sheetname.cell_value(d_s, d_1)))
            elif sheetname.cell_type(d_s, d_1) == 0:
                zh_data1.append("数据为空")
            elif sheetname.cell_type(d_s, d_1) == 1:
                zh_data1.append(sheetname.cell_value(d_s, d_1).encode('utf-8'))
            else:
                zh_data1.append(sheetname.cell_value(d_s, d_1))
        if d_s > 0:
           #  pass
              zh_data1.append(pay_process(1, zh_data1,f_date[sheet_nums]))
              zh_data1.append(pay_process(2, zh_data1,f_date[sheet_nums]))
              zh_data1.append(pay_process(3, zh_data1, f_date[sheet_nums]))
              zh_data1.append(pay_process(4, zh_data1, f_date[sheet_nums]))
              zh_data1.append(pay_process(5, zh_data1, f_date[sheet_nums]))
        module_list1.append(zh_data1)
    return module_list1

##处理表中字段，定位计算使用的字段
def target_list(head_list):
    head_target=[]
    dict = {'1':1,'2':2,'3':3,'4':4}
    head_list_num = len(head_list)
    for i in range(0,head_list_num):
        if head_list[i].encode('utf-8') == "发放日期":
           dict['1']=i
        elif head_list[i].encode('utf-8') == "到期日期":
           dict['2']=i
        elif head_list[i].encode('utf-8') == "贷款金额":
           dict['3'] = i
        elif head_list[i].encode('utf-8') == "利率":
           dict['4']=i
    for s in range(1,5):
        head_target.append(dict[str(s)])
    head_target.append(head_list_num)
    return  head_target





#print "test"
##读取文件以及对文件的基础信息进行提取
try:
   data = xlrd.open_workbook('data.xls')
except:
    log_wr("需要处理的文件似乎不存在！")
    quit()
times1('正在检查文件状态')
times1('开始处理文件，计算开始')
##提取sheetl_list
table = data.sheet_names()
sheet_num = 0
#f_datas = [['3','7','6','8','12'],['5','6','4','7','12'],['4','5','3','9','12']]
sheet_nums = 0
###定义sheet名称
list_sheet = []
######处理表头数据
head = 0
listdd = []
dds = []
for sheet_num in table:
    sheetname = data.sheet_by_name(sheet_num)
    ds = sheetname.ncols
    head = []
    for i in range(0,ds):
        head.append(sheetname.cell_value(0,i))
   # target_list(head)
    dds = target_list(head)
    listdd.append(dds)
######
for sheet_num in table:
    sheetname = data.sheet_by_name(sheet_num)
    nrow = sheetname.nrows
    ncol = sheetname.ncols
 #   list222 = func.sheet_headinfor(sheetname,ncol)
    list_sheet.append(data_process(nrow,ncol,sheetname,listdd,sheet_nums))
    sheet_nums = sheet_nums + 1
list_sheet[0][0].append("正常利息")
list_sheet[0][0].append("逾期罚息")
list_sheet[1][0].append("正常利息")
list_sheet[1][0].append("逾期罚息")
list_sheet[2][0].append("正常利息")
list_sheet[2][0].append("逾期罚息")
list_sheet[0][0].append("本金复利")
list_sheet[0][0].append("利息复利")
list_sheet[1][0].append("本金复利")
list_sheet[1][0].append("利息复利")
list_sheet[2][0].append("本金复利")
list_sheet[2][0].append("利息复利")
list_sheet[0][0].append("逾期天数")
list_sheet[1][0].append("逾期天数")
list_sheet[2][0].append("逾期天数")


workbook = xlwt.Workbook(encoding = 'utf-8')
# 创建一个worksheet
#sheetname = data.sheet_by_name(sheet_num)
sumd = 0
for snums in table:
    sheetname = data.sheet_by_name(snums)
    worksheet = workbook.add_sheet(snums)
    list_lens = len(list_sheet[sumd])
    li2 = len(list_sheet[sumd][0])
    i = 0
    s = 1
    for i in range(list_lens):
     for s in range(li2):
        worksheet.write(i,s,label = list_sheet[sumd][i][s])
    sumd = sumd + 1
workbook.save('data_1.xls')



