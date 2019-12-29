# coding=UTF-8
import datetime
import xlrd, xlwt, xlutils

def pay_process(flag,lists,f_date):
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
        module_list1.append(zh_data1)
    return module_list1