import xlrd
import xlwt
import numpy
import matplotlib.dates
from matplotlib import pylab
from pylab import *
from datetime import datetime

workbook = xlrd.open_workbook('lab2.xlsx')
worksheet = workbook.sheet_by_index(0)
dtype = ["desktop","mobile","tablet"]
utype = ["New Visitor","Returning Visitor"]

def rows():
    rows = []
    for rownum in range(1, worksheet.nrows):
        row = worksheet.row_values(rownum)
        rows.append(row)
    return rows
	
r=rows()

def parse_date(d):
    dd = datetime.strptime(d,'%Y%m%d%H')
    return dd

def date_str(d):
    dd = datetime.strftime(d,'%d %b %Y %H:00')
    return dd

def date_list():
    d_lst = []
    for rownum in range(1, worksheet.nrows):
        row = worksheet.row_values(rownum)
        dd = parse_date(row[0])
        d_lst.append(dd)
    return d_lst

def stat(lst):
    ev = numpy.mean(lst)
    v = numpy.var(lst)
    std = numpy.std(lst)
    return ev,v,std

def plot_distrib_dev(dev1,dev2,dev3,name):
    d_lst = date_list()
    date_set = set(d_lst)
    d_list = list(date_set)
    d_list.sort()
    date = []
    for i in range(len(d_list)):
        date.append(d_list[i])
    plot(date,dev1,label=u'desktop')
    plot(date, dev2, label=u'mobile')
    plot(date, dev3, label=u'tablet')
    xlabel('Date')
    title(name)
    grid(True)
    legend()
    show()

def dtype_matrix():
    desktop_list = []
    mobile_list = []
    tablet_list = []
    for i in range(len(r)):
        desktop = mobile = tablet = 0
        r1 = r[i]
        desktop = desktop + r1.count(dtype[0])
        desktop_list.append(desktop)
        mobile = mobile + r1.count(dtype[1])
        mobile_list.append(mobile)
        tablet = tablet + r1.count(dtype[2])
        tablet_list.append(tablet)
    return desktop_list,mobile_list,tablet_list
	
M = dtype_matrix()

def date_index():
    d_lst = date_list()
    date_set = set(d_lst)
    d_list = list(date_set)
    d_list.sort()
    date_index = []
    for i in range(len(d_list)):
        temp = d_lst.count(d_list[i])
        date_index.append(temp)
    return date_index,d_list

def dtype_num_matrix():
    M = dtype_matrix()
    M0 = M[0]
    M1 = M[1]
    M2 = M[2]
    index = date_index()
    index = index[0]
    desk_l = []
    mob_l = []
    tabl_l = []
    temp = 0
    for i in range(len(index)):
        desktop = mobile = tablet = 0
        desktop = sum(M0[temp:index[i]+temp])
        desk_l.append(desktop)
        mobile = sum(M1[temp:index[i]+temp])
        mob_l.append(mobile)
        tablet = sum(M2[temp:index[i]+temp])
        tabl_l.append(tablet)
        temp = temp+index[i]
    return desk_l,mob_l,tabl_l
	
DT = dtype_num_matrix()
plot_distrib_dev(DT[0],DT[1],DT[2],"Device distribution")

def normalization(dev):
    norm_dtype = []
    div = dev[0]
    for i in range(len(dev)):
        dev[i]=dev[i]/div
    return dev
	
DT_p = [normalization(DT[0]),normalization(DT[1]),normalization(DT[2])]

plot_distrib_dev(DT_p[0],DT_p[1],DT_p[2],"Device distribution after normalization")

def loadtime_dev_distrib():
    M = dtype_matrix()
    M0 = M[0]
    M1 = M[1]
    M2 = M[2]
    d_load_time = []
    m_load_time = []
    t_load_time = []
    for i in range (len(r)):
        if M0[i]==1:
            d_load_time.append(r[i][8])
        elif M1[i]==1:
            m_load_time.append(r[i][8])
        elif M2[i]==1:
            t_load_time.append(r[i][8])
    return d_load_time,m_load_time,t_load_time
	
DLT = loadtime_dev_distrib()

def dlt_deviation(n):
    stat_d = stat(DLT[n])
    dev_date = []
    temp = 0
    for i in range(len(r)):
        if M[n][i]==1:
            if r[i][8]>stat_d[0]+2*stat_d[1] or r[i][8]<stat_d[0]-2*stat_d[1]:
                r_d = parse_date(r[i][0])
                dev_date.append(date_str(r_d))
                temp = temp+1
    return dev_date
	
dlt_deviation(0)

def compare_dlt_deviation(dev1,dev2,dev3):
    dlt_date = []
    dlt_date_s = []
    dlt_date = dev1+dev2+dev3
    dlt_date_s = set(dlt_date)
    dlt_date_s = list(dlt_date_s)
    for i in range(len(dlt_date_s)):
        dev_count = dlt_date.count(dlt_date_s[i])
        if dev_count>1:
            print ("{}:{}".format(dlt_date_s[i],dev_count))
			
compare_dlt_deviation(dlt_deviation(0),dlt_deviation(1),dlt_deviation(2))

def write_dlt(dev1,dev2,dev3):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('DLT1')
    for i in range(len(dev1)):
        ws.write(i,0,dev1[i])
        ws.write(i, 1, dev2[i])
        ws.write(i, 2, dev3[i])
        i=i+1
    wb.save('d:/lab2/dlt.xls')
	
write_dlt(DT_p[0],DT_p[1],DT_p[2])

def dt_deviation(n):
    stat_d = stat(DT[n])
    temp = 0
    for i in range(len(r)):
        if M[n][i] == 1:
            if r[i][8] > stat_d[0] + 2 * stat_d[1] or r[i][8] < stat_d[0] - 2 * stat_d[1]:
                r_d = parse_date(r[i][0])
                print(date_str(r_d))
                temp = temp + 1
    print("Number of deviation:{}".format(temp))
	
def t_deviation(n):
    stat_d = stat(DT[n])
    temp = 0
    d_lst = date_list()
    date_set = set(d_lst)
    d_list = list(date_set)
    d_list.sort()
    for i in range(len(d_list)):
        if DT[n][i]>stat_d[0] + stat_d[1] or DT[n][i] < stat_d[0]- stat_d[1]:
            print (d_list[i])
            temp=temp+1
    print("Number of deviation:{}".format(temp))
	
t_deviation(2)

#peak
def find_peak(dev1):
    is_peak = []
    for i in range (len(dev1)):
        is_peak.append(0)
    i=0
    for i in range(len(dev1)-10):
        mx = max(dev1[i:i+10])
        mn = min(dev1[i:i + 10])
        mxi = dev1[i:i+10].index(mx)
        mni = dev1[i:i+10].index(mn)
        is_peak[mxi+i]=1
        is_peak[mni+i]=-1
        i=i+10
    return is_peak
	
plot_distrib_dev(find_peak(DT_p[0]),find_peak(DT_p[1]),find_peak(DT_p[2]),"Device distribution peak")

def compare_peak(dev1,dev2,dev3):
    sum = []
    for i in range(len(dev1)):
        if (dev1[i]==1 and dev2[i]==-1 and dev3[i]==-1) or (dev2[i]==1 and dev1[i]==-1 and dev3[i]==-1) or (dev3[i]==1 and dev2[i]==-1 and dev1[i]==-1):
            sum.append(1)
        elif (dev1[i]==-1 and dev2[i]==1 and dev3[i]==1) or (dev2[i]==-1 and (dev1[i]==1 and dev3[i]==1)) or (dev3[i]==-1 and (dev2[i]==1 and dev1[i]==1)):
            sum.append(1)
        else:
            sum.append(0)
    return sum
	
cp = compare_peak(find_peak(DT_p[0]),find_peak(DT_p[1]),find_peak(DT_p[2]))
plot_distrib_dev(cp,cp,cp,"Device distribution: compare peaks")
print (stat(cp))
