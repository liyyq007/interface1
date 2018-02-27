# encoding=utf8
#完成递增

from xlutils.copy import copy
import xlrd
import sys

reload(sys)
sys.setdefaultencoding('utf8')

x = raw_input(u'增长行数:')
x=int(x)+1
xl_data = xlrd.open_workbook('D:\\test\\API4.xls')


table = xl_data.sheet_by_name(u'Sheet1')
rows = table.ncols
print '一共多少列:',rows
w_xls = copy(xl_data)
w_xls.encoding='utf8'
sheet_write = w_xls.get_sheet(0)
for i in range(1,x):
    print '运行条数:',i
    #begin num值
    beginnum=0
    num=beginnum+i+000
    a=str(num)
    b=str(beginnum+i+100)
    print b
    c=b
    params = u'{"source":3,"name":"客户'+a+'","primaryMobile":"189762300'+b+'"}'
    print params
    sheet_write.write(i,3, (params))

w_xls.save('.\\API4.xls')