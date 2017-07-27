# -*- coding: UTF-8 -*-
import xlrd,logging,urllib,urllib2,json,sys
import requests
import time
from flask import Flask, request, render_template

#定义系统输出编码
reload(sys)
sys.setdefaultencoding('utf-8')

#定义日志输出
logging.basicConfig(level=logging.DEBUG,
                format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S')#,
#                 filename='myapp.log',
#                 filemode='w')

#定义一个StreamHandler，将INFO级别或更高的日志信息打印到标准错误，并将其添加到当前的日志处理对象#
logger = logging.getLogger("fib")#fib
logger.setLevel(logging.DEBUG)#fib
#
console = logging.StreamHandler()
# console.setLevel(logging.INFO)
formatter = logging.Formatter('[%(asctime)s] %(name)s:%(levelname)s: %(message)s')
console.setFormatter(formatter)
#
logger.addHandler(console)#fib

# logging.getLogger('').addHandler(console)
logger.removeHandler(console)


#处理excel表格
xl_data = xlrd.open_workbook('D:\\test\\API.xlsx')
logging.debug("打开%s excel表格成功"%xl_data)
table = xl_data.sheet_by_name(u'Sheet1')
logging.debug("打开%s表成功"%table)
nrows = table.nrows#行数
logging.debug("表中有%s行"%nrows)
ncols = table.ncols#列数
logging.debug("表中有%s列"%ncols)

rightside = []
result_1 = []
rst_data = []
leftside = []
row=[]
passed = 0
fail = 0
noresult = 0
for i in range(1,nrows):#遍历
    cell_A3 =table.row_values(i)
    id= cell_A3[0]
    title= cell_A3[1]
    url= cell_A3[2]
    params= cell_A3[3]
    leixing = cell_A3[4]
    result =cell_A3[5]
    send_headers =cell_A3[6]
    Remarks = cell_A3[7]

    paramsdict=json.loads(params)
    params_turn =urllib.urlencode(paramsdict)  #参数化处理
    send_headers_dict=json.loads("{"+send_headers+"}")
    url2 = urllib2.Request(url=url,data=params_turn,headers=send_headers_dict)
    # url2.add_header(send_headers.send_headerssplit(":")[0],send_headers.split(":")[1])
    response = urllib2.urlopen(url2)
    apicontent = response.read()
    apicontent = json.dumps(apicontent).decode("unicode-escape")

    lf_data = {
                'code': id,
                'title': unicode(title).encode("utf-8")
            }
    leftside.append(lf_data)

    rs_data = {
                "fssj": u"测试数据",
                "csbt": unicode(title).encode("utf-8"),
                "fsfs": leixing,
                "alms": Remarks,
                "fsdz": url,
                "fscs": str(params).decode("unicode-escape").encode("utf-8").replace("u\'","\'"),
                'testid': id,
                "apicontent": apicontent
            }
    rightside.append(rs_data)

    try:
        if cmp(apicontent, result) == 0:
            passed += 1
            row.append("pass")
        else:
            fail += 1
            row.append("fail")
    except Exception:
        noresult += 1
        row.append("no except result")

    rs_data = []
    for y in row:
        rs_data.append(y)
    result_1.append(rs_data)


    for a, b in zip(rightside, result_1):
        data = {
            "sendData": a,
            "dealData": b,
            "result": b[len(b) - 1]
        }
    rst_data.append(data)

logging.debug('result:%s'%rs_data)

app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html',leftside=leftside,rst_data=rst_data, pppp=passed, ffff=fail, noresult=noresult)

if __name__ == '__main__':
    app.run()
