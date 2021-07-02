import requests
import xlrd
from xlutils.copy import copy
from time import sleep, time
import sys
from datetime import datetime


def get_student_information(id_number, cookie):
    url = 'http://121.196.193.218/sxjgpt/student.do?list'
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Connection': 'keep-alive',
        'Content-Length': '113',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': cookie,
        'Host': '121.196.193.218',
        'Origin': 'http://121.196.193.218',
        'Referer': 'http://121.196.193.218/sxjgpt/student.do?main&menuId=402881e45aa6c5ae015aa6ca05cc0001',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36 Edg/91.0.864.59',
        'X-Requested-With': 'XMLHttpRequest',
    }
    data = {
        'qCityCode': '3306',
        'qCountyCode': '',
        'qSchoolCode': '',
        'idnum': id_number,
        'stuName': '',
        'stuid': '',
        'phone': '',
        'traincar': '',
        'page': '1',
        'rows': '10',
    }
    # 获取请求的响应
    r = requests.post(url, data=data, headers=headers)
    response = r.json()
    if response['total'] == 0:
        applydate = '未报名'
        insName = '无'
    else:
        applydate = response['rows'][0]['applydate']
        insName = response['rows'][0]['insName']
    return applydate, insName


def get_training_record(id_number, cookie):
    url = 'http://121.196.193.218/sxjgpt/stagetrainningtime.do?list'
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Connection': 'keep-alive',
        'Content-Length': '156',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': cookie,
        'Host': '121.196.193.218',
        'Origin': 'http://121.196.193.218',
        'Referer': 'http://121.196.193.218/sxjgpt/stagetrainningtime.do?main&menuId=402881e45ac68366015ac6996e360001',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36 Edg/91.0.864.59',
        'X-Requested-With': 'XMLHttpRequest',
    }
    data = {
        'qCityCode': '3306',
        'qCountyCode': '',
        'qSchoolCode': '',
        'subject': '',
        'auditstate': '-1',
        'applybegin': '',
        'applyend': '',
        'stuname': '',
        'idnum': id_number,
        'auditbegin': '',
        'auditend': '',
        'page': '1',
        'rows': '50',
    }
    # 获取请求的响应
    r = requests.post(url, data=data, headers=headers)
    response = r.json()
    return response


def information_processing(response):
    if response['total'] == 0:
        record = '无报审记录'
    else:
        record_model = '科目{0}报审时间：{1}\n'
        record = ''
        for i in range(response['total']):
            record += record_model.format(response['rows'][i]['subject'], response['rows'][i]['auditdate'])
    return record.rstrip()


def get_id_numbers_from_workbook(file_name):
    """从一个workbook中获取身份证信息"""    
    # 打开xlsx或者xls表
    workbook = xlrd.open_workbook(file_name)
    sheet = workbook.sheet_by_index(0)    
    id_numbers = sheet.col_values(2, start_rowx=1, end_rowx=None)
    return id_numbers


def get_information(file_name, cookie):
    id_numbers = get_id_numbers_from_workbook(file_name)
    num = len(id_numbers)
    information = [['' for i in range(3)] for j in range(num)]
    for i in range(num):
        count = 1  # 尝试连接的次数，初始化为0
        id_number = id_numbers[i]
        message = "正在查询{0}，这是第{1:>2d}次查询，进程为{2:>4d}/{3:>4d}, 进度为{4:>6.2f}%"
        while True:
            print('\r', message.format(id_number, count, i+1, num, (i+1)*100/num), end='', flush=True)
            try:
                applydate, insName = get_student_information(id_number, cookie)
                information[i][0] = applydate
                information[i][1] = insName        
                information[i][2] = information_processing(get_training_record(id_number, cookie))
                break
            except requests.exceptions.RequestException:  # 获取异常,查询异常时，如网速过慢超时
                sleep(1)
                count += 1  # 查询次数加1
                if count > 20:
                        sys.exit("网络不畅，请稍后再试！")
    return information


def output(file_name):
    information = get_information(file_name, cookie)
    title = ['报名时间', '培训机构', '报审情况']  # 自定义标题行
    col_num = len(title)  # 获取列数
    work_book = xlrd.open_workbook(file_name)  # 打开xls表
    sheet = work_book.sheet_by_index(0)
    ncols = sheet.ncols
    copy_work_book = copy(work_book)
    copy_sheet = copy_work_book.get_sheet(0)
    [copy_sheet.write(0, ncols+j, title[j]) for j in range(col_num)]
    [[copy_sheet.write(i+1, ncols+j, information[i][j]) for j in range(col_num)] for i in range(len(information))]
    copy_work_book.save('结果-{0}.xls'.format(datetime.now().strftime('%Y%m%d-%H%M%S')))


start = time()  # 程序开始计时
cookie = 'JSESSIONID=9E64AF00D4271C393FE98A177C027597; td_cookie=266670215'
file_name = '6月1日至30日变更考试地.xls'
output(file_name)
end = time()  # 程序计时结束
print("本次批量处理结束，共用时：{0:>6.2f}s".format(end - start))  # 输出程序运行时间
