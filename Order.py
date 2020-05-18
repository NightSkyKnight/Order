# coding=utf-8

import threading
import _thread
import operator
import codecs
import jsonpath
import json
import requests
from jsonpath import jsonpath
import os
import random
import copy
import time
import win32api
import easygui
import sys
import datetime
import signal


# 存储获取到的原数据
b = {}
# 对源数据进行整理
gd = {'gd_title': [], 'gd_name': [], 'gd_nr': [], 'gd_kf': []}
gd_a = {'gd_title': [], 'gd_name': [], 'gd_nr': [], 'gd_kf': []}
# 初始数目
total = 0
Account_password = []
# 获取的cookies
getcookies = []
# 标识：cookies
mycookies = []
# 浏览器标识池
user_agents = ['Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3704.400 QQBrowser/10.4.3587.400', 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36',
               'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Trident/4.0; SE 2.X MetaSr 1.0; SE 2.X MetaSr 1.0; .NET CLR 2.0.50727; SE 2.X MetaSr 1.0)', 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; 360SE)', 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 UBrowser/6.2.4094.1 Safari/537.36']
cookies_b = 0


# 输入
def sr():
    a = ['0', '1']
    a[0] = input("输入账号：")
    wenj = open("账号密码.txt", "w", encoding='utf-8')
    print(a[0], file=wenj)
    a[1] = input("输入密码：")
    wenj = open("账号密码.txt", "a", encoding='utf-8')
    print(a[1], file=wenj)


# 登录并获取COOKIES
def login():
    name = []
    try:
        with open('账号密码.txt', 'r+', encoding='utf-8') as f:
            Account_password = f.readlines()
            if Account_password == None:
                print('请输入账号密码')
                sr()
    except FileNotFoundError:
        sr()
        with open('账号密码.txt', 'r+', encoding='utf-8') as f:
            Account_password = f.readlines()

    for i in range(0, 2):
        Account_password[i] = Account_password[i].rstrip('\n')
    # 登录Post数据
    postdata_2 = {
        "act": "login",
        "adminuser": Account_password[0],
        "password": Account_password[1]
    }
    print("请耐心等待。。。。")
    # 循环请求
    for h in range(len(user_agents)):
        # 构造header
        header = {
            'Host': 'crm.114my.cn:1808',
            'User-Agent': user_agents[h],
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Referer': 'http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/public/login.php',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        header_2 = {
            'Host': 'crm.114my.cn:1808',
            'User-Agent': user_agents[h],
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Referer': 'http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/public/login.php',
            'Connection': 'close'
        }
        header_3 = {
            'Host': 'crm.114my.cn:1808',
            'User-Agent': user_agents[h],
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-index.php',
            'Connection': 'close'
        }
        header_4 = {
            'Host': 'crm.114my.cn:1808',
            'User-Agent': user_agents[h],
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-index.php',
            'Connection': 'close'
        }
        # 开始请求
        login_cw = len(
            '<script>alert("用户或密码错误");window.location.href="/";</script>')
        s = requests.session()
        r_1 = s.post("http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/public/login.php",
                     headers=header, data=postdata_2)
        login_acw = len(r_1.text)
        if login_acw == login_cw:
            easygui.msgbox('用户或密码错误(ps:重新打开程序并在“账号密码.tx”里重新修改为正确的账号密码)')
            sys.exit()
        time.sleep(1)
        r_2 = s.get(
            'http://crm.114my.cn:1808/html/admin/view/public-index.php', headers=header_2)
        time.sleep(1)
        r_3 = s.get(
            'http://crm.114my.cn:1808/html/admin/ng-include/inc-page-info.php', headers=header_3)
        time.sleep(1)
        r_4 = s.post(
            'http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/common/global-data.php', headers=header_4)
        # 添加COOKIES ID
        getcookies.append(s.cookies['PHPSESSID'])
        # 集合
        # 存储账号密码
        mycookies.append([user_agents[h], getcookies[h]])


def get_cookies():
    cookies_a = 0
    for h in range(len(user_agents)):
        # 从主页过到内页所需请求
        get_gd = requests.session()
        curl_b = {
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'User-Agent': mycookies[cookies_a][0],
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-index.php',
            'Cookie': 'PHPSESSID='+mycookies[cookies_a][1]
        }
        b = get_gd.get(
            'http://crm.114my.cn:1808/html/admin/view/public-ticket-manage.php',  headers=curl_b)
        time.sleep(1)

        curl_dd = {
            'Host': 'crm.114my.cn:1808',
            'User-Agent': mycookies[cookies_a][0],
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-index.php',
            'Connection': 'close',
            'Cookie': 'PHPSESSID='+mycookies[cookies_a][1]
        }
        dd = get_gd.post(
            'http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/common/global-data.php',  headers=curl_dd)
        time.sleep(1)

        curl_d = {
            'Host': 'crm.114my.cn:1808',
            'User-Agent': mycookies[cookies_a][0],
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-ticket-manage.php',
            'Connection': 'close',
            'Cookie': 'PHPSESSID='+mycookies[cookies_a][1]
        }
        d = get_gd.post(
            'http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/common/global-data.php',  headers=curl_d)
        # print(d.json())
        time.sleep(1)
        curl_c = {
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'User-Agent': mycookies[cookies_a][0],
            'Content-Type': 'application/json;charset=utf-8',
            'Accept': 'application/json, text/plain, */*',
            'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-ticket-manage.php',
            'Cookie': 'PHPSESSID='+mycookies[cookies_a][1]
        }
        DATAC = {"page": "1", "page_size": 10, "sdata": {"client_type": "1", "client_keyword": "", "promotion_type": "1", "promotion_keyword": "", "salesman_type": "1", "salesman_keyword": "", "salesman_department": 0, "time_type": "1", "start_time": "", "end_time": "", "status": "-1", "order_no": "", "designer_department": "", "designer_u_id": "", "search_tc_is_pay": "1", "search_design_date_type": "1", "for_design_start_time": "", "for_design_end_time": "", "companyname": "", "bill_type": "0", "bill_status": "0",
                                                         "is_winpop": 'false', "contract_no": "", "pro_package_id": "0", "order_state": "0", "arrear_type": "0", "order_by_type": "0", "need_icp": "1", "action_type": "0", "isExpired": "0", "order_chargeback": "0", "query_type": "2", "remark_type": "1", "remark_keyword": "", "do_tclass": "", "department_id": 0, "subtype": "0", "handled_select": {"department_id": "0", "login_name": "", "nickname": "", "zy_user_id": "0"}, "to_department_id": "", "huifang": "-1", "chkoneself": "n", "an_type": "n"}}
        c = get_gd.post('http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/common/tickets-manage-list.php',
                        data=json.dumps(DATAC), headers=curl_c)
        time.sleep(1)
        curl_e = {
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'User-Agent': mycookies[cookies_a][0],
            'Content-Type': 'application/json;charset=utf-8',
            'Accept': 'application/json, text/plain, */*',
            'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-ticket-manage.php',
            'Cookie': 'PHPSESSID='+mycookies[cookies_a][1]
        }
        DATAe = {
            "href": "http://crm.114my.cn:1808/html/admin/view/public-ticket-manage.php"}
        e = requests.post('http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/common/get-staff-info.php',
                          data=json.dumps(DATAe), headers=curl_e)
        cookies_a = cookies_a+1


# 请求模块
def qinqiu():
    global cookies_b
    # 发送数据包
    DATA = {"page": 1, "page_size": 10, "sdata": {"client_type": "1", "client_keyword": "", "promotion_type": "1", "promotion_keyword": "", "salesman_type": "1", "salesman_keyword": "", "salesman_department": 0, "time_type": "1", "start_time": "", "end_time": "", "status": "-1", "order_no": "", "designer_department": "", "designer_u_id": "", "search_tc_is_pay": "1", "search_design_date_type": "1", "for_design_start_time": "", "for_design_end_time": "", "companyname": "", "bill_type": "0", "bill_status": "0",
                                                  "is_winpop": 'false', "contract_no": "", "pro_package_id": "0", "order_state": "0", "arrear_type": "0", "order_by_type": "0", "need_icp": "1", "action_type": "0", "isExpired": "0", "order_chargeback": "0", "query_type": "1", "remark_type": "1", "remark_keyword": "", "do_tclass": "", "department_id": 0, "subtype": "0", "handled_select": {"department_id": "0", "login_name": "", "nickname": "", "zy_user_id": "0"}, "to_department_id": "", "huifang": "-1", "chkoneself": "n", "an_type": "n"}}

    DATA_2 = {"page": 2, "page_size": 10, "sdata": {"client_type": "1", "client_keyword": "", "promotion_type": "1", "promotion_keyword": "", "salesman_type": "1", "salesman_keyword": "", "salesman_department": 0, "time_type": "1", "start_time": "", "end_time": "", "status": "-1", "order_no": "", "designer_department": "", "designer_u_id": "", "search_tc_is_pay": "1", "search_design_date_type": "1", "for_design_start_time": "", "for_design_end_time": "", "companyname": "", "bill_type": "0", "bill_status": "0",
                                                    "is_winpop": 'false', "contract_no": "", "pro_package_id": "0", "order_state": "0", "arrear_type": "0", "order_by_type": "0", "need_icp": "1", "action_type": "0", "isExpired": "0", "order_chargeback": "0", "query_type": "1", "remark_type": "1", "remark_keyword": "", "do_tclass": "", "department_id": 0, "subtype": "0", "handled_select": {"department_id": "0", "login_name": "", "nickname": "", "zy_user_id": "0"}, "to_department_id": "", "huifang": "-1", "chkoneself": "n", "an_type": "n"}}

    DATA_3 = {"page": 3, "page_size": 10, "sdata": {"client_type": "1", "client_keyword": "", "promotion_type": "1", "promotion_keyword": "", "salesman_type": "1", "salesman_keyword": "", "salesman_department": 0, "time_type": "1", "start_time": "", "end_time": "", "status": "-1", "order_no": "", "designer_department": "", "designer_u_id": "", "search_tc_is_pay": "1", "search_design_date_type": "1", "for_design_start_time": "", "for_design_end_time": "", "companyname": "", "bill_type": "0", "bill_status": "0",
                                                    "is_winpop": 'false', "contract_no": "", "pro_package_id": "0", "order_state": "0", "arrear_type": "0", "order_by_type": "0", "need_icp": "1", "action_type": "0", "isExpired": "0", "order_chargeback": "0", "query_type": "1", "remark_type": "1", "remark_keyword": "", "do_tclass": "", "department_id": 0, "subtype": "0", "handled_select": {"department_id": "0", "login_name": "", "nickname": "", "zy_user_id": "0"}, "to_department_id": "", "huifang": "-1", "chkoneself": "n", "an_type": "n"}}
    DATA_4 = {"page": 4, "page_size": 10, "sdata": {"client_type": "1", "client_keyword": "", "promotion_type": "1", "promotion_keyword": "", "salesman_type": "1", "salesman_keyword": "", "salesman_department": 0, "time_type": "1", "start_time": "", "end_time": "", "status": "-1", "order_no": "", "designer_department": "", "designer_u_id": "", "search_tc_is_pay": "1", "search_design_date_type": "1", "for_design_start_time": "", "for_design_end_time": "", "companyname": "", "bill_type": "0", "bill_status": "0",
                                                    "is_winpop": 'false', "contract_no": "", "pro_package_id": "0", "order_state": "0", "arrear_type": "0", "order_by_type": "0", "need_icp": "1", "action_type": "0", "isExpired": "0", "order_chargeback": "0", "query_type": "1", "remark_type": "1", "remark_keyword": "", "do_tclass": "", "department_id": 0, "subtype": "0", "handled_select": {"department_id": "0", "login_name": "", "nickname": "", "zy_user_id": "0"}, "to_department_id": "", "huifang": "-1", "chkoneself": "n", "an_type": "n"}}

    # URL地址
    url = 'http://crm.114my.cn:1808/ctrl/ctrl.php?path=admin/common/tickets-manage-list-cla0.php'
    # 请求头
    curl = {
        'Origin': 'http://crm.114my.cn:1808',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'User-Agent': mycookies[cookies_b][0],
        'Content-Type': 'application/json;charset=UTF-8',
        'Accept': 'application/json, text/plain, */*',
        'Referer': 'http://crm.114my.cn:1808/html/admin/view/public-ticket-manage.php',
        'Cookie': 'PHPSESSID='+mycookies[cookies_b][1]
    }

    get_gd = requests.session()
    if total == 0 or total <= 10:
        a = get_gd.post(url, data=json.dumps(DATA), headers=curl)
        print('工单数：', total)
    elif total > 10 and total <= 20:
        a = get_gd.post(url, data=json.dumps(DATA_2), headers=curl)
        print('工单数：', total)
    elif total > 20 and total <= 30:
        a = get_gd.post(url, data=json.dumps(DATA_3), headers=curl)
        print('工单数：', total)
    else:
        a = get_gd.post(url, data=json.dumps(DATA_4), headers=curl)
        print('工单数：', total)
    # 检查返回码
    # print(a.status_code)  # 获取服务器返回的状态码 2为正常
    # 更换标识
    cookies_b = cookies_b+1
    if cookies_b == 5 or cookies_b > 5:
        cookies_b = 0
    return a

# END


# 解析并输出模块
# 循环遍历字典里含有列表的值
def huoq(c):  # 获取字典里面的列表   b['data']['data']
    for item in c:  # 循环赋值
        gd['gd_title'].append(item['step_title'])
        gd['gd_name'].append(item['company']['companyname'])
        gd['gd_nr'].append(item['step_remark'])
        gd['gd_kf'].append(item['do_nickname'])


# 输出到文本
def shuchu():
    data = open("one.txt", 'w', encoding='utf-8')
    j = 0
    for dayin in gd['gd_title']:
        print('标题：', end='', file=data)
        print(gd['gd_title'][j], file=data)
        print('', file=data)
        print('公司名：', end='', file=data)
        print(gd['gd_name'][j], file=data)
        print('', file=data)
        print('内容：', end='', file=data)
        print(gd['gd_nr'][j], file=data)
        print('', file=data)
        print('客服：', end='', file=data)
        print(gd['gd_kf'][j], file=data)
        print('', file=data)
        print('------------------', file=data)
        j += 1
    data.flush()
    data.close()

# END


# 时间点判断，并决定sleep时间
def sleep_time(new_time):
    # 获取新时间的年月日
    new_time_ymd = new_time.strftime('%Y-%m-%d')
    # 把年月日拼接到指定时分秒
    end_time = datetime.datetime.strptime(
        new_time_ymd+' 14:00:00', '%Y-%m-%d %H:%M:%S')
    # 计算差值
    sleep_time_seconds = (end_time-new_time).__getattribute__('seconds')

    return sleep_time_seconds


# 暂停程序
def handler(signal_num, frame):
    if(bool(signal_num)):
        print('你已暂停！！！')
        print()
        os.system('pause')


if __name__ == "__main__":
    os.system('title windows update')
    # 登录
    login()
    # 过到内容页
    get_cookies()
    while 1:
        # 当按下'ctrl+c'时触发信号，暂停程序
        signal.signal(signal.SIGINT, handler)
        # 获取最新时间
        new_time = datetime.datetime.now()
        if(new_time.__getattribute__('hour') >= int(12) and new_time.__getattribute__('hour') < int(14)):
            sleep_time_seconds = sleep_time(new_time)
            if(sleep_time_seconds != 0):
                print('砸瓦鲁多')
                time.sleep(sleep_time_seconds)
                # 清空秒数
                sleep_time_seconds = 0
                # 更新时间
                continue
        else:
            gd = {'gd_title': [], 'gd_name': [],
                  'gd_nr': [], 'gd_kf': []}  # 重新创建字典把原先值清空
            # time.sleep(random.uniform(3,4))#随机休眠时间
            # time.sleep(1)
            try:
                b = qinqiu().json()  # 获取返回的JSON数据
            except (ValueError, NameError, TypeError, KeyError, Exception):
                easygui.msgbox(msg="出现错误，程序正在重启", title="Error")
                time.sleep(3)
            try:
                huoq(b['data']['data'])  # 提取JSON数据存放
                total = (b['data']['page_info']['total'])
            except (KeyError, TypeError):
                easygui.msgbox(msg="登录过期！请重新登录！", title="Error")
                login()
                get_cookies()
                continue
            if(operator.eq(gd_a['gd_title'], gd['gd_title'])):  # 比较2个列表的gd_title值是不是相同
                print('无变化')
                print(time.strftime('%H:%M:%S', time.localtime(time.time())))  # 输出时间
                print()
            else:
                gd_a['gd_title'].clear()  # 清除列表
                gd_a['gd_title'].extend(copy.deepcopy(gd['gd_title']))  # 深拷贝
                shuchu()
                win32api.ShellExecute(0, 'open', 'one.txt', '', '', 1)  # 打开文本
                gd = {'gd_title': [], 'gd_name': [], 'gd_nr': [], 'gd_kf': []}

            time.sleep(1)
            gd = {'gd_title': [], 'gd_name': [],
                  'gd_nr': [], 'gd_kf': []}  # 重新创建字典把原先值清空
            try:
                b = qinqiu().json()  # 获取返回的JSON数据
            except (ValueError, NameError, TypeError, KeyError, Exception):
                easygui.msgbox(msg="出现错误，程序正在重启", title="Error")
                time.sleep(3)
            try:
                huoq(b['data']['data'])  # 提取JSON数据存放
                total = (b['data']['page_info']['total'])
            except (KeyError, TypeError):
                easygui.msgbox(msg="登录过期！请重新登录！", title="Error")
                login()
                get_cookies()
                continue
            if(operator.eq(gd_a['gd_title'], gd['gd_title'])):  # 比较2个列表的gd_title值是不是相同
                print('无变化')
                print(time.strftime('%H:%M:%S', time.localtime(time.time())))  # 输出时间
                print()
            else:
                gd_a['gd_title'].clear()  # 清除列表
                gd_a['gd_title'].extend(copy.deepcopy(gd['gd_title']))  # 深拷贝
                shuchu()
                win32api.ShellExecute(0, 'open', 'one.txt', '', '', 1)  # 打开文本
