import re
import os
import csv
import socket
import logging
import requests
import paramiko
import threading
import pandas as pd
import matplotlib.pyplot as plt
from scp import SCPClient
from time import sleep
from numpy import shape
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.support.ui import Select


csvfile = open("data.csv", "a", newline='',encoding='gbk')
writer = csv.writer(csvfile)
requests.packages.urllib3.disable_warnings()
now = datetime.now().strftime("%Y%m%d")
today = datetime.today()
logging.basicConfig(level=logging.WARNING,
                    format="%(asctime)s %(levelname)s %(message)s",
                    datefmt="%Y-%m-%d %H:%M",
                    handlers=[logging.FileHandler("SMO.log", "w", "utf-8"), ])


def usage(img, color="black"):  # return the maxium and minium of a graph
    Max = 0
    Min = 101
    img = plt.imread(img)
    point1 = []
    point2 = []
    x, y = shape(img)[0:2]

    if color=="black":
        target = [0,0,0]
    elif color == "blue":
        target = [0,0,1]

    for i in range(x):
        for j in range(y):
            if sum(img[i, j][0:3]) == 3:
                point1 = [i, j]
                break
    for i in range(x-1, 0, -1):
        for j in range(y-1, 0, -1):
            if sum(img[i, j][0:3]) == 3:
                point2 = [i, j]
                break
    img = img[point2[0]:point1[0], point1[1]:point2[1]]
    x, y = shape(img)[0:2]
    for i in range(x):
        for j in range(y):
            if sum(img[i, j][0:3] - target) == 0:
                (Max) = i
                break
    for i in range(x-1, 0, -1):
        for j in range(y-1, 0, -1):
            if sum(img[i, j][0:3] - target) == 0:
                Min = i
                break
   
    return [(x-Max) / x, (x-Min) / x]


def is_avail(IP, port=22):  # return True if device is reachable
    try:
        host = socket.gethostbyname(IP)
        s = socket.create_connection((host, port), 2)
        return True
    except:
        logging.error("無法連線到 " + IP)
    return False


def get_qkview(client, hostname):  # generate qkview and saved at C:\qkview
    try:
        stdin, stdout, stderr = client.exec_command(
            "qkview;mv /var/tmp/"+hostname+".qkview /var/tmp/" + hostname + "_" + now + ".qkview")
        dummy = stdout.readlines()

        scp = SCPClient(client.get_transport())
        scp.get("/shared/tmp/" + hostname + "_" + now + ".qkview", "\\qkviews")
        print(hostname, " Qkview saved")
        return "OK"
    except:
        print("Error occur")
        return "Error"


def get_ucs(client, hostname):  # generate ucs and saved at C:\ucs
    try:
        stdin, stdout, stderr = client.exec_command(
            "tmsh save /sys ucs /var/local/ucs/" + hostname + '_' + now + ".ucs")
        dummy = stdout.readlines()

        scp = SCPClient(client.get_transport())
        scp.get("/var/local/ucs/" + hostname + "_" + now + ".ucs", "\\ucs")
        print(hostname, " UCS saved")
        return "OK"
    except:
        return "Error"


def get_data(IP, ACC, PASS, sleep_time=5):

    if is_avail(IP):
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(IP, 22, username=ACC,password=PASS, timeout=10)
    else:
        print("無法連線到 " + IP)

    sys_host = "ERROR"
    sys_sn = "ERROR"
    sys_ver = "ERROR"
    sys_time = "ERROR"
    sys_ha = "ERROR"
    sys_uptime = "ERROR"
    sys_mem = "ERROR"
    sys_cpu = "ERROR"
    sys_cert = "ERROR"
    sys_ntp = "ERROR"
    sys_snmp = "ERROR"
    sys_ucs = "ERROR"
    sys_qkview = "ERROR"
# TODO ===========
    sys_ac = "ERROR"
    sys_nc = "ERROR"
    sys_tp = "ERROR"
    sys_log = "ERROR"
# ================
    options = webdriver.ChromeOptions()
    options.add_argument('ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument("--disable-extensions")

    driver = webdriver.Chrome(chrome_options=options)
    driver.get("https://" + IP + "/tmui/login.jsp")
    driver.find_element_by_id("username").send_keys(ACC)
    driver.find_element_by_id("passwd").send_keys(PASS)
    driver.find_element_by_xpath("//button[1]").click()
    sleep(sleep_time)
# ============= time =============
    try:
        time = driver.find_element_by_id("dateandtime")
        system_time = time.text.split('\n')[-1].split(' ')[0:2]
        sys_time_hr, sys_time_min = system_time[0].split(':')
        sys_time_hr = str(int(sys_time_hr) + 12 * int(system_time[1] == "PM"))
        system_time = sys_time_hr.zfill(2)+':'+sys_time_min.zfill(2)

        local_time = datetime.strftime(datetime.now(), "%H:%M")
        sys_time = "OK" if system_time == local_time else "NO"
    except:
        logging.error("無法獲取時間資訊 " + IP)
        sys_time = "ERROR"
# ============= ha =============
    try:
        sys_ha = driver.find_element_by_id("status").text.split('\n')[1]
    except:
        logging.error("無法取得HA資訊 " + IP)
        sys_ha = "ERROR"

# ============= hostname =============
    try:
        sys_host = driver.find_element_by_id("deviceid").text.split('\n')[1]
    except:
        logging.error("無法取得hostname資訊 " + IP)
        sys_host = "ERROR"
# ============= sn, sys_ver =============
    try:
        driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/device/properties_general.jsp")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        items = driver.find_elements_by_class_name("settings")
        sys_sn, sys_ver = [item.text for item in items][1:3]
    except:
        logging.error("無法取得 S/N 或版本資訊 " + IP)
# ============= ucs, qkview =============
    # t1 = threading.Thread(target=get_qkview, args=(client,sys_host))
    # t2 = threading.Thread(target=get_ucs, args=(client,sys_host))
    # t1.start()
    # t2.start()
# ============= uptime =============
    try:
        driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/service/list.jsp")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        uptime_text = driver.find_element_by_id("list_body").text.split('\n')[0].split(' ')[-2:]
        sys_uptime = uptime_text[0] + ' ' + uptime_text[1]
    except:
        logging.error("無法取得uptime資訊 " + IP)
 
# ============= certificate =============
    try:
        driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/locallb/ssl_certificate/list.jsp?&startListIndex=0&showAll=true")
        sleep(sleep_time*3)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))

        certificates = driver.find_element_by_id("list_body").text.split('\n')
        expired = []
        near_expired = []
        for i in range(len(certificates)):
            if(len(certificates[i]) < 13 and certificates[i] != "Common"):
                d1 = datetime.strptime(certificates[i], "%b %d, %Y")
                if d1 < today:
                    expired.append(certificates[i-1].split(' ')[0])
                elif d1 >= today:
                    near_expired.append(certificates[i-1].split(' ')[0])
        if len(expired)==0 and len(near_expired)==0:
            sys_cert = "OK"
        else:
            sys_cert = "已過期: "+ str(len(expired)) + " 快過期:" + str(len(near_expired))
            # print("已過期:",len(expired))
            # [print(item) for item in expired]
            # print("快過期:",len(near_expired))
            # [print(item) for item in near_expired]
    except:
        logging.error("無法取得憑證資訊 " + IP)

# ============= NTP =============
    try:
        driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/device/properties_ntp.jsp")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        lst = [item for item in driver.find_element_by_id("ntp.servers").text.replace(' ', '').split('\n') if item != '']

        if len(lst) == 0:
            sys_ntp = "N/A"
        else:
            sys_ntp = "OK"
    except:
        logging.error("無法取得NTP資訊 " + IP)

# ============= SNMP ============= 
    try:
        driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/snmp/configuration_agent.jsp")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        lst = [item for item in driver.find_element_by_id("snmp_allow_list").text.replace(' ', '').split('\n') if item != '']

        if len(lst) == 0 or lst[0] == "127.0.0.0/8":
            sys_snmp = "N/A"
        else:
            sys_snmp = "OK"
    except:
        logging.error("無法取得SNMP資訊 " + IP)
# ============= mem,cpu,active con., new con., throughput =============
    name_lst = ["mem", "cpu", "ac", "nc", "tp"]
    try:
        driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/stats/list.jsp?subset=All")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        s = Select(driver.find_element_by_name("int_select"))
        s.select_by_value("3")

        img = driver.find_elements_by_tag_name("img")
        img_lst = [item.get_attribute('src') for item in img if item.get_attribute('src')[-3:] == "png"]

        s = requests.session()
        for cookie in driver.get_cookies():
            c = {cookie['name']: cookie['value']}
            s.cookies.update(c)

        for i in range(len(name_lst)):
            r = s.get(img_lst[i], allow_redirects=True, verify=False)
            open(IP + name_lst[i] + '.png', 'wb').write(r.content)

        res = []
        for i in range(len(name_lst)):
            if i == len(name_lst) - 1:
                res.append(usage(IP + name_lst[i] + ".png","blue"))
            else:
                res.append(usage(IP + name_lst[i] + ".png"))
        [os.system("del " + IP + name + ".png") for name in name_lst]

        if res[0][0] == res[0][1]:
            sys_mem = str(int(res[0][0] * 100)) + "%"
        else:
            sys_mem = str(int(res[0][0] * 100)) + "% ~ " + str(int(res[0][1] * 100)) + "%"

        if res[1][0] == res[1][1]:
            sys_cpu = str(int(res[1][0] * 100)) + "%"
        else:
            sys_cpu = str(int(res[1][0] * 100)) + "% ~ " + str(int(res[1][1] * 100)) + "%"

        
    except Exception as e:
        logging.error("無法取得CPU或記憶體資訊 " + IP + str(e))
# ============= end =============
    driver.close()
    # t1.join()
    # t2.join()
    # sys_qkview = "OK" if os.path.exists("C:\\qkviews\\" + sys_host + '_' + now + ".qkview") else "ERROR"
    # sys_ucs = "OK" if os.path.exists("C:\\ucs\\" + sys_host + '_' + now + ".ucs") else "ERROR"

    outgo = [sys_host, sys_sn, sys_uptime, sys_mem, sys_cpu, sys_ac, sys_nc, sys_tp, sys_log, sys_ntp, sys_snmp, sys_ucs, sys_qkview, sys_time, sys_cert, sys_ha, sys_ver]
    print(outgo)
    writer.writerow(outgo)


if __name__ == "__main__":
    process_count = 0
    devices = pd.read_excel("SMO_ex.xls").values.tolist()
    try:
        PATH = os.path.abspath(os.getcwd())
        os.chdir("\\")
        os.system("mkdir qkviews, ucs")  
        os.chdir(PATH)
    except:
        print("please run as administrator")

    for device in devices:
        process_count += 1
        IP = device[0]
        ACCOUNT = device[1]
        PASSWD = device[2]
        t = threading.Thread(target=get_data,args=(IP,ACCOUNT,PASSWD,10))
        t.start()
        if process_count == 4:
            t.join()
            process_count = 0