import re
import os
import csv
import socket
import shutil
import logging
import requests
import paramiko
import threading
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from numpy import shape
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from scp import SCPClient
from time import sleep, time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import Select

PATH = os.path.abspath(os.getcwd())
pass_count = 0
csvfile = open("data.csv", "a", newline='', encoding='UTF-8')
writer = csv.writer(csvfile)
requests.packages.urllib3.disable_warnings()
now = datetime.now().strftime("%Y%m%d")
today = datetime.today()
logging.basicConfig(level=logging.WARNING,
                    format="%(asctime)s %(levelname)s %(message)s",
                    datefmt="%Y-%m-%d %H:%M",
                    handlers=[logging.FileHandler("SMO.log", "w", "utf-8"), ])



def words(ti, da, num):
    global word_nn, doc, t0, filecount
    # print(ti, ": ", da)
    if word_nn == 1:
        t0.cell(num, word_nn).text = da
    elif word_nn == 3:
        t0.cell(num, word_nn).text = da
    elif word_nn == 5:
        t0.cell(num, word_nn).text = da
    else:
        t0.cell(num, 6).text = da
    while True:
        try:
            doc.save("SMO_" + str(filecount) + ".docx")
            break
        except:
            pass


def paste(lst):
    sys_host, sys_sn, sys_uptime, sys_mem, sys_cpu, sys_ac, sys_nc, sys_tp, sys_log, sys_ntp, sys_snmp, sys_ucs, sys_qkview, sys_time, sys_cert, sys_ha, sys_ver = lst
    words("hostname", sys_host, 0)
    words("S/N", sys_sn, 1)
    words("uptime", sys_uptime, 2)
    words("Memory", sys_mem, 3)
    words("CPU", sys_cpu, 4)
    words("Active Connections", sys_ac, 5)
    words("New Connections", sys_nc, 6)
    words("Throughput", sys_tp, 7)
    words("Syslog", sys_log, 8)
    words("NTP", sys_ntp, 9)
    words("SNMP", sys_snmp, 10)
    words("Config Backup",sys_ucs,11)
    words("Qkview Backup",sys_qkview,12)
    words("Time", sys_time, 13)
    words("Certificate status",sys_cert,14)
    words("HA status", sys_ha, 15)
    words("Version", sys_ver, 16)

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


def change_unit(value):
    units = ['bps', 'Kbps', 'Mbps', 'Gbps', 'Tbps']
    count = 0
    while(value/1000 >= 1):
        value = round(value/1000, 2)
        count += 1
    return str(value)+units[count]


def get_data(IP, ACC, PASS, sleep_time=5):

    if is_avail(IP) and is_avail(IP, 443):
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(IP, 22, username=ACC, password=PASS, timeout=10)
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
    sys_ac = "ERROR"
    sys_nc = "ERROR"
    sys_tp = "ERROR"
# TODO ========
    sys_log = "ERROR"
# =============
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": PATH + "\\" + IP}
    options.add_experimental_option("prefs", prefs)
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
        time_device = driver.find_element_by_id("dateandtime")
        system_time = time_device.text.split('\n')[-1].split(' ')[0:2]
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
        sys_ha = driver.find_element_by_id("status").text.split('\n')[-1]
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
        driver.get("https://" + IP +
                   "/tmui/Control/jspmap/tmui/system/device/properties_general.jsp")
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
        driver.get("https://" + IP +
                   "/tmui/Control/jspmap/tmui/system/service/list.jsp")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        uptime_text = driver.find_element_by_id("list_body").text.split('\n')[0].split(',')[0].split(' ')[-2:]
        sys_uptime = uptime_text[0] + ' ' + uptime_text[1]
    except:
        logging.error("無法取得uptime資訊 " + IP)
# ============= certificate =============
    try:
        driver.get("https://" + IP +
                   "/tmui/Control/jspmap/tmui/locallb/ssl_certificate/list.jsp?&startListIndex=0&showAll=true")
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
        if len(expired) == 0 and len(near_expired) == 0:
            sys_cert = "OK"
        else:
            sys_cert = "expired: " + str(len(expired)) + " near expire:" + str(len(near_expired))
            with open(IP+"_Certificate.txt","w",encoding="utf-8") as cert_file:
                cert_file.write("Near Expire:\n")
                [cert_file.write(row + "\n") for row in near_expired]
                cert_file.write("\nExpired:\n")
                [cert_file.write(row + "\n") for row in expired]
            # print("已過期:",len(expired))
            # [print(item) for item in expired]
            # print("快過期:",len(near_expired))
            # [print(item) for item in near_expired]
    except:
        logging.error("無法取得憑證資訊 " + IP)
# ============= NTP =============
    try:
        driver.get("https://" + IP +
                   "/tmui/Control/jspmap/tmui/system/device/properties_ntp.jsp")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        lst = [item for item in driver.find_element_by_id(
            "ntp.servers").text.replace(' ', '').split('\n') if item != '']

        if len(lst) == 0:
            sys_ntp = "N/A"
        else:
            sys_ntp = "OK"
    except:
        logging.error("無法取得NTP資訊 " + IP)
# ============= SNMP =============
    try:
        driver.get("https://" + IP +
                   "/tmui/Control/jspmap/tmui/system/snmp/configuration_agent.jsp")
        sleep(sleep_time)
        driver.switch_to.frame(driver.find_element_by_id("contentframe"))
        lst = [item for item in driver.find_element_by_id(
            "snmp_allow_list").text.replace(' ', '').split('\n') if item != '']

        if len(lst) == 0 or lst[0] == "127.0.0.0/8":
            sys_snmp = "N/A"
        else:
            sys_snmp = "OK"
    except:
        logging.error("無法取得SNMP資訊 " + IP)
# ============= mem =============
    try:
        driver.get("https://" + IP + "/tmui/tmui/util/ajax/data_viz.jsp?cache=" + str(int(time())) + "&name=throughput")
        sleep(sleep_time)
        os.rename(PATH + "\\" + IP +"\\data_viz.jsp", PATH + "\\" + IP + "\\throughput.csv")
        sleep(sleep_time)
        df = pd.read_csv(IP + "\\throughput.csv")
        mem = df[["Rtmmused", "Rtmmmemory"]]
        used = mem["Rtmmused"].values.tolist()
        total = mem["Rtmmmemory"].values.tolist()
        mem_max = 0
        mem_min = 101
        for i in range(len(total)):
            value = round((used[i]/total[i]) * 100)
            mem_max = value if value > mem_max else mem_max
            mem_min = value if value < mem_min else mem_min
        
        if mem_max == mem_min:
            sys_mem = str(mem_min) + "%"
        else:
            sys_mem = str(mem_min) + "% ~ " + str(mem_max) + "%"

    except :
        logging.error("無法取得記憶體用量" + IP)
# ============= cpu =============
    try:
        cpu = df[["Ruser", "Rniced","Rsystem","Ridle","Rirq","Rsoftirq","Riowait"]]
        used =[sum(item) for item in cpu[["Ruser", "Rniced","Rsystem"]].values.tolist()]
        total = [sum(item) for item in cpu.values.tolist()]
        cpu_max = 0
        cpu_min = 101
        for i in range(len(total)):
            value = round((used[i]/total[i]) * 100)
            cpu_max = value if value > cpu_max else cpu_max
            cpu_min = value if value < cpu_min else cpu_min

        if cpu_max == cpu_min:
            sys_cpu = str(cpu_min) + "%"
        else:
            sys_cpu = str(cpu_min) + "% ~ " + str(cpu_max) + "%"

    except :
        logging.error("無法取得CPU用量" + IP)
# ============= throughput =============
    try:
        tp = np.array(df["tput_bytes_in"].values.tolist())
        maxium = int(max(tp))
        minimum = int(min(tp))
        sys_tp = change_unit(minimum * 8) + " ~ " + change_unit(maxium * 8)
    except:
        logging.error("無法取得 throughput " + IP)
# ============= active connection =============
    try:
        driver.get("https://" + IP + "/tmui/tmui/util/ajax/data_viz.jsp?cache=" + str(int(time())) + "&name=connections")
        sleep(sleep_time)
        os.rename(PATH + "\\" + IP + "\\data_viz.jsp", PATH + "\\" + IP + "\\connections.csv")
        sleep(sleep_time)
        df = pd.read_csv(IP + "\\connections.csv")
        ac = np.array(df["curclientconns"].values.tolist())
        maxium = int(round(max(ac)))
        minimum = int(round(min(ac)))
        sys_ac = str(minimum) + " ~ " + str(maxium) + "/s"
    except:
        logging.error("無法取得 active connection " + IP)
# ============= new connection =============
    try:
        nc = np.array(df["totclientconns"].values.tolist())
        maxium = int(round(max(nc)))
        minimum = int(round(min(nc)))
        sys_nc = str(minimum) + " ~ " + str(maxium) + "/s"
    except:
        logging.error("無法取得 new connection " + IP)
# ============= end =============
    shutil.rmtree(IP, ignore_errors=True)
    driver.close()
    # t1.join()
    # t2.join()
    # sys_qkview = "OK" if os.path.exists("C:\\qkviews\\" + sys_host + '_' + now + ".qkview") else "ERROR"
    # sys_ucs = "OK" if os.path.exists("C:\\ucs\\" + sys_host + '_' + now + ".ucs") else "ERROR"

    outgo = [sys_host, sys_sn, sys_uptime, sys_mem, sys_cpu, sys_ac, sys_nc, sys_tp,
             sys_log, sys_ntp, sys_snmp, sys_ucs, sys_qkview, sys_time, sys_cert, sys_ha, sys_ver]
    # print(outgo)
    writer.writerow(outgo)


if __name__ == "__main__":
    threads = []
    process_count = 0
    devices = pd.read_excel("SMO_ex.xls").values.tolist()
    try:
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
        t = threading.Thread(target=get_data, args=(IP, ACCOUNT, PASSWD, 20))
        threads.append(t)
        t.start()
        if process_count == 4:
            t.join()
            process_count = 0

    for x in threads:
        x.join()

    global doc, t0, word_nn, filecount
    filecount = 0
    data_lst = []
    with open("data.csv", "r", encoding="utf-8") as csvf:
        data = csv.reader(csvf)
        for line in data:
            data_lst.append(line)

    for row in range(len(data_lst)):
        if row % 4 == 0:
            doc = Document('example.docx')
            doc.styles['Normal'].font.name = "Times New Roman"
            doc.styles['Normal'].font.size = Pt(10)
            t0 = doc.tables[0]
            word_nn = 1
            filecount += 1
        paste(data_lst[row])
        word_nn += 2