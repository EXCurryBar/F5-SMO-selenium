#
#                       _oo0oo_
#                      o8888888o
#                      88" . "88
#                      (| -_- |)
#                      0\  =  /0
#                    ___/`---'\___
#                  .' \\|     |# '.
#                 / \\|||  :  |||# \
#                / _||||| -:- |||||- \
#               |   | \\\  -  #/ |   |
#               | \_|  ''\---/''  |_/ |
#               \  .-\__  '-'  ___/-. /
#             ___'. .'  /--.--\  `. .'___
#          ."" '<  `.___\_<|>_/___.' >' "".
#         | | :  `- \`.;`\ _ /`;.`/ - ` : | |
#         \  \ `_.   \_ __\ /__ _/   .-` /  /
#     =====`-.____`.___ \_____/___.-`___.-'=====
#                       `=---='
#
#
#     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#               佛祖保佑         永无BUG
#
#
#
# ============ import thingie =============
import re
import os
import csv
import gzip
import socket
import shutil
import logging
import paramiko
import threading
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from scp import SCPClient
from time import sleep, time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import Select

# ============= initialize thingie ============
PATH = os.path.abspath(os.getcwd())
csvfile = open("data.csv", "a", newline='', encoding='UTF-8')
writer = csv.writer(csvfile)
now = datetime.now().strftime("%Y%m%d")
today = datetime.today()
logging.basicConfig(level=logging.WARNING,
                    format="%(asctime)s %(levelname)s %(message)s",
                    datefmt="%Y-%m-%d %H:%M",
                    handlers=[logging.FileHandler("SMO.log", "w", "utf-8"), ])


def words(ti, da, num):
    # 這裡問henry
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
    # 這裡問henry
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


def is_avail(IP, port=22):
    # is_avail(IP)       --> 如果22port能連接回傳True
    # is_avail(IP, port) --> 如果指定port能連接回傳True
    try:
        host = socket.gethostbyname(IP)
        s = socket.create_connection((host, port), 2)   # 檢查指定IP:port是否有回應，timeout 2 秒
        return True
    except:
        logging.error("無法連線到 " + IP)
    return False


def get_qkview(client, hostname):
    # 產生qkview然後存到 \qkviews
    # 參數 client 用來進行ssh連線
    # 參數 hostname 進行檔案命名
    try:
        stdin, stdout, stderr = client.exec_command(
            "qkview;mv /var/tmp/"+hostname+".qkview /var/tmp/" + hostname + "_" + now + ".qkview")  # 在底層下產生qkview的指令並重新命名為hostname_日期.qkview
        dummy = stdout.readlines()  # 不加這行程式會繼續執行而不會等qkview產完

        scp = SCPClient(client.get_transport()) # 準備以scp協定傳輸檔案
        scp.get("/shared/tmp/" + hostname + "_" + now + ".qkview", PATH + "\\qkviews\\"+ hostname + "_" + now + ".qkview")  # 將qkview 複製到本地端qkviews資料夾
        print(hostname, " Qkview saved")
        client.exec_command("rm -rf /shared/tmp/" + hostname + "_" + now + ".qkview")   # 移除此台f5上的qkview檔案
        return "OK"
    except:
        logging.error("無法取得Qkview " + IP)
        return "Error"


def get_ucs(client, hostname):
    # 產生ucs然後存到 \ucs
    # 參數 client 用來進行ssh連線
    # 參數 hostname 進行檔案命名
    try:
        stdin, stdout, stderr = client.exec_command(
            "tmsh save /sys ucs /var/local/ucs/" + hostname + '_' + now + ".ucs") # 在底層下產生ucs的指令並重新命名為hostname_日期.ucs
        dummy = stdout.readlines()  # 不加這行程式會繼續執行而不會等ucs產完

        scp = SCPClient(client.get_transport()) # 準備以scp協定傳輸檔案
        scp.get("/var/local/ucs/" + hostname + "_" + now + ".ucs", PATH + "\\ucs\\" + hostname + "_" + now + ".ucs")    # 將ucs複製到本地端ucs資料夾
        print(hostname, " UCS saved")
        return "OK"
    except:
        logging.error("無法取得UCS " + IP)
        return "Error"


def ltm(IP, ACC, PASS): 
    # 找出錯誤log並儲存到 \IP_ERR_LOG中
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(IP, username=ACC, password=PASS)                 # 以paramiko進行ssh連線
        _, stdout, _ = client.exec_command("cd /var/log; ls")           # 列出 /var/log目錄下所有檔案

        log = ""
        lst = [line.replace('\n', '') for line in stdout.readlines() if line[:3] == "ltm"]  # 尋找ltm log並儲存到lst中
        scp = SCPClient(client.get_transport()) # 準備以scp協定傳輸檔案
        for line in lst:
            scp.get("/var/log/" + line, PATH + "\\" + IP + "_log\\" + line) # 取得/var/log中的ltm log
            if line[-2:] == "gz":               # 如果該log為壓縮擋 則解壓縮再讀取
                with gzip.open(IP+"_log\\"+line, "rb") as f_in:
                    log += f_in.read().decode() # log字串增加讀取到的檔案內容
            else:                               # 如果該log不是壓縮擋 則直接讀取
                with open(IP+"_log\\"+line, "rb") as f_in:  
                    log += f_in.read().decode() # log字串增加讀取到的檔案內容
    except Exception as e:
        print(e)
        return
# == HA state change
    P = re.compile("\n.*HA unit.*\n")
    res = re.findall(P, log)    
    if len(res) != 0:
        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
        with open(IP+"_ERR_LOG\\HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*No failover status messages received for.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
        with open(IP+"_ERR_LOG\\HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*Active\n")
    res = re.findall(P, log)
    if len(res) != 0:
        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
        with open(IP+"_ERR_LOG\\HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*Offline\n")
    res = re.findall(P, log)
    if len(res) != 0:
        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
        with open(IP+"_ERR_LOG\\HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*Standby\n")
    res = re.findall(P, log)
    if len(res) != 0:
        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
        with open(IP+"_ERR_LOG\\HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)
# == VS state change
    P = re.compile("\n.*Virtual Address .*GREEN to RED.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
        with open(IP+"_ERR_LOG\\VS_ERR.log", "w", newline='') as ef:
            ef.writelines(res)
# == Pool
    P = re.compile("\n.*Pool.*GREEN to RED.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
        with open(IP+"_ERR_LOG\\Pool_ERR.log", "w", newline='') as ef:
            ef.writelines(res)

    # == Template
    # P = re.compile("\n.<這裡輸入log特徵>.*\n")
    # res = re.findall(P, log)
    # if len(res) != 0:        
    #     if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
    #     with open(IP+"_ERR_LOG\\<這裡輸入錯誤名稱>_ERR.log", "a", newline='') as ef:
    #         ef.writelines(res)


def syst(IP, ACC, PASS):
    # 找出錯誤log並儲存到 \IP_ERR_LOG中
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(IP, username=ACC, password=PASS)
        _, stdout, _ = client.exec_command("cd /var/log; ls")

        lst = [line.replace('\n', '')
               for line in stdout.readlines() if line[:8] == "messages"]
        log = ""
        scp = SCPClient(client.get_transport())
        for line in lst:
            scp.get("/var/log/" + line, PATH + "\\" + IP + "_log\\" + line)
            if line[-2:] == "gz":
                with gzip.open(IP+"_log\\"+line, "rb") as f_in:
                    log += f_in.read().decode()
            else:
                with open(IP+"_log\\"+line, "rb") as f_in:
                    log += f_in.read().decode()
    except Exception as e:
        print(e)
        return


def healthCheck(IP):
    # check if device is alive
    global pass_count
    if not (is_avail(IP) and is_avail(IP, 443)):
        shutil.rmtree(IP, ignore_errors=True)
        pass_count += 1
        logging.error("與 " + IP + " 連線中斷")
        return True
    return False


def change_unit(value, unit = "bps"):
    # 1000 --> 1k<unit>, 1000000 --> 1M<unit> ...
    scale = ['', 'k', 'M', 'G', 'T']
    count = 0
    while(value/1000 >= 1):
        value = round(value/1000, 2)
        count += 1
    return str(value)+scale[count]+unit


def get_data(IP, ACC, PASS, sleep_time=5):
    # IP            --> 設備IP
    # ACC           --> 設備帳號
    # PASS          --> 設備密碼
    # sleep_time    --> 每步delay時間，預設5秒
    global pass_count
    if is_avail(IP) and is_avail(IP, 443):  # 先判斷22 443能不能連線
        client = paramiko.SSHClient()       # 如果可以則先宣告ssh client
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())    # 問henry
        client.connect(IP, 22, username=ACC, password=PASS, timeout=10)
        try:
            os.makedirs(IP + "_log")        # 建立log資料夾
        except:
            pass
    else:                                   # 若無法連線
        print("\n無法連線到 " + IP)
        pass_count += 1                     # 通過計數+1
        return
    
    # 初始化表格資料，因為收集到資料才會改動，所以宣告為ERROR
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
# FIXING ========
    # 待更新
    sys_log = "OK"
# =============
    # 初始化 selenium
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": PATH + "\\" + IP}
    options.add_experimental_option("prefs", prefs)
    options.add_argument('ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument("--disable-extensions")

    driver = webdriver.Chrome(chrome_options=options, executable_path=PATH + "\\chromedriver.exe")
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

        loc_time_hr = int(local_time.split(':')[0])
        loc_time_min = int(local_time.split(':')[1])

        diff_hr = int(sys_time_hr) - loc_time_hr
        diff_min = int(sys_time_min) - loc_time_min

        diff_time = diff_min + diff_hr * 60

        if diff_time > 0:
            sys_time = "快" + (str(diff_hr) + "小時") * int(diff_hr != 0) + str(diff_min) + "分鐘"
        elif diff_time < 0:
            sys_time = "慢" + (str(abs(diff_hr)) + "小時") * int(diff_hr != 0) + str(abs(diff_min)) + "分鐘"
        else:
            sys_time = "OK"
    except:
        logging.error("無法獲取時間資訊 " + IP)
        sys_time = "ERROR"

    if healthCheck(IP):
        driver.close()
        return
# ============= ha =============
    try:
        sys_ha = driver.find_element_by_id("status").text.split('\n')[-1]
    except:
        logging.error("無法取得HA資訊 " + IP)
        sys_ha = "ERROR"

    if healthCheck(IP):
        driver.close()
        return
# ============= hostname =============
    try:
        sys_host = driver.find_element_by_id("deviceid").text.split('\n')[1]
    except:
        logging.error("無法取得hostname資訊 " + IP)
        sys_host = "ERROR"

    if healthCheck(IP):
        driver.close()
        return
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
        
    if healthCheck(IP):
        driver.close()
        return
# ============= ucs, qkview =============
    t1 = threading.Thread(target=get_qkview, args=(client,sys_host))
    t2 = threading.Thread(target=get_ucs, args=(client,sys_host))
    t1.start()
    t2.start()

    if healthCheck(IP):
        driver.close()
        return
# ============= syslog =============
    t3 = threading.Thread(target=ltm, args=(IP,ACC,PASS))
    t4 = threading.Thread(target=syst, args=(IP,ACC,PASS))

    t3.start()
    t4.start()

    if healthCheck(IP):
        driver.close()
        return    
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

    if healthCheck(IP):
        driver.close()
        return
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
                    expired.append(certificates[i-1].split(' ')[0] + "\t\t" + certificates[i])
                elif d1 >= today:
                    near_expired.append(certificates[i-1].split(' ')[0] + "\t\t" + certificates[i])
        if len(expired) == 0 and len(near_expired) == 0:
            sys_cert = "OK"
        else:
            sys_cert = "expired: " + str(len(expired)) + " near expire:" + str(len(near_expired))
            with open(IP+"_Certificate.txt","w",encoding="utf-8") as cert_file:
                cert_file.write("Near Expire:\n")
                [cert_file.write(row + "\n") for row in near_expired]
                cert_file.write("\nExpired:\n")
                [cert_file.write(row + "\n") for row in expired]
    except:
        logging.error("無法取得憑證資訊 " + IP)

    if healthCheck(IP):
        driver.close()
        return
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

    if healthCheck(IP):
        driver.close()
        return
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

    if healthCheck(IP):
        driver.close()
        return
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
        logging.error("無法取得記憶體用量 " + IP)
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
        logging.error("無法取得CPU用量 " + IP)
# ============= throughput =============
    try:
        tp = np.array(df["tput_bytes_in"].values.tolist())
        maxium = int(max(tp))
        minimum = int(min(tp))
        sys_tp = change_unit(minimum * 8) + " ~ " + change_unit(maxium * 8)
    except:
        logging.error("無法取得 throughput " + IP)

    if healthCheck(IP):
        driver.close()
        return
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
        sys_ac = change_unit(minimum, "/sec") + " ~ " + change_unit(maxium, "/sec")
    except:
        logging.error("無法取得 active connection " + IP)
# ============= new connection =============
    try:
        nc = np.array(df["totclientconns"].values.tolist())
        maxium = int(round(max(nc)))
        minimum = int(round(min(nc)))
        sys_nc = change_unit(minimum,"/sec") + " ~ " + change_unit(maxium,"/sec")
    except:
        logging.error("無法取得 new connection " + IP)

    if healthCheck(IP):
        driver.close()
        return
# ============= end =============
    shutil.rmtree(IP, ignore_errors=True)
    shutil.rmtree(PATH + "\\" + IP + "_log", ignore_errors=True)
    if len(os.listdir(IP+"_ERR_LOG"))==0:
        shutil.rmtree(PATH + "\\" + IP + "_ERR_LOG", ignore_errors=True)
    driver.close()
    t1.join()
    t2.join()
    sys_qkview = "OK" if os.path.exists(PATH + "\\qkviews\\" + sys_host + '_' + now + ".qkview") else "ERROR"
    sys_ucs = "OK" if os.path.exists(PATH + "\\ucs\\" + sys_host + '_' + now + ".ucs") else "ERROR"
    t3.join()
    t4.join()
    d = os.listdir(IP+"_ERR_LOG")
    for item in d:
        if item[:len(IP)] == IP and item[-7:] == "ERR.log":
            sys_log = "ERROR"
            break

    outgo = [sys_host, sys_sn, sys_uptime, sys_mem, sys_cpu, sys_ac, sys_nc, sys_tp,
             sys_log, sys_ntp, sys_snmp, sys_ucs, sys_qkview, sys_time, sys_cert, sys_ha, sys_ver]
    writer.writerow(outgo)
    print(IP + " 蒐集完畢")
    pass_count += 1


if __name__ == "__main__":
    global doc, t0, word_nn, filecount, pass_count
    # doc, t0, word_nn  --> 問henry
    # filecount         --> 寫檔案用於文件編號
    # pass_count        --> 用於計算執行完畢數量
    pass_count = 0
    filecount = 0
    process_count = 0 # --> 計算同時進行的執行續數量
    devices = pd.read_excel("SMO_ex.xls").values.tolist()   # 讀取excel中的IP、帳號、密碼並儲存於devices中

    try:
        os.makedirs("qkviews")
        os.makedirs("ucs")      # 新增ucs、qkviews資料夾用於存放資料
    except Exception as e:
        print(e)

    for device in devices:      # 逐項抓取設備
        process_count += 1      # 執行續計數+1
        IP = device[0]
        ACCOUNT = device[1]
        PASSWD = device[2]      # 抓取設備IP、帳號、密碼
        t = threading.Thread(target=get_data, args=(IP, ACCOUNT, PASSWD, 25)) # 開啟一個執行續 設定get_data參數延遲25秒(lab太慢)
        t.start()
        if process_count == 4:  # 限制同時搜尋4台設備
            t.join()
            process_count = 0   # 等待其中一台搜尋完畢重設process_count

    while(pass_count<len(devices)): # 等待全部設備搜尋完畢再繼續
        sleep(1)
        
    csvfile.close()             # 關閉收集資料檔
    data_lst = []               # 設定資料暫存的list
    with open("data.csv", "r", encoding="utf-8") as csvf:
        data = csv.reader(csvf)
        for line in data:
            data_lst.append(line)   # 將資料存放到list中

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