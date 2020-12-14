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
from progressbar import ProgressBar
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
pbar = ProgressBar()
csvfile = open("data.csv", "a", newline='', encoding='UTF-8')
writer = csv.writer(csvfile)
now = datetime.now().strftime("%Y%m%d")
today = datetime.today()
logging.basicConfig(level=logging.WARNING,
                    format="%(asctime)s %(levelname)s %(message)s",
                    datefmt="%Y-%m-%d %H:%M",
                    handlers=[logging.FileHandler("SMO.log", "w", "utf-8"), ])


def words(ti, da, num):
    # 這裡問henry 我複製過來的👍
    global word_nn, doc, t0, filecount
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
    # 這裡問henry 我複製過來的👍
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
        # logging.error("無法連線到 " + IP)
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
        logging.error("與 " + IP + " 連線中斷")
        pass_count += 1
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


def get_data(IP, ACC, PASS):
    # IP            --> 設備IP
    # ACC           --> 設備帳號
    # PASS          --> 設備密碼
    global pass_count
    if is_avail(IP) and is_avail(IP, 443):  # 先判斷22 443能不能連線
        client = paramiko.SSHClient()       # 如果可以則先宣告ssh client
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(IP, 22, username=ACC, password=PASS, timeout=10)
        try:
            os.makedirs(IP + "_log")        # 建立log資料夾
        except:
            pass
    else:                                   # 若無法連線 執行完畢計數+1
        print("\n無法連線到 " + IP)
        pass_count += 1
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
    options.add_experimental_option("prefs", prefs)     # 設置預設下載路徑為 \IP資料夾
    options.add_argument('ignore-certificate-errors')   # 設置忽略憑證錯誤
    options.add_argument('--ignore-ssl-errors')         # 設置忽略ssl錯誤
    options.add_argument("--disable-extensions")        # 關閉外掛功能

    driver = webdriver.Chrome(chrome_options=options, executable_path=PATH + "\\chromedriver.exe")
    driver.get("https://" + IP + "/tmui/login.jsp")     # get登入頁面
    driver.find_element_by_id("username").send_keys(ACC)# 輸入帳號
    driver.find_element_by_id("passwd").send_keys(PASS) # 輸入密碼
    driver.find_element_by_xpath("//button[1]").click()
# ============= time =============
    while not healthCheck(IP):
        try:
            time_device = driver.find_element_by_id("dateandtime")                  # 抓取f5 UI介面上的時間字串
            system_time = time_device.text.split('\n')[-1].split(' ')[0:2]
            sys_time_hr, sys_time_min = system_time[0].split(':')                   # 分出時、分
            sys_time_hr = str(int(sys_time_hr) + 12 * int(system_time[1] == "PM"))  # 將 '時' 轉換成24小時制
            local_time = datetime.strftime(datetime.now(), "%H:%M")                 # 抓取本機時間

            loc_time_hr = int(local_time.split(':')[0])
            loc_time_min = int(local_time.split(':')[1])                            # 分出時、分

            diff_hr = int(sys_time_hr) - loc_time_hr
            diff_min = int(sys_time_min) - loc_time_min

            diff_time = diff_min + diff_hr * 60                                     # 算出相差時間並轉換成分鐘

            if diff_time > 0:                                                       # 讓sys_time儲存時間差，若時間一樣則儲存OK
                sys_time = "快" + (str(diff_hr) + "小時") * int(diff_hr != 0) + str(diff_min) + "分鐘"
            elif diff_time < 0:
                sys_time = "慢" + (str(abs(diff_hr)) + "小時") * int(diff_hr != 0) + str(abs(diff_min)) + "分鐘"
            else:
                sys_time = "OK"
            break
        except:
            pass

    if healthCheck(IP):
        driver.close()
        return
# ============= ha =============
    while not healthCheck(IP):
        try:
            sys_ha = driver.find_element_by_id("status").text.split('\n')[-1]       # 抓取f5 UI顯示的HA資訊
            break
        except:
            pass

    if healthCheck(IP):
        driver.close()
        return
# ============= hostname =============
    while not healthCheck(IP):
        try:
            sys_host = driver.find_element_by_id("deviceid").text.split('\n')[1]    # 抓取f5 UI顯示的hostname
            break
        except:
            pass

    if healthCheck(IP):
        driver.close()
        return
# ============= sn, sys_ver =============
    driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/device/properties_general.jsp")  # 載入 System >> Configuration
    while not healthCheck(IP):
        try:
            driver.switch_to.frame(driver.find_element_by_id("contentframe"))       # 切換到包含資料的frame 
            items = driver.find_elements_by_class_name("settings")                  # 複製UI上表格內的所有資料
            sys_sn, sys_ver = [item.text for item in items][1:3]                    # SN、版本分別為第2與第3項
            break
        except:
            pass
        
    if healthCheck(IP):
        driver.close()
        return
# ============= ucs, qkview =============
    # thread_qkview = threading.Thread(target=get_qkview, args=(client,sys_host))    # 因為產出時間較久，開一個執行續在背景跑
    # thread_ucs = threading.Thread(target=get_ucs, args=(client,sys_host))       # 因為產出時間較久，開一個執行續在背景跑
    # thread_qkview.start()
    # thread_ucs.start()

    if healthCheck(IP):
        driver.close()
        return
# ============= syslog =============
    thread_ltmLog = threading.Thread(target=ltm, args=(IP,ACC,PASS))       # 因為產出時間較久，開一個執行續在背景跑
    thread_systemLog = threading.Thread(target=syst, args=(IP,ACC,PASS))      # 因為產出時間較久，開一個執行續在背景跑

    thread_ltmLog.start()
    thread_systemLog.start()

    if healthCheck(IP):
        driver.close()
        return    
# ============= uptime =============
    driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/service/list.jsp")       # 抓取 System>>Services頁面
    while not healthCheck(IP):
        try:
            driver.switch_to.frame(driver.find_element_by_id("contentframe"))                       # 切換至含有資料的frame
            uptime_text = driver.find_element_by_id("list_body").text.split('\n')[0].split(',')[0].split(' ')[-2:]  # 擷取big3d的運行時間
            sys_uptime = uptime_text[0] + ' ' + uptime_text[1]
            break
        except:
            pass

    if healthCheck(IP):
        driver.close()
        return
# ============= certificate =============
    driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/locallb/ssl_certificate/list.jsp?&startListIndex=0&showAll=true")
    # 抓取憑證頁面並顯示全部憑證
    while not healthCheck(IP):
        try:
            driver.switch_to.frame(driver.find_element_by_id("contentframe"))

            certificates = driver.find_element_by_id("list_body").text.split('\n') 
            # 以換行符號區分各憑證

            # 過期和即將過期憑證的警示符號會被解讀成一個換行符號，
            # 若該憑證過期或即將過期會解讀為以下格式:
            #   憑證資訊 \n 到期日 \n partition \n
            # 而正常未過期憑證則是:
            #   憑證資訊    到期日    partition \n
            # 日期字串長度不會超過13個字元(Jan 26, 2021 --> 12個字元)，且通常憑證儲存在Common
            # 故以此邏輯進行憑證是否過期判斷:

            # 若該項字串長度小於13且不為Common，則其前一項必為該憑證資訊

            expired = []        # 存放過期憑證
            near_expired = []   # 存放即將過期憑證
            for i in range(len(certificates)):
                if(len(certificates[i]) < 13 and certificates[i] != "Common"):  # 判斷是否為快過期或過期憑證的時間
                    d1 = datetime.strptime(certificates[i], "%b %d, %Y")        # 將該時間轉換格式以便比對
                    if d1 < today:                                              # 若憑證過期則加入過期憑證list
                        expired.append(certificates[i-1].split(' ')[0] + "\t\t" + certificates[i])
                    elif d1 >= today:                                           # 若憑證尚未過期則加入即將過期憑證list
                        near_expired.append(certificates[i-1].split(' ')[0] + "\t\t" + certificates[i])
            if len(expired) == 0 and len(near_expired) == 0:                    # 若兩個list皆未儲存任何憑證則讓sys_cert=OK
                sys_cert = "OK"
            else:                                                               # 若有任何一個list有憑證則寫入IP_Certificate.txt中
                sys_cert = "expired: " + str(len(expired)) + " near expire:" + str(len(near_expired))
                with open(IP+"_Certificate.txt","w",encoding="utf-8") as cert_file:
                    cert_file.write("Near Expire:\n")
                    [cert_file.write(row + "\n") for row in near_expired]
                    cert_file.write("\nExpired:\n")
                    [cert_file.write(row + "\n") for row in expired]
            break
        except:
            pass

    if healthCheck(IP):
        driver.close()
        return
# ============= NTP =============
    driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/device/properties_ntp.jsp") 
    # 抓取System>>Configuration:NTP頁面
    while not healthCheck(IP):
        try:
            driver.switch_to.frame(driver.find_element_by_id("contentframe"))
            lst = [item for item in driver.find_element_by_id("ntp.servers").text.replace(' ', '').split('\n') if item != '']   
            # 抓取ntp server內容 如有設定則加入list中

            if len(lst) == 0:       # 如未設定NTP server 則 sys_ntp = N/A， 有設定則填入OK
                sys_ntp = "N/A"
            else:
                sys_ntp = "OK"
            break
        except:
            pass

    if healthCheck(IP):
        driver.close()
        return
# ============= SNMP =============
    driver.get("https://" + IP + "/tmui/Control/jspmap/tmui/system/snmp/configuration_agent.jsp")
    # 抓取System>>SNMP頁面
    while not healthCheck(IP):
        try:
            driver.switch_to.frame(driver.find_element_by_id("contentframe"))
            lst = [item for item in driver.find_element_by_id(
                "snmp_allow_list").text.replace(' ', '').split('\n') if item != '']
            # 抓取snmp內容 加入list中
            if len(lst) == 0 or lst[0] == "127.0.0.0/8":    # 如未設定SNMP則 sys_snmp= N/A， 有設定則填入OK
                sys_snmp = "N/A"
            else:
                sys_snmp = "OK"
            break
        except:
            pass

    if healthCheck(IP):
        driver.close()
        return
# ============= mem =============
    driver.get("https://" + IP + "/tmui/tmui/util/ajax/data_viz.jsp?cache=" + str(int(time())) + "&name=throughput")
    # 抓取Statistic>>Performance>>Traffic Report背景傳輸的throughput CSV檔
    # 檔案內容為30天以來每20分鐘取樣一次的系統效能及流量紀錄
    # 記憶體用量、CPU用量以及throughput皆記錄在同一檔案
    while not healthCheck(IP):
        try:
            os.rename(PATH + "\\" + IP +"\\data_viz.jsp", PATH + "\\" + IP + "\\throughput.csv")
            # 將下載後的檔案改名 方便讀取
            df = pd.read_csv(IP + "\\throughput.csv")
            mem = df[["Rtmmused", "Rtmmmemory"]]
            used = mem["Rtmmused"].values.tolist()
            total = mem["Rtmmmemory"].values.tolist()
            # 讀取csv檔中記憶體使用量以及記憶體大小
            mem_max = 0
            mem_min = 101
            for i in range(len(total)):
                value = round((used[i]/total[i]) * 100)
                mem_max = value if value > mem_max else mem_max
                mem_min = value if value < mem_min else mem_min
                # 計算每20分鐘的用量最大值及最小值，並記錄在mem_max與mem_min中

            if mem_max == mem_min:      # 將最大最小值以範圍的型態表示
                sys_mem = str(mem_min) + "%"
            else:
                sys_mem = str(mem_min) + "% ~ " + str(mem_max) + "%"
            break
        except :
            pass
# ============= cpu =============
    try:
        cpu = df[["Ruser", "Rniced","Rsystem","Ridle","Rirq","Rsoftirq","Riowait"]]     # 讀取計算cpu用量的參數
        used =[sum(item) for item in cpu[["Ruser", "Rniced","Rsystem"]].values.tolist()]
        total = [sum(item) for item in cpu.values.tolist()]
        # CPU 用量的計算方式:
        #                   user + niced + system
        # -----------------------------------------------------------  x 100%
        #   user + niced + system + idle + irq + softirq + iowait
        cpu_max = 0
        cpu_min = 101
        for i in range(len(total)):
            value = round((used[i]/total[i]) * 100)
            cpu_max = value if value > cpu_max else cpu_max
            cpu_min = value if value < cpu_min else cpu_min
            # 計算每20分鐘的用量最大值及最小值，並記錄在cpu_max與cpu_min中
        if cpu_max == cpu_min:      # 將最大最小值以範圍的型態表示
            sys_cpu = str(cpu_min) + "%"
        else:
            sys_cpu = str(cpu_min) + "% ~ " + str(cpu_max) + "%"
    except :
        logging.error("無法取得CPU用量 " + IP)
# ============= throughput =============
    try:
        tp = np.array(df["tput_bytes_in"].values.tolist())  # 讀取throughput的數據
        maxium = int(max(tp))
        minimum = int(min(tp))
        sys_tp = change_unit(minimum * 8) + " ~ " + change_unit(maxium * 8)
        # 抓取大最小值轉換單位後以範圍形式表示
    except:
        logging.error("無法取得 throughput " + IP)

    if healthCheck(IP):
        driver.close()
        return
# ============= active connection =============
    driver.get("https://" + IP + "/tmui/tmui/util/ajax/data_viz.jsp?cache=" + str(int(time())) + "&name=connections")
    # 抓取Statistic>>Performance>>Traffic Report背景傳輸的connection CSV檔
    while not healthCheck(IP):
        try:
            os.rename(PATH + "\\" + IP + "\\data_viz.jsp", PATH + "\\" + IP + "\\connections.csv")
            # 將下載後的檔案改名 方便讀取
            df = pd.read_csv(IP + "\\connections.csv")
            ac = np.array(df["curclientconns"].values.tolist())
            # 讀取active connection數據
            maxium = int(round(max(ac)))
            minimum = int(round(min(ac)))
            sys_ac = change_unit(minimum, "/sec") + " ~ " + change_unit(maxium, "/sec")
            # 抓取大最小值轉換單位後以範圍形式表示
            break
        except:
            pass
# ============= new connection =============
    try:
        nc = np.array(df["totclientconns"].values.tolist())     # 讀取new connection數據
        maxium = int(round(max(nc)))
        minimum = int(round(min(nc)))
        sys_nc = change_unit(minimum,"/sec") + " ~ " + change_unit(maxium,"/sec")
        # 抓取大最小值轉換單位後以範圍形式表示
    except:
        logging.error("無法取得new connection " + IP)

    if healthCheck(IP):
        driver.close()
        return
# ============= end =============
    # 讀取完資料收尾
    shutil.rmtree(IP, ignore_errors=True)                           # 將存放csv檔的資料夾刪除
    driver.close()                                                  # 關閉瀏覽器
    # thread_qkview.join()                                            # 等待qkview收集完畢
    # thread_ucs.join()                                               # 等待ucs收集完畢
    sys_qkview = "OK" if os.path.exists(PATH + "\\qkviews\\" + sys_host + '_' + now + ".qkview") else "ERROR"   # 若沒有收集到qkview則sys_qkview=ERROR
    sys_ucs = "OK" if os.path.exists(PATH + "\\ucs\\" + sys_host + '_' + now + ".ucs") else "ERROR"             # 若沒有收集到ucs則sys_ucs=ERROR

    thread_ltmLog.join()                                            # 等待ltm log 收集完畢
    thread_systemLog.join()                                         # 等待system log 收集完畢
    shutil.rmtree(PATH + "\\" + IP + "_log", ignore_errors=True)    # 將存放log檔的資料夾刪除

    outgo = [sys_host, sys_sn, sys_uptime, sys_mem, sys_cpu, sys_ac, sys_nc, sys_tp,
             sys_log, sys_ntp, sys_snmp, sys_ucs, sys_qkview, sys_time, sys_cert, sys_ha, sys_ver]
    writer.writerow(outgo)                                          # 儲存讀取到的資料
    print(IP + " 蒐集完畢") 
    pass_count += 1                                                 # 執行完畢計數 + 1


if __name__ == "__main__":
    global doc, t0, word_nn, filecount, pass_count
    # doc, t0, word_nn  --> 問henry 我複製過來的👍
    # filecount         --> 寫檔案用於文件編號
    # pass_count        --> 用於計算執行完畢數量
    pass_count = 0
    filecount = 0
    process_count = 0 # --> 計算同時進行的執行續數量
    devices = pd.read_excel("SMO_ex.xls").values.tolist()   # 讀取excel顯示的IP、帳號、密碼並儲存於devices中

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
        t = threading.Thread(target=get_data, args=(IP, ACCOUNT, PASSWD)) # 開啟一個執行續 設定get_data參數
        t.start()
        if process_count == 4:  # 限制同時搜尋4台設備
            t.join()
            process_count = 0   # 等待其中一台搜尋完畢重設process_count

    while(pass_count<=len(devices)): # 等待全部設備搜尋完畢再繼續
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