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


sleep_time = 10

PATH = os.path.abspath(os.getcwd())
# try:
#     os.makedirs(IP + "_log")
# except Exception as e:
#     print(e)

def ltm(IP, ACC, PASS):
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(IP, username=ACC, password=PASS)
        _, stdout, _ = client.exec_command("cd /var/log; ls")

        lst = [line.replace('\n', '')
               for line in stdout.readlines() if line[:3] == "ltm"]
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
    # log = str(log)
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
    # if len(res) != 0:        if not os.path.exists(IP+"_ERR_LOG"): os.makedirs(IP+"_ERR_LOG") 
    #     with open(IP+"_ERR_LOG\\<這裡輸入錯誤名稱>_ERR.log", "a", newline='') as ef:
    #         ef.writelines(res)



def syst(IP, ACC, PASS):
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
    # log = str(log)


def cert(IP, ACC, PASS):
    today = datetime.today()
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
    except Exception as e:
        logging.error("無法取得憑證資訊 " + IP + str(e))

    driver.close()


def getTime(client):
    try:
        start = time()
        _, stdout, _ = client.exec_command("date")
        end = time()

        local_time = datetime.strftime(datetime.now(), "%H:%M:%S")

        sys_time_hr, sys_time_min, sys_time_sec = stdout.readlines()[0].split(' ')[4].split(':')
        

        loc_time_hr = int(local_time.split(':')[0])
        loc_time_min = int(local_time.split(':')[1])
        loc_time_sec = int(local_time.split(':')[2])

        diff_hr = int(sys_time_hr) - loc_time_hr
        diff_min = int(sys_time_min) - loc_time_min
        diff_sec = int(sys_time_sec) - loc_time_sec

        diff_time = diff_sec + diff_min * 60 + diff_hr * 3600

        if diff_time > 0:
            sys_time = "快" + (str(diff_hr) + "小時") * int(diff_hr != 0) + (str(diff_min) + "分鐘") * int(diff_min != 0) + str(diff_sec) + "秒"
        elif diff_time < 0:
            sys_time = "慢" + (str(abs(diff_hr)) + "小時") * int(diff_hr != 0) + (str(abs(diff_min)) + "分鐘") * int(diff_min != 0) + str(diff_sec) + "秒"
        else:
            sys_time = "OK"
        print(end - start)
        print(sys_time)
    except Exception as e:
        print(e)
        return


if __name__ == "__main__":
    # cert(IP, ACC, PASS)
    # t1 = threading.Thread(target=ltm, args=(IP, ACC, PASS))
    # t2 = threading.Thread(target=syst, args=(IP, ACC, PASS))
    # t1.start()
    # t2.start()

    # t1.join()
    # t2.join()
    # shutil.rmtree(PATH + "\\" + IP + "_log", ignore_errors=True)
    IP = "192.168.51.192"
    ACC = "admin"
    PASS = "admin"
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(IP, username=ACC, password=PASS)

    getTime(client)