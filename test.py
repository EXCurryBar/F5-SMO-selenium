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
IP = "192.168.51.192"
ACC = "admin"
PASS = "admin"
try:
    os.makedirs(IP + "_log")
except Exception as e:
    print(e)


def ltm():
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
    with open(IP+"_log\\ltm.log", "w", encoding='UTF-8') as lf:
        lf.write(log)
    log = ""
    with open(IP+"_log\\ltm.log", "r", encoding='UTF-8') as lf:
        log = lf.read()

    # == HA state change
    P = re.compile("\n.*HA unit.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_HA_ERR.log", "w", newline='') as ef:
            ef.writelines(res)

    # == VS state change
    P = re.compile("\n.*Virtual Address .*GREEN to RED.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_VS_ERR.log", "w", newline='') as ef:
            ef.writelines(res)

    # == Pool monitor down
    # P = re.compile("\n.*Pool.*status down.*\n")
    # res = re.findall(P, log)
    # if len(res) != 0:
    #     with open(IP + "_Pool_ERR.log", "w", newline='') as ef:
    #         ef.writelines(res)


def syst():
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
    with open(IP+"_log\\messages.log", "w", encoding='UTF-8') as lf:
        lf.write(log)
    log = ""
    with open(IP+"_log\\messages.log", "r", encoding='UTF-8') as lf:
        log = lf.read()


if __name__ == "__main__":
    options = webdriver.ChromeOptions()
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
            sys_time = "快" + (str(diff_hr) + "小時") * diff_hr + str(diff_min) + "分鐘"
        elif diff_time < 0:
            sys_time = "慢" + (str(abs(diff_hr)) + "小時") * abs(diff_hr) + str(abs(diff_min)) + "分鐘"
        else:
            sys_time = "OK"
        
    except:
        logging.error("無法獲取時間資訊 " + IP)
        sys_time = "ERROR"

    print(sys_time)





    driver.close()