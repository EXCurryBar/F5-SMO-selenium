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
        with open(IP + "_HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*No failover status messages received for.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*Active\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*Offline\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)

    P = re.compile("\n.*Standby\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_HA_ERR.log", "a", newline='') as ef:
            ef.writelines(res)
# == VS state change
    P = re.compile("\n.*Virtual Address .*GREEN to RED.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_VS_ERR.log", "w", newline='') as ef:
            ef.writelines(res)
# == Pool
    P = re.compile("\n.*Pool.*status down.*\n")
    res = re.findall(P, log)
    if len(res) != 0:
        with open(IP + "_Pool_ERR.log", "w", newline='') as ef:
            ef.writelines(res)

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

if __name__ == "__main__":
    # cert(IP, ACC, PASS)
    t1 = threading.Thread(target=ltm)
    t2 = threading.Thread(target=syst)
    t1.start()
    t2.start()

    t1.join()
    t2.join()
    shutil.rmtree(PATH + "\\" + IP + "_log", ignore_errors=True)
