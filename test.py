import paramiko
from scp import SCPClient
import os
import re
import gzip
import shutil
import threading

PATH = os.path.abspath(os.getcwd())
IP = "192.168.51.160"
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
    t1 = threading.Thread(target=ltm)
    t2 = threading.Thread(target=syst)
    t1.start()
    t2.start()

    t1.join()
    t2.join()
    shutil.rmtree(PATH + "\\" + IP + "_log", ignore_errors=True)

    d = os.listdir()
    for item in d:
        if item[:len(IP)] == IP and item[-7:] == "ERR.log":
            print(item)