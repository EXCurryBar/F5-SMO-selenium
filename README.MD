# F5 SMO 表單自動化填寫
這支程式是用來填寫F5的維護表單，已達到減少人力成本以及保留維護紀錄。
### 環境
開發程式時使用 Python 3.8.5 ，安裝時可以下載zip，或是安裝git並輸入下面指令
    
    git clone https://github.com/EXCurryBar/F5-SMO-selenium.git
使用到的函示庫如下:

    matplotlib == 3.3.1
    numpy == 1.18.5
    pandas == 1.1.2
    paramiko == 2.7.2
    scp == 0.13.2
    selenium == 3.141.0
需要安裝函式庫時運行:

    pip install -r requirement.txt
即可一次安裝

### 已知問題

    syslog : 
        1.無判斷system log錯誤
        2.無法按照log時間排序

