import os
import time
import datetime
import threading
from jkdk.dk import fuc_dk
from zhs.zhs import zhs_main

sign = 1  # 结束标志

if __name__ == "__main__":
    path_dk = os.getcwd() + '\\jkdk\\dk.xlsx'
    path_zhs = os.getcwd() + '\\zhs\\zhs.xlsx'
    now = datetime.datetime.now()
    dk = datetime.datetime.strptime(str(now.year) + "-" + str(now.month) + "-" + str(now.day) + " 07:10:00", "%Y-%m-%d "
                                                                                                             "%H:%M:%S")
    zhs = datetime.datetime.strptime(str(now.year) + "-" + str(now.month) + "-" + str(now.day) + " 13:00:00",
                                     "%Y-%m-%d "
                                     "%H:%M:%S")
    f = open('main_timestamp.txt', 'r')
    time_stamp = f.readlines()  # 读取上次运行时间
    dk_0 = datetime.datetime.strptime(time_stamp[0].replace('\n', ''), "%Y-%m-%d %H:%M:%S")  # 文本=>datetime
    zhs_0 = datetime.datetime.strptime(time_stamp[1], "%Y-%m-%d %H:%M:%S")
    f.close()
    while sign:
        now = datetime.datetime.now()
        dk_time = (dk - now).total_seconds()
        zhs_time = (zhs - now).total_seconds()
        if (dk - dk_0).days:  # 当 上次运行时间间隔>一天 时执行
            dk_save = now
            if dk_time > 0:
                threading.Timer(dk_time, fuc_dk, path_dk).start()
            else:
                fuc_dk(path_dk)
        if (zhs - zhs_0).days:
            sign = 0
            zhs_save = now
            if zhs_time > 0:
                threading.Timer(zhs_time, zhs_main, path_zhs).start()
            else:
                threading.Thread(target=zhs_main,args=path_zhs).start()
        time.sleep(600)
    sign = 1
    while sign:
        time.sleep(600)
    f = open('main_timestamp.txt', 'w')
    f.write(str(dk))
    f.write('\n' + str(zhs))
    f.close()
