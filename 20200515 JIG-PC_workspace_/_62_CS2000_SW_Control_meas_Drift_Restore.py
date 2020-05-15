"""
===============================
This script is to control CS2000 SW to do optical measure continuously.
v0. @ZL, 20191118
v1. added Excel window handler to distinguish CA310 vba procedure and CS2000 log XL file.
      @ZL, 20191121
===============================
"""

import time, threading, pyautogui
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta
from CS2000_package.CS2000_SW_Control import main_meas_IRE

def init():
    global s
    s = BackgroundScheduler()

def task_20IRE():
    """20IRE"""
    print(datetime.now(), "20IRE")
    threading.Thread(target=main_meas_IRE, args=(20, SCC_No, '20', CA_Mode, User_PSG500_op_time)).start()
    
def task_100IRE():
    """100IRE"""
    print(datetime.now(), "100IRE")
    threading.Thread(target=main_meas_IRE, args=(12, SCC_No , '100', CA_Mode, User_PSG500_op_time)).start()

if __name__ == '__main__':
    now = datetime.now()
    time_before_enter_loop = 3
    plan_start_time = now + timedelta(seconds=time_before_enter_loop) #<~ N secs from 'now', get a timestamp node

    SCC_No = 'NX85_CS_SCC9300101'
    CA_Mode = False
    User_PSG500_op_time = 5
    n = 25 #<~ time gap between 100IRE, and 20IRE at a same timestamp
    init()

    print('{} Program is ready. \n {} seconds later, start to measure'.format(datetime.now(),time_before_enter_loop))
    for i in range(time_before_enter_loop):
        print(i+1)
        time.sleep(1)

    """===Paset codes of  scheduled node-times here if trying to restore measurement==="""
    #paste codes of node-times generated from _65_XL_node-time_for_Driftpy.py
    """===Paset codes of  scheduled node-times here if trying to restore measurement==="""

    try:
        s.start()
        while True:
            time.sleep(2)
    except(KeyboardInterrupt, SystemExit):
        s.shutdown()
    except pywinauto.findwindows.WindowNotFoundError:
        s.shutdown()

    print('job done')
