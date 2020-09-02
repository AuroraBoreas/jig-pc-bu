"""
===============================
This script is to control CS2000 SW to do optical measure continuously.
v0. @ZL, 20191118
v1. added Excel window handler to distinguish CA310 vba procedure and CS2000 log XL file.
      @ZL, 20191121
===============================
"""

import time, threading, pyautogui, pywinauto
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta
import sys, os
sys.path.append(os.path.dirname(__file__))
# from lib.pkg_cs2000_manager.drift import main_meas_IRE
from lib.pkg_cs2000_manager.core import CS2kMeasurer

def main_meas_IRE(mt, pmodsccno, color, ca_mode, user_psg500optime):
    dict_setup = dict(meas_time           = mt,
                      pmod_scc_no         = pmodsccno,
                      color               = color,
                      ca_mode             = ca_mode,
                      user_psg500_op_time = user_psg500optime,
                      ageing_hour         = 2)
    measurer1 = CS2kMeasurer(**dict_setup)
    measurer1.start()

def init():
    global s
    s = BackgroundScheduler()

def task_20IRE():
    """20IRE"""
    print(datetime.now(), "20IRE")
    threading.Thread(target=main_meas_IRE, args=(20, SCC_No, 'white', CA_Mode, User_PSG500_op_time)).start()
    
def task_100IRE():
    """100IRE"""
    print(datetime.now(), "100IRE")
    threading.Thread(target=main_meas_IRE, args=(12, SCC_No , 'white', CA_Mode, User_PSG500_op_time)).start()

if __name__ == '__main__':
    now = datetime.now()
    time_before_enter_loop = 3
    plan_start_time = now + timedelta(seconds=time_before_enter_loop) #<~ N secs from 'now', get a timestamp node

    SCC_No = 'SB49_Panel_SCC0000007_LS(DA04)'
    CA_Mode = False
    User_PSG500_op_time = 5
    n = 25 #<~ time gap between 100IRE, and 20IRE at a same timestamp
    init()

    info_fmt = '{0:%Y%m%d %H:%M:%S} Program is ready. \n {1} seconds later, start to measure'
    print(info_fmt.format(datetime.now(), time_before_enter_loop))
    for i in range(time_before_enter_loop):
        print(i+1)
        time.sleep(1)

    """===Paset codes of  scheduled node-times here if trying to restore measurement==="""
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=1.0))
    # s.add_job(task_20IRE, 'date', run_date=plan_start_time+timedelta(seconds=n)) #<~ time to measure 100IRE and swap PSG500, is like 20 seconds
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=120.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=300.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=600.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=1200.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=1800.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=3600.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=7200.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=14400.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=28800.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=43200.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=86400.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=172800.0))
    # s.add_job(task_20IRE, 'date', run_date=plan_start_time+timedelta(seconds=172800.0 + n)) #<~ time to measure 100IRE and swap PSG500, is like 20 seconds
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=259200.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=360000.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=604800.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=720000.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=900000.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=1080000.0))
    s.add_job(task_100IRE, 'date', run_date=plan_start_time+timedelta(seconds=1440000.0))
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
