#!/usr/bin/env python
# coding: utf-8

# Try to auto CS2000 SW to do meas. <br>
# @ZL, 20191115 <br>
# 
# **basic env. info**
# - Lang: Python
# - IDEL: Jupyter
# - Version: 3.7
# - System: Win7

# In[118]:


import pyautogui, time, win32gui, win32con, psutil, pyttsx3
from datetime import datetime
from pywinauto.findwindows import find_window
from win32con import (SW_SHOW, SW_RESTORE)

pyautogui.PAUSE = 0.25


# In[119]:


def get_windows_placement(window_id):
    return win32gui.GetWindowPlacement(window_id)[1]

def set_active_window(window_id):
    if get_windows_placement(window_id) == 2:
        win32gui.ShowWindow(window_id, SW_RESTORE)
    else:
        win32gui.ShowWindow(window_id, SW_SHOW)
        
    win32gui.SetForegroundWindow(window_id)
    win32gui.SetActiveWindow(window_id)

def close_excel_window():
    PROCNAME = "EXCEL.EXE"
    for proc in psutil.process_iter():
        if proc.name() == PROCNAME:
            proc.kill()


# In[120]:


def reset_ui_windows():
    """click nothing on task bar"""
    pyautogui.moveTo(x=1056, y=748)
    pyautogui.click()
    
def click_cs2000sw_on_taskbar():
    """click icon on task bar"""
    pyautogui.moveTo(x=291, y=749)
    pyautogui.click()    

def find_excel_on_taskbar():
    """find excel and bring to top"""
    window_id = find_window(title='Microsoft Excel - Sheet1')
    set_active_window(window_id)
    
def click_cs2000sw_saveToExcel():
    """save to excel"""
    pyautogui.moveTo(x=928, y=530)
    pyautogui.click()
    
def cs2000_sw_meas(meas_time, IRE, Check_All_Mode=False, User_Change_PSG500_RGBB_time=10):
    engine = pyttsx3.init()
    
    """meas and save"""
    if not Check_All_Mode: #<~ meas white only
        engine.say("Start to measure momentarily. {}IRE".format(IRE))
        engine.runAndWait()

        #make cs2000sw as top window
        reset_ui_windows()#reset ui windows
        click_cs2000sw_on_taskbar()
        
        #<~ meas
        pyautogui.moveTo(x=933, y=422)
        pyautogui.click() 
        pyautogui.hotkey('Enter')
        time.sleep(meas_time) #<~ waiting for cs2000 meas, N seconds
        pyautogui.hotkey('Enter')
        #<~ save to excel
        click_cs2000sw_saveToExcel()
        time.sleep(1) #<~ waiting for excel operation
        pyautogui.hotkey('Enter')
        time.sleep(1.5)
        engine.say("Record successfully")
        engine.runAndWait()
      
    if Check_All_Mode: #<~ WRGBB
        engine.say("Start to measure momentarily, WRGBB mode")
        engine.runAndWait()
        
        #make cs2000sw as top window
        reset_ui_windows()#reset ui windows
        click_cs2000sw_on_taskbar()
        
        ##<~ meas
        # White
        engine.say("Start to measure White, {} seconds later".format(User_Change_PSG500_RGBB_time))
        engine.runAndWait()
        time.sleep(User_Change_PSG500_RGBB_time) #<~  wait for user to change PSG500
        pyautogui.moveTo(x=933, y=422)
        pyautogui.click() 
        pyautogui.hotkey('Enter')
        time.sleep(12) #<~ waiting for cs2000 meas, N seconds
        # Red
        engine.say("Start to measure Red, {} seconds later".format(User_Change_PSG500_RGBB_time))
        engine.runAndWait()
        time.sleep(User_Change_PSG500_RGBB_time+3) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter')
        time.sleep(5) #<~ waiting for cs2000 meas, N seconds
        # Green
        engine.say("Start to measure Green, {} seconds later".format(User_Change_PSG500_RGBB_time))
        engine.runAndWait()
        time.sleep(User_Change_PSG500_RGBB_time+3) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter') 
        time.sleep(5) #<~ waiting for cs2000 meas, N seconds
        # Blue
        engine.say("Start to measure Blue, {} seconds later".format(User_Change_PSG500_RGBB_time))
        engine.runAndWait()
        time.sleep(User_Change_PSG500_RGBB_time+3) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter')
        time.sleep(5) #<~ waiting for cs2000 meas, N seconds
        # Black
        engine.say("Start to measure Black, {} seconds later".format(User_Change_PSG500_RGBB_time))
        engine.runAndWait()
        time.sleep(User_Change_PSG500_RGBB_time) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter')
        time.sleep(52) #<~ waiting for cs2000 meas, N seconds
        pyautogui.hotkey('Enter')
        
        #<~ save to excel
        click_cs2000sw_saveToExcel()
        time.sleep(1.5) #<~ waiting for excel operation
        pyautogui.hotkey('Enter')
        time.sleep(1.5)
        engine.say("Record successfully")
        engine.runAndWait()     

def save_excel_data_to_local_documents_folder(file_name):
    """save cs2000 log data"""
    #make excel as top window
    find_excel_on_taskbar()
    pyautogui.hotkey('ALT','F','S')   
    time.sleep(3)
    pyautogui.typewrite(file_name)
    time.sleep(3)
    pyautogui.hotkey('Enter')
    time.sleep(3)
    close_excel_window()


# In[121]:

def main_meas_IRE(meas_time, SCC_No, IRE, CA_Mode, User_PSG500_op_time, Ageing_Hour):
    cs2000_sw_meas(meas_time, IRE, Check_All_Mode=CA_Mode, User_Change_PSG500_RGBB_time=User_PSG500_op_time)
    file_name = datetime.now().strftime('%Y-%m-%d %H%M%S') + " {0}_{1}H".format(SCC_No, Ageing_Hour)
    file_name = file_name.upper()
    save_excel_data_to_local_documents_folder(file_name)

if __name__ == '__main__':
    meas_time = 12
    SCC_No = 'SCC9300101'
    Ageing_Hour = 100
    CA_Mode = False
    User_PSG500_op_time = 10
    Ageing_Hour = 2
    main_meas_IRE(meas_time, SCC_No, Ageing_Hour, CA_Mode, User_PSG500_op_time, Ageing_Hour)

