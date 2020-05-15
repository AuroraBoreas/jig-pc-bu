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


import pyautogui, time, win32gui, win32con, psutil
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

#def close_excel_window():
#    PROCNAME = "EXCEL.EXE"
#   for proc in psutil.process_iter():
#        if proc.name() == PROCNAME:
#            proc.kill()

def close_excel_window(title_name):
    window_id = find_window(title=title_name)
    win32gui.PostMessage(window_id, win32con.WM_CLOSE,0,0)
            

# In[120]:


def reset_ui_windows():
    """click nothing on task bar"""
    pyautogui.moveTo(x=1056, y=748)
    pyautogui.click()
    
def click_cs2000sw_on_taskbar():
    """click icon on task bar"""
    pyautogui.moveTo(x=291, y=749)
    pyautogui.click()    

def find_excel_on_taskbar(title_name ='Microsoft Excel - Sheet1'):
    """find excel and bring to top"""
    window_id = find_window(title=title_name)
    set_active_window(window_id)
    
def click_cs2000sw_saveToExcel():
    """save to excel"""
    pyautogui.moveTo(x=928, y=530)
    pyautogui.click()
    
def cs2000_sw_meas(meas_time, IRE, Check_All_Mode=False, User_Change_PSG500_RGBB_time=10):
    
    """meas and save"""
    if not Check_All_Mode: #<~ meas white only
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
      
    if Check_All_Mode: #<~ WRGBB       
        #make cs2000sw as top window
        reset_ui_windows()#reset ui windows
        click_cs2000sw_on_taskbar()
        
        ##<~ meas
        # White
        time.sleep(User_Change_PSG500_RGBB_time) #<~  wait for user to change PSG500
        pyautogui.moveTo(x=933, y=422)
        pyautogui.click() 
        pyautogui.hotkey('Enter')
        time.sleep(12) #<~ waiting for cs2000 meas, N seconds
        # Red
        time.sleep(User_Change_PSG500_RGBB_time+3) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter')
        time.sleep(5) #<~ waiting for cs2000 meas, N seconds
        # Green
        time.sleep(User_Change_PSG500_RGBB_time+3) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter') 
        time.sleep(5) #<~ waiting for cs2000 meas, N seconds
        # Blue
        time.sleep(User_Change_PSG500_RGBB_time+3) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter')
        time.sleep(5) #<~ waiting for cs2000 meas, N seconds
        # Black
        time.sleep(User_Change_PSG500_RGBB_time) #<~  wait for user to change PSG500
        pyautogui.hotkey('Enter')
        time.sleep(60) #<~ waiting for cs2000 meas, N seconds
        pyautogui.hotkey('Enter')
        
        #<~ save to excel
        click_cs2000sw_saveToExcel()
        time.sleep(1) #<~ waiting for excel operation
        pyautogui.hotkey('Enter')
        time.sleep(1.5)

def save_excel_data_to_desktop(file_name):
    """save cs2000 log data"""
    #make excel as top window
    find_excel_on_taskbar(title_name ='Microsoft Excel - Sheet1')
    pyautogui.hotkey('ALT','F','S')   
    time.sleep(3)
    pyautogui.typewrite(file_name)
    time.sleep(3)
    pyautogui.hotkey('Enter')
    time.sleep(1)
    title_name = 'Microsoft Excel - {}.xlsx'.format(file_name)
    close_excel_window(title_name)


# In[121]:

def main_meas_IRE(meas_time, SCC_No, IRE, CA_Mode, User_PSG500_op_time):
    cs2000_sw_meas(meas_time, IRE, Check_All_Mode=CA_Mode, User_Change_PSG500_RGBB_time=User_PSG500_op_time)
    file_name = datetime.now().strftime('%Y-%m-%d %H%M%S') + " {0}_cs2000_meas_log_{1}IRE".format(SCC_No, IRE) 
    save_excel_data_to_desktop(file_name)

if __name__ == '__main__':
    meas_time = 12
    SCC_No = 'SCC9300101'
    IRE = 100
    CA_Mode = False
    User_PSG500_op_time = 10
    main_meas_IRE(meas_time, SCC_No, IRE, CA_Mode, User_PSG500_op_time)

