#!/usr/bin/env python
# coding: utf-8

# In[1]:


"""
============================================================================================
This is an automation project developed by ZL, 20191105
This project tries to automate Optical measurement software[cas25w] w/o an exposed API.

#v0, ZL, 20191105. Project was starting, developing, and debugging
#v1, ZL, 20191106. Polishing. Added a logic/function to detect if data name duplicates than twice.
#v2, ZL, 20200424. breakthrough 24 rows. It can fetch ultimate rows from CA2500


WARNING! 
=sys info=
system      : 'Windows'
node         : 'adminpanel-Desi'
release      : '7'
version      : '6.1.7601'
machine    : 'x86'
processor : 'x86 Family 6 Model 61 Stepping 4, GenuineIntel'

=screen=
size            : 'Size(width=1366, height=768)'

=CA-25w=
version     : 'Ver.1.00.0005'

=Excel=
version     : '14.0.4760.1000(32bit)'

============================================================================================
"""

import pyautogui, time, win32clipboard, sys
#import numpy as np
from io import BytesIO

pyautogui.PAUSE = 0.15
pyautogui.FAILSAFE = True


# In[2]:


pyautogui.position()


# In[3]:


def clear_clipboard():
    """Empty clipboard"""
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.CloseClipboard()
    
def get_clipboard_data():
    """get data from clipboard"""
    win32clipboard.OpenClipboard()
    d = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()
    return d

def choose_data_set_for_tab_swap():
    """
    this is a workaround to settle tab swapping between 参考点 and 仿真色 beforehand
    """
    pyautogui.moveTo(x=18, y=215)
    pyautogui.click()
    time.sleep(1.5)

def change_cas25w_layout(is_data=True):
    """swatch tabs between 参考点 and 仿真色"""
    #back to cas25w main window
    if is_data:      
        #<~ select tab "参考点"
        pyautogui.moveTo(x=548, y=108)
        pyautogui.click()
        time.sleep(1)
    else:
        #<~ select tab "仿真色"
        pyautogui.moveTo(x=411, y=109)
        pyautogui.click()
        time.sleep(1)
        
def get_data_name(date_name_coor):
    """get name of data set"""
    #! attention: always presume cas25w window is active 
    pyautogui.moveTo(*date_name_coor)
    pyautogui.click()
    time.sleep(2)
    pyautogui.doubleClick()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    win32clipboard.OpenClipboard()
    d = win32clipboard.GetClipboardData()
#     win32clipboard.EmptyClipboard()
    win32clipboard.CloseClipboard()
    return d

def verify_data_name_data(dn):
    """verify data name: string contains 100IRE or 50IRE or CENTRE
    U may definetly added more conditions here 
    expect data set: 100IRE, 50IRE @ZL
    """
    identifiers = ['100IRE', '50IRE']
    identifier_center = 'ZOOM'
    identifier_SET = 'SET'
    dn = dn.upper()
    if identifier_center in dn or identifier_SET in dn:
        return False
    else:
        return any(i in dn for i in identifiers)

def verify_data_name_img(dn):
    """verify data name: string contains 100IRE or 50IRE or 0IRE or CENTRE
    U may definetly added more conditions here
    Expect data set: 100IRE, 50IRE, 0IRE @ZL
    """
    identifiers = ['100IRE', '50IRE','0IRE', '87IRE']
    identifier_center = 'ZOOM'
    identifier_mura = 'MURA'
    #identifier_SET = 'SET'
    dn = dn.upper()
    if identifier_center in dn:
        return False
    else:
        return any(i in dn for i in identifiers)

def transfer_data_from_cas250w_to_xl(x, y):
    """transfer data via cas250w -> ctrl + c to copy -> Excel -> ctrl + v to paste"""
    #activate Exel
    pyautogui.moveTo(x=332, y=743)
    pyautogui.click()
    #active destine cell
    pyautogui.moveTo(x=x, y=y)
    pyautogui.click()
    #store data
    pyautogui.hotkey('ctrl', 'v')
    #back to cas25w main window
    pyautogui.moveTo(x=274, y=747)
    pyautogui.click()
    #<~ clear clipboard
    clear_clipboard()

def transfer_data_name(from_coor1, to_coor1):
    """transfer data name from cas25w to Excel
    from_coor1 is coordinate(x, y) at cas25w
    to_coor1 is coodinate(x, y) at Excel
    """
    #! attention: always presume cas25w window is active 
    pyautogui.moveTo(*from_coor1)
    pyautogui.click()
    time.sleep(1)
    pyautogui.doubleClick()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    transfer_data_from_cas250w_to_xl(*to_coor1)

def transfer_dot_data(from_coor1, to_coor1):
    """
    similar as the function above, but for its dot date
    """
    #! attention: always presume cas25w window is active 
    pyautogui.moveTo(x=71, y=133)
    pyautogui.click()
    time.sleep(1)
    #open position setting
    pyautogui.moveTo(x=1128, y=79)
    pyautogui.click()
    time.sleep(2) #<~delay 1s to wait for small window popout
    #select position setting1
    pyautogui.moveTo(*from_coor1) ##<~ here is to be dynamic
    pyautogui.doubleClick()
    time.sleep(2) #<~delay 1s to wait for cas25w response
    #confirm and back to main window
    pyautogui.moveTo(x=1008, y=671)
    pyautogui.click()
    #move to copy data button
    pyautogui.moveTo(x=685, y=142)
    pyautogui.click()
    #copy dot data
    pyautogui.moveTo(x=678, y=252)
    pyautogui.click()
    #paste into Excel
    transfer_data_from_cas250w_to_xl(*to_coor1)

def scroll_xl_window_down():
    """
    to suit Excel format
    """
    #activate Exel
    pyautogui.moveTo(x=332, y=743)
    pyautogui.click()
    #<~ Position: click Excel scrollbar down button
    pyautogui.moveTo(x=1353, y=682)
    for _ in range(17):
        pyautogui.click()
        time.sleep(0.2)
    # back to cas25w main window
    pyautogui.moveTo(x=274, y=747)
    pyautogui.click()

def send_to_clipboard(clip_type, data):
    """
    put something on clipboard
    """
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()
    
def shoot_img(x1,y1,x2,y2):
    """
    take screenshot -> convert to binary data -> put on clipboard
    """
    w = x2 - x1
    h = y2 - y1
    img = pyautogui.screenshot(region=(x1,y1,w,h))
    output = BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    send_to_clipboard(win32clipboard.CF_DIB, data)

def transfer_img(to_coor1):
    """
    UF img to Excel
    """
    #! attention: always presume cas25w window is active 
    #shoot_img
    shoot_img(417,322,824,551)
    time.sleep(1)
    #paste into Excel
    transfer_data_from_cas250w_to_xl(*to_coor1)

def switch_color():
    """
    switch between colorful and gray 
    """
    pyautogui.moveTo(x=845, y=142)
    pyautogui.click()
    time.sleep(0.5)

def scroll_xl_to_right(n=6):
    """
    to suit Excel format
    """
    #<~ Position: click Excel scrollbar down button
    pyautogui.moveTo(x=1328, y=697)
    for _ in range(n):
        pyautogui.click()
        time.sleep(0.6)
    # back to cas25w main window
    pyautogui.moveTo(x=274, y=747)
    pyautogui.click()

def activate_src_ws():
    """activate src worksheet.
    Attention: Excel worksheet names and positions matter!
    """
    #activate Exel
    pyautogui.moveTo(x=332, y=743)
    pyautogui.click()
    #select src ws
    pyautogui.moveTo(x=80, y=698)
    pyautogui.click()
    
def activate_img_ws():
    """activate img worksheet.
    Attention: Excel worksheet names and positions matter!
    """
    #activate Exel
    pyautogui.moveTo(x=332, y=743)
    pyautogui.click()
    #select img ws cell
    pyautogui.moveTo(x=117, y=699)
    pyautogui.click()
    
def activate_cas25w():
    """activate cas25w.
    cas25w on windows bar matters!
    """
    # back to cas25w main window
    pyautogui.moveTo(x=274, y=747)
    pyautogui.click()  
    
def reset_window():
    """
    click on windows task bar. this action put all open software with UI background
    """
    pyautogui.moveTo(x=1059, y=744)
    pyautogui.click()    

def make_vba_report():
    """activate Excel and click prefabricated vba module button"""
    pyautogui.moveTo(x=274, y=747)
    pyautogui.click()
    pyautogui.moveTo(x=164, y=11)
    pyautogui.click()
    pyautogui.moveTo(x=947, y=467)
    pyautogui.click()

def get_rgb_under_mouse_cursor(x, y):
    """
    return r,g,b value of pixel under mouse cursor (x=337, y=191)
    # 246 247 249, no vertical scroll-bar
    #(46, 151, 207), when ca2500 has vertical scroll-bar
    """
    from ctypes import windll
    pyautogui.moveTo(x, y) #color changes when mouse event occurs
    dc = windll.user32.GetDC(0)
    rgb = windll.gdi32.GetPixel(dc, x, y)
    r = rgb & 0xff
    g = (rgb >> 8) & 0xff
    b = (rgb >> 16) & 0xff
    return (r, g, b)

def move_to_slidebar_downarrow_and_click():
    """when N(sample) is greater than 24
    move to the fix row to get data.  @ZL, 20200424
    Point(x=334, y=683)"""
    pyautogui.moveTo(x=334, y=683)
    pyautogui.click()
    time.sleep(0.5)

def reset_cas25w_GUI_slidebar():
    """reset GUI slidebar
    move to the fix row. @ZL, 20200424
    Point(x=337, y=191)"""
    x, y = 337, 191
    if abs(get_rgb_under_mouse_cursor(x, y)[0] - 46) <= 5:
        pyautogui.moveTo(x, y)
        for _ in range(100):
            pyautogui.click()
    else:
        pass
    
# In[4]:


def main_get_data(n):
    STEP = 19 #<~ cas25w_data_name_row_step
    reset_window()
    
    ###<~ transfer dataname from cas25w to Excel: #<~well, push to limit: max 25 sets of data
    activate_src_ws()#<~ active SRC worksheet in Excel
    activate_cas25w()
    reset_cas25w_GUI_slidebar()
    choose_data_set_for_tab_swap()
    change_cas25w_layout(is_data=True) #<~select tab "参考点"
    
    tmp = ""
    cnt = 0

    for i in range(n):       
        if i <= 24:
            cas25w_data_name_coor = (18, 215 + i * STEP)
            try:
                dn = get_data_name(cas25w_data_name_coor)
                if tmp == dn:
                    cnt += 1
                if cnt >= 1: #<~~ break if data name duplicates about twice
                    break
                if verify_data_name_data(dn):
                    # data set name
                    xl_paste_coor = (55, 121) #<~ paste into Excel file
                    transfer_data_name(cas25w_data_name_coor, xl_paste_coor)
                    #<~ get data of 01_UF9 POINT_1_18.evl
                    cas25w_setting_coor = (559,251)
                    xl_paste_coor = (141, 139)
                    transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
                    #<~ get data of 02_UF9 POINT_60_40.evl
                    cas25w_setting_coor = (559,271)
                    xl_paste_coor = (427, 139)
                    transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
                    #<~ get data of 03_UF9 POINT_1_6.evl
                    cas25w_setting_coor = (559,291)
                    xl_paste_coor = (716, 139)
                    transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
                    #<~ setting up Excel before proceeding with next data set
                    scroll_xl_window_down()            
            except TypeError: #<~ stops when hitting blank data row
                break
            tmp = dn

        if i > 24:
            move_to_slidebar_downarrow_and_click()
            cas25w_data_name_coor = (18, 671) #<- fixed row, height = 671
            try:
                dn = get_data_name(cas25w_data_name_coor)
                if tmp == dn:
                    cnt += 1
                if cnt >= 1: #<~~ break if data name duplicates about twice
                    break
                if verify_data_name_data(dn):
                    # data set name
                    xl_paste_coor = (55, 121) #<~ paste into Excel file
                    transfer_data_name(cas25w_data_name_coor, xl_paste_coor)
                    #<~ get data of 01_UF9 POINT_1_18.evl
                    cas25w_setting_coor = (559,251)
                    xl_paste_coor = (141, 139)
                    transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
                    #<~ get data of 02_UF9 POINT_60_40.evl
                    cas25w_setting_coor = (559,271)
                    xl_paste_coor = (427, 139)
                    transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
                    #<~ get data of 03_UF9 POINT_1_6.evl
                    cas25w_setting_coor = (559,291)
                    xl_paste_coor = (716, 139)
                    transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
                    #<~ setting up Excel before proceeding with next data set
                    scroll_xl_window_down()            
            except TypeError: #<~ stops when hitting blank data row
                break
            tmp = dn

# In[5]:


def main_get_img(n):
    STEP = 19 #<~ cas25w_data_name_row_step
    reset_window()
    
    ###<~ transfer dataname from cas25w to Excel: #<~well, push to limit: max 25 sets of data
    activate_cas25w()
    reset_cas25w_GUI_slidebar()
    choose_data_set_for_tab_swap()
    change_cas25w_layout(is_data=False) #<~select tab "仿真色"
    #<~ active Img worksheet in Excel
    activate_img_ws()
    scroll_xl_to_right(n=1) #<~ match Img format in VBA

    tmp = ""
    cnt = 0
    for i in range(n):
        if i <= 24:
            cas25w_data_name_coor = (18, 215 + STEP * i)
            try:
                dn = get_data_name(cas25w_data_name_coor)
                if tmp == dn:
                    cnt += 1
                if cnt >= 1: #<~~ break if data name duplicates about twice
                    break
                if verify_data_name_img(dn):
                    # data set name
                    xl_paste_coor = (62, 121) #<~ paste into Excel file
                    transfer_data_name(cas25w_data_name_coor, xl_paste_coor)
                    #<~ get img: colorful        
                    xl_paste_coor = (62, 139) #<~ it supposes to start with colorful!
                    transfer_img(xl_paste_coor)
                    #<~ get img: gray
                    switch_color()
                    xl_paste_coor = (62, 373)
                    transfer_img(xl_paste_coor)
                    activate_img_ws()
                    scroll_xl_to_right(n=6)
                    switch_color()
            except TypeError:
                break
            tmp = dn

        if i > 24:
            move_to_slidebar_downarrow_and_click()
            cas25w_data_name_coor = (18, 671) #<- fixed row, height = 671            
            try:
                dn = get_data_name(cas25w_data_name_coor)
                if tmp == dn:
                    cnt += 1
                if cnt >= 1: #<~~ break if data name duplicates about twice
                    break
                if verify_data_name_img(dn):
                    # data set name
                    xl_paste_coor = (62, 121) #<~ paste into Excel file
                    transfer_data_name(cas25w_data_name_coor, xl_paste_coor)
                    #<~ get img: colorful        
                    xl_paste_coor = (62, 139) #<~ it supposes to start with colorful!
                    transfer_img(xl_paste_coor)
                    #<~ get img: gray
                    switch_color()
                    xl_paste_coor = (62, 373)
                    transfer_img(xl_paste_coor)
                    activate_img_ws()
                    scroll_xl_to_right(n=6)
                    switch_color()
            except TypeError:
                break
            tmp = dn

# In[6]:


def main():
    n = 100
    #clear_clipboard()
    #main_get_data(n)
    #time.sleep(1)
    clear_clipboard()
    main_get_img(n)
    time.sleep(1)
    #make_vba_report()

if __name__ == "__main__":
    main()


# In[ ]:




