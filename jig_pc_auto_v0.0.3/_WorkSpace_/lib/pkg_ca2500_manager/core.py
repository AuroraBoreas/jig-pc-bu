#!/usr/bin/env python
# coding: utf-8

"""============================================================================================
This is an automation project developed by ZL, 20191105
This project tries to automate Optical measurement software[cas25w] w/o an exposed API.

#v0, ZL, 20191105. Project was starting, developing, and debugging
#v1, ZL, 20191106. Polishing. Added a logic/function to detect if data name duplicates than twice.
#v2, ZL, 20200424. breakthrough 24 rows. It can fetch ultimate rows from CA2500
#v3, ZL, 20200527. refactor, module

WARNING, DEV ENV! 
=sys info=
system   : 'Windows'
node     : 'adminpanel-Desi'
release  : '7'
version  : '6.1.7601'
machine  : 'x86'
processor: 'x86 Family 6 Model 61 Stepping 4, GenuineIntel'
=screen=
size     : 'Size(width=1366, height=768)'
=CA-25w=
version  : 'Ver.1.00.0005'
=Excel=
version  : '14.0.4760.1000(32bit)'
============================================================================================"""

import pyautogui, time, win32clipboard
from io import BytesIO
from .point import Point

pyautogui.PAUSE = 0.15
pyautogui.FAILSAFE = True

class CA2500_Extractor(object):
    #application positions on taskbar
    ca25_position             = Point(description="ca25 position on taskbar", x=274, y=743)
    excel_position            = Point(description="excel position on taskbar", x=332, y=743)
    # reset_position_on_taskbar = Point(description="reset position on taskbar", x=1059, y=743)

    #ca2500 GUI, button positions which are likely affected by user dragging windows
    tab_img     = Point(description="ca25, img tab", x=411, y=109)
    tab_data    = Point(description="ca25, data tab", x=548, y=109)
    #ca2500 GUI, data log files, button positions which are NOT affected by user dragging windows
    ca25_data_rows_visible_wo_scrolling = 24 #ca25, data set, each log file per row
    ca25_data_row_step                  = 19 #ca25, data set, row height
    ca25_data_row_first_position        = Point(description="ca25, the first ca2500 data row position", x=18, y=209)
    ca25_data_row_last_position         = Point(description="ca25, the last ca2500 data row position", x=18, y=665)
    #ca2500 GUI, evaluation sub GUI
    evaluation_setting_button      = Point(description="ca25, data evaluation setting button", x=71, y=133)
    open_evaluation_file_button    = Point(description="ca25, open evaluation file button", x=1128, y=79)
    confirm_evaluation_file_button = Point(description="ca25, confirm evaluation file button", x=1008, y=671)
    ca25_data_template_1_18_evl    = Point(description="ca25, date template 1/18.evl", x=559, y=251)
    ca25_data_template_60_40_evl   = Point(description="ca25, date template 60/40.evl", x=559, y=271)
    ca25_data_template_1_6_evl     = Point(description="ca25, date template 1/6.evl", x=559, y=291)

    #excel GUI button positions
    excel_scrollbar_down_button  = Point(description="excel, scrollbar down button", x=1353, y=682)
    excel_scrollbar_down_times   = 17 #excel src worksheet, VBA calculation format
    excel_scrollbar_right_button = Point(description="excel, scrollbar down button", x=1328, y=697)
    #excel worksheets
    excel_tab_src_ws = Point(description="excel, tab 'src' worksheet", x=80, y=698)
    excel_tab_img_ws = Point(description="excel, tab img worksheet", x=117, y=698)
    #excel data/img paste-into-cell positions
    exel_data_paste_into_cell_logname    = Point(description="excel, data paste into cell", x=55, y=121)
    excel_data_paste_into_cell_1_18_evl  = Point(description="excel, date paste into 1/18.evl", x=141, y=139)
    excel_data_paste_into_cell_60_40_evl = Point(description="excel, date paste into 60/40.evl", x=427, y=139)
    excel_data_paste_into_cell_1_6_evl   = Point(description="excel, date paste into 1/6.evl", x=716, y=139)
    excel_img_paste_into_cell_logname    = Point(description="excel, img name paste-into-cell", x=62, y=121)
    excel_img_paste_into_cell_colorful   = Point(description="excel, colorful img paste-into-cell", x=62, y=139)
    excel_img_paste_into_cell_gray       = Point(description="excel, gray img paste-into-cell", x=62, y=355)

    def __init__(self, args={
            'sample_numbers'                  : 200,
            'needs_color_switch'              : True,
            'excel_img_ws_scroll_right_times' : 6,
            'data_wanted'                     : ['100IRE', '50IRE'],
            'data_notwanted'                  : ['ZOOM', 'SET'],
            'img_wanted'                      : ['100IRE', '50IRE','0IRE'],
            'img_notwanted'                   : ['ZOOM', 'SET'],
            'img_shooting_coors'              : (415, 311, 791, 523),
            'copy_data_button'                : Point(description="copy data button", x=688, y=142),
            'copy_data_action'                : Point(description="copy data action", x=688, y=252),
            'color_switch_button'             : Point(description="colorful-gray switch button", x=843, y=142),
            'ca25_data_scrollup_up_button'    : Point(description="ca2500 data scrollup up button", x=340, y=191),
            'ca25_data_scrollbar_checkpoint'  : Point(description="ca2500 data scrollbar checkpoint", x=340, y=204),
            'ca25_data_scrolldown_down_button': Point(description="ca2500 data scrolldown down button", x=340, y=700),
            }):
        """init parameters that are often changed
        """
        self.sample_numbers                  = args['sample_numbers']
        self.needs_color_switch              = args['needs_color_switch']
        self.excel_img_ws_scroll_right_times = args['excel_img_ws_scroll_right_times']
        
        self.data_wanted    = args['data_wanted']
        self.data_notwanted = args['data_notwanted']
        self.img_wanted     = args['img_wanted']
        self.img_notwanted  = args['img_notwanted']

        self.img_shooting_coors               = args['img_shooting_coors']
        self.copy_data_button                 = args['copy_data_button']
        self.copy_data_action                 = args['copy_data_action']
        self.color_switch_button              = args['color_switch_button']
        self.ca25_data_scrollup_up_button     = args['ca25_data_scrollup_up_button']
        self.ca25_data_scrollbar_checkpoint   = args['ca25_data_scrollbar_checkpoint']

        tmp_point = args['ca25_data_scrolldown_down_button']
        rgb_presume_scrolldown_button = self.get_rgb_under_mouse_cursor(tmp_point.x, tmp_point.y)
        if self.is_close(a=rgb_presume_scrolldown_button[0], b=240, c=10):
            #the horizontal scrollbar's height = 17
            self.ca25_data_scrolldown_down_button = Point(description="ca2500 data scrolldown button", x=tmp_point.x, y=tmp_point.y-17)
        else:
            self.ca25_data_scrolldown_down_button = tmp_point

    def clear_clipboard(self):
        """Empty clipboard"""
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        
    def get_clipboard_data(self):
        """get data from clipboard"""
        win32clipboard.OpenClipboard()
        d = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return d

    def choose_data_set_for_tab_swap(self):
        """
        this is a workaround to settle tab swapping between data tab and image tab beforehand
        """
        pyautogui.moveTo(x=self.ca25_data_row_first_position.x, y=self.ca25_data_row_first_position.y)
        pyautogui.click()
        time.sleep(1.5)

    def change_cas25w_layout(self, is_data=True):
        """swatch tabs between data tab and image tab"""
        #back to cas25w main window
        if is_data:      
            #<~ select tab "data"
            pyautogui.moveTo(x=self.tab_data.x, y=self.tab_data.y)
            pyautogui.click()
            time.sleep(1)
        else:
            #<~ select tab "image"
            pyautogui.moveTo(x=self.tab_img.x, y=self.tab_img.y)
            pyautogui.click()
            time.sleep(1)
            
    def get_data_name(self, date_name_coor):
        """get name of data set"""
        pyautogui.moveTo(*date_name_coor)
        pyautogui.click()
        time.sleep(2)
        pyautogui.doubleClick()
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        win32clipboard.OpenClipboard()
        d = win32clipboard.GetClipboardData()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        return d

    def verify_data_name_data(self, dn):
        """verify data name: filter ca2500 log file based on wanted and NotWanted conditions"""
        dn = dn.upper()
        if any(i.upper() in dn for i in self.data_notwanted):
            return False
        else:
            return any(i.upper() in dn for i in self.data_wanted)

    def verify_data_name_img(self, dn):
        """verify data name: filter ca2500 log file based on wanted and NotWanted conditions"""
        dn = dn.upper()
        if any(i.upper() in dn for i in self.img_notwanted):
            return False
        else:
            return any(i.upper() in dn for i in self.img_wanted)

    def transfer_data_from_cas250w_to_xl(self, x, y):
        """transfer data via cas250w -> ctrl + c to copy -> Excel -> ctrl + v to paste"""
        #activate Exel
        pyautogui.moveTo(x=self.excel_position.x, y=self.excel_position.y)
        pyautogui.click()
        #active destine cell
        pyautogui.moveTo(x=x, y=y)
        pyautogui.click()
        #store data
        pyautogui.hotkey('ctrl', 'v')
        #back to cas25w main window
        pyautogui.moveTo(x=self.ca25_position.x, y=self.ca25_position.y)
        pyautogui.click()
        #<~ clear clipboard
        self.clear_clipboard()

    def transfer_data_name(self, from_coor1, to_coor1):
        """
        transfer data name from cas25w to Excel
        """
        pyautogui.moveTo(*from_coor1)
        pyautogui.click()
        time.sleep(1)
        pyautogui.doubleClick()
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        self.transfer_data_from_cas250w_to_xl(*to_coor1)

    def transfer_dot_data(self, from_coor1, to_coor1):
        """
        similar as the function above, but for its dot date
        """
        pyautogui.moveTo(x=self.evaluation_setting_button.x, y=self.evaluation_setting_button.y)
        pyautogui.click()
        time.sleep(1)
        #open position setting
        pyautogui.moveTo(x=self.open_evaluation_file_button.x, y=self.open_evaluation_file_button.y)
        pyautogui.click()
        time.sleep(2) #<~delay 1s to wait for small window popout
        #select position setting1
        pyautogui.moveTo(*from_coor1)
        pyautogui.doubleClick()
        time.sleep(2) #<~delay 1s to wait for cas25w response
        #confirm and back to main window
        pyautogui.moveTo(x=self.confirm_evaluation_file_button.x, y=self.confirm_evaluation_file_button.y)
        pyautogui.click()
        #move to copy data button
        pyautogui.moveTo(x=self.copy_data_button.x, y=self.copy_data_button.y)
        pyautogui.click()
        #copy dot data
        pyautogui.moveTo(x=self.copy_data_action.x, y=self.copy_data_action.y)
        pyautogui.click()
        #paste into Excel
        self.transfer_data_from_cas250w_to_xl(*to_coor1)

    def scroll_xl_window_down(self):
        """
        to suit Excel format
        """
        #activate Exel
        pyautogui.moveTo(x=self.excel_position.x, y=self.excel_position.y)
        pyautogui.click()
        #<~ Position: click Excel scrollbar down button
        pyautogui.moveTo(x=self.excel_scrollbar_down_button.x, y=self.excel_scrollbar_down_button.y)
        for _ in range(self.excel_scrollbar_down_times):
            pyautogui.click()
            time.sleep(0.2)
        # back to cas25w main window
        pyautogui.moveTo(x=self.ca25_position.x, y=self.ca25_position.y)
        pyautogui.click()

    def send_to_clipboard(self, clip_type, data):
        """
        put something on clipboard
        """
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()
        
    def shoot_img(self, x1, y1, x2, y2):
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
        self.send_to_clipboard(win32clipboard.CF_DIB, data)

    def transfer_img(self, to_coor1):
        """
        UF img to Excel
        """
        x1, y1, x2, y2 = self.img_shooting_coors
        self.shoot_img(x1, y1, x2, y2)
        time.sleep(1)
        #paste into Excel
        self.transfer_data_from_cas250w_to_xl(*to_coor1)

    def switch_color(self):
        """
        switch between colorful and gray 
        """
        pyautogui.moveTo(x=self.color_switch_button.x, y=self.color_switch_button.y)
        pyautogui.click()
        time.sleep(0.5)

    def scroll_xl_to_right(self, n=6):
        """
        to suit Excel format
        """
        #<~ Position: click Excel scrollbar down button
        pyautogui.moveTo(x=self.excel_scrollbar_right_button.x, y=self.excel_scrollbar_right_button.y)
        for _ in range(n):
            pyautogui.click()
            time.sleep(0.6)
        # back to cas25w main window
        pyautogui.moveTo(x=self.ca25_position.x, y=self.ca25_position.y)
        pyautogui.click()

    def activate_src_ws(self):
        """activate src worksheet.
        Attention: Excel worksheet names and positions matter!
        """
        #activate Exel
        pyautogui.moveTo(x=self.excel_position.x, y=self.excel_position.y)
        pyautogui.click()
        #select src ws
        pyautogui.moveTo(x=self.excel_tab_src_ws.x, y=self.excel_tab_src_ws.y)
        pyautogui.click()
        
    def activate_img_ws(self, x, y):
        """activate img worksheet.
        Attention: Excel worksheet names and positions matter!
        """
        #activate Exel
        pyautogui.moveTo(x=self.excel_position.x, y=self.excel_position.y)
        pyautogui.click()
        #select img ws cell
        pyautogui.moveTo(x=x, y=y)
        pyautogui.click()
        
    def activate_cas25w(self):
        """activate cas25w.
        """
        # back to cas25w main window
        pyautogui.moveTo(x=self.ca25_position.x, y=self.ca25_position.y)
        pyautogui.click()  

    def reset_window(self):
        """
        click on windows task bar. this action put all open software with UI background
        """
        # pyautogui.moveTo(x=self.reset_position_on_taskbar.x, y=self.reset_position_on_taskbar.y)
        pyautogui.hotkey('win', 'd')
        # pyautogui.click()

    def get_rgb_under_mouse_cursor(self, x, y):
        """
        return r,g,b value of pixel under mouse cursor (x=337, y=191)
        """
        from ctypes import windll
        pyautogui.moveTo(x, y) #color changes when mouse event occurs
        dc = windll.user32.GetDC(0)
        rgb = windll.gdi32.GetPixel(dc, x, y)
        r = rgb & 0xff
        g = (rgb >> 8) & 0xff
        b = (rgb >> 16) & 0xff
        return (r, g, b)

    def move_to_slidebar_downarrow_and_click(self):
        """when N(sample) is greater than 24"""
        pyautogui.moveTo(x=self.ca25_data_scrolldown_down_button.x, y=self.ca25_data_scrolldown_down_button.y)
        pyautogui.click()
        time.sleep(0.5)

    def is_close(self, a, b, c):
        """
        check if a is close to b.
        math: |a - b| <= c
        """
        if abs(a-b) <= c:
            return True
        else:
            return False

    def reset_cas25w_GUI_slidebar(self):
        """reset GUI slidebar
        magic number 46. it's from RBG(46, 151, 207) of scroll_up_arrow_color when mouse is on it
        """
        scrollup_rgb_under_mouse_cursor = self.get_rgb_under_mouse_cursor(self.ca25_data_scrollup_up_button.x, self.ca25_data_scrollup_up_button.y)
        if self.is_close(a=scrollup_rgb_under_mouse_cursor[0], b=46, c=5):
            pyautogui.moveTo(self.ca25_data_scrollup_up_button.x, self.ca25_data_scrollup_up_button.y)
            for i in range(self.sample_numbers):
                if i < 10:
                    pyautogui.click()
                else:
                    pyautogui.moveTo(self.ca25_data_scrollbar_checkpoint.x, self.ca25_data_scrollbar_checkpoint.y)
                    time.sleep(1)
                    #magic number 169 is from RBG(169, 219, 246) when mouse is on scrollbar. otherwise, its gray RBG(240, 240, 240)
                    scrollbar_rgb_under_mouse_cursor = self.get_rgb_under_mouse_cursor(self.ca25_data_scrollbar_checkpoint.x, self.ca25_data_scrollbar_checkpoint.y)
                    if self.is_close(a=scrollbar_rgb_under_mouse_cursor[0], b=169, c=5):
                        #scrollbar is reset
                        break
                    else:
                        pyautogui.moveTo(self.ca25_data_scrollup_up_button.x, self.ca25_data_scrollup_up_button.y)
                        for _ in range(10):
                            pyautogui.click()

    def transfer_ca25_data_to_xl(self, cas25w_data_name_coor):
        """refactor to keep DIY"""
        # data set name
        xl_paste_coor = (self.exel_data_paste_into_cell_logname.x, self.exel_data_paste_into_cell_logname.y) #<~ paste into Excel file
        self.transfer_data_name(cas25w_data_name_coor, xl_paste_coor)
        #<~ get data of 01_UF9 POINT_1_18.evl
        cas25w_setting_coor = (self.ca25_data_template_1_18_evl.x, self.ca25_data_template_1_18_evl.y)
        xl_paste_coor = (self.excel_data_paste_into_cell_1_18_evl.x, self.excel_data_paste_into_cell_1_18_evl.y)
        self.transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
        #<~ get data of 02_UF9 POINT_60_40.evl
        cas25w_setting_coor = (self.ca25_data_template_60_40_evl.x, self.ca25_data_template_60_40_evl.y)
        xl_paste_coor = (self.excel_data_paste_into_cell_60_40_evl.x, self.excel_data_paste_into_cell_60_40_evl.y)
        self.transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
        #<~ get data of 03_UF9 POINT_1_6.evl
        cas25w_setting_coor = (self.ca25_data_template_1_6_evl.x, self.ca25_data_template_1_6_evl.y)
        xl_paste_coor = (self.excel_data_paste_into_cell_1_6_evl.x, self.excel_data_paste_into_cell_1_6_evl.y)
        self.transfer_dot_data(cas25w_setting_coor, xl_paste_coor)
        #<~ setting up Excel before proceeding with next data set
        self.scroll_xl_window_down()
        return

    def transfer_ca25_img_to_xl(self, cas25w_data_name_coor):
        """refactor to keep DIY"""
        # data set name
        xl_paste_coor = (self.excel_img_paste_into_cell_logname.x, self.excel_img_paste_into_cell_logname.y)
        self.transfer_data_name(cas25w_data_name_coor, xl_paste_coor)
        #<~ get img: colorful        
        xl_paste_coor = (self.excel_img_paste_into_cell_colorful.x, self.excel_img_paste_into_cell_colorful.y)
        self.transfer_img(xl_paste_coor)
        #<~ get img: gray
        if self.needs_color_switch:
            self.switch_color()
            xl_paste_coor = (self.excel_img_paste_into_cell_gray.x, self.excel_img_paste_into_cell_gray.y)
            self.transfer_img(xl_paste_coor)
        self.activate_img_ws(x=self.excel_tab_img_ws.x, y=self.excel_tab_img_ws.y)
        self.scroll_xl_to_right(n=self.excel_img_ws_scroll_right_times)
        if self.needs_color_switch:
            self.switch_color()
        return

    def main_get_data(self):
        self.reset_window()     
        ###<~ transfer dataname from cas25w to Excel: #<~well, push to limit: max 25 sets of data
        self.activate_src_ws()
        self.activate_cas25w()
        self.reset_cas25w_GUI_slidebar()
        self.choose_data_set_for_tab_swap()
        self.change_cas25w_layout(is_data=True) #<~select tab "data"
        
        tmp = ""
        cnt = 0

        for i in range(self.sample_numbers):       
            if i <= self.ca25_data_rows_visible_wo_scrolling:
                cas25w_data_name_coor = (self.ca25_data_row_first_position.x, self.ca25_data_row_first_position.y + i * self.ca25_data_row_step)
                try:
                    dn = self.get_data_name(cas25w_data_name_coor)
                    if tmp == dn:
                        cnt += 1
                    if cnt >= 1: #<~~ break if data name duplicates about twice
                        break
                    if self.verify_data_name_data(dn):
                        self.transfer_ca25_data_to_xl(cas25w_data_name_coor)  
                except TypeError: #<~ stops when hitting blank data row
                    break
                tmp = dn

            if i > self.ca25_data_rows_visible_wo_scrolling:
                self.move_to_slidebar_downarrow_and_click()
                cas25w_data_name_coor = (self.ca25_data_row_last_position.x, self.ca25_data_row_last_position.y)
                try:
                    dn = self.get_data_name(cas25w_data_name_coor)
                    if tmp == dn:
                        cnt += 1
                    if cnt >= 1: #<~~ break if data name duplicates about twice
                        #get data from real final row
                        #magic number 51 is from RBG(51, 153, 255) when data row is seletec
                        datarow_rgb_under_mouse_cursor = self.get_rgb_under_mouse_cursor(self.ca25_data_row_last_position.x, self.ca25_data_row_last_position.y)
                        if self.is_close(a=datarow_rgb_under_mouse_cursor[0], b=51, c=5):
                            cas25w_data_name_coor = (self.ca25_data_row_last_position.x, self.ca25_data_row_last_position.y + self.ca25_data_row_step)
                            dn = self.get_data_name(cas25w_data_name_coor)
                            if self.verify_data_name_data(dn):
                                self.transfer_ca25_data_to_xl(cas25w_data_name_coor)
                                break
                            else:
                                break                             
                        else:
                            break
                    if self.verify_data_name_data(dn):
                        self.transfer_ca25_data_to_xl(cas25w_data_name_coor)
                except TypeError: #<~ stops when hitting blank data row
                    break
                tmp = dn

    def main_get_img(self):

        self.reset_window()
        ###<~ transfer dataname from cas25w to Excel: #<~well, push to limit: max 25 sets of data
        self.activate_cas25w()
        self.reset_cas25w_GUI_slidebar()
        self.choose_data_set_for_tab_swap()
        self.change_cas25w_layout(is_data=False) #<~select tab "image"
        #<~ active Img worksheet in Excel
        self.activate_img_ws(x=self.excel_tab_img_ws.x, y=self.excel_tab_img_ws.y)
        self.scroll_xl_to_right(n=1) #<~ match Img format in VBA

        tmp = ""
        cnt = 0
        for i in range(self.sample_numbers):
            if i <= self.ca25_data_rows_visible_wo_scrolling:
                cas25w_data_name_coor = (self.ca25_data_row_first_position.x, self.ca25_data_row_first_position.y + self.ca25_data_row_step * i)
                try:
                    dn = self.get_data_name(cas25w_data_name_coor)
                    if tmp == dn:
                        cnt += 1
                    if cnt >= 1: #<~~ break if data name duplicates about twice
                        break
                    if self.verify_data_name_img(dn):
                        self.transfer_ca25_img_to_xl(cas25w_data_name_coor)
                except TypeError:
                    break
                tmp = dn

            if i > self.ca25_data_rows_visible_wo_scrolling:
                self.move_to_slidebar_downarrow_and_click()
                cas25w_data_name_coor = (self.ca25_data_row_last_position.x, self.ca25_data_row_last_position.y)           
                try:
                    dn = self.get_data_name(cas25w_data_name_coor)
                    if tmp == dn:
                        cnt += 1
                    if cnt >= 1: #<~~ break if data name duplicates about twice
                        #get data from real final row
                        #magic number 51 is from RBG(51, 153, 255) when data row is seletec
                        datarow_rgb_under_mouse_cursor = self.get_rgb_under_mouse_cursor(self.ca25_data_row_last_position.x, self.ca25_data_row_last_position.y)
                        if self.is_close(a=datarow_rgb_under_mouse_cursor[0], b=51, c=5):
                            cas25w_data_name_coor = (self.ca25_data_row_last_position.x, self.ca25_data_row_last_position.y + self.ca25_data_row_step)
                            dn = self.get_data_name(cas25w_data_name_coor)
                            if self.verify_data_name_img(dn):
                                self.transfer_ca25_img_to_xl(cas25w_data_name_coor)
                                break
                            else:
                                break                            
                        else:
                            break
                    if self.verify_data_name_img(dn):
                        self.transfer_ca25_img_to_xl(cas25w_data_name_coor)
                except TypeError:
                    break
                tmp = dn

def main():
    extractor1 = CA2500_Extractor()
    extractor1.clear_clipboard()
    extractor1.main_get_data()
    time.sleep(1)
    extractor1.clear_clipboard()
    extractor1.main_get_img()
    time.sleep(1)

if __name__ == "__main__":
    main()