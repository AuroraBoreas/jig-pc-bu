""" A modue manages measurement of cs2000 for an easy operation by one man, or serial samples testing

It has basic functionalities as follows.
- auto open cs2000 software
- auto setup cs2000 per instance
- auto save log with precise measurement datetime and name

Changelog
- v0.0.1, inital release
- v0.0.2, refactor as package

About
- author: @ZL, 202007
- LICENSE: MIT

"""
import pyautogui, time, win32gui, win32con, pyttsx3, win32com.client, subprocess
from datetime import datetime
from pywinauto.findwindows import find_window
from win32con import (SW_SHOW, SW_RESTORE)
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from lib.pkg_cs2000_manager.point import Point
from lib.pkg_Windows_Sound_Manager import sound

pyautogui.PAUSE = 0.25

class CS2kMeasurer:
    """CS2000 measurement helper

    Returns:
        class: general purpose control of cs2000 measurement
    """
    wait_time_white = 12
    wait_time_red   = 5
    wait_time_green = 5
    wait_time_blue  = 5
    wait_time_black = 25

    wait_time_xl_react  = 1.5
    wait_time_xl_rename = 2.5

    cs2000_close_button = Point(description="cs2000 close button position", x=951, y=116)
    cs2000_checkall_box = Point(description="cs2000 checkall box", x=427, y=171)
    cs2000_white_box    = Point(description="cs2000 White box", x=427, y=201)
    cs2000_red_box      = Point(description="cs2000 RED box", x=427, y=231)
    cs2000_green_box    = Point(description="cs2000 GREEN box", x=427, y=261)
    cs2000_blue_box     = Point(description="cs2000 BLUE box", x=427, y=291)
    cs2000_black_box    = Point(description="cs2000 BLACK box", x=427, y=321)

    cs2000_remote_button     = Point(description="cs2000 remote button", x=927, y=371)
    cs2000_meas_button       = Point(description="cs2000 meas button", x=933, y=422)
    cs2000_save2excel_button = Point(description="cs2000 save to excel button position", x=928, y=530)

    cs2000_warn_tag_connection  = Point(description="cs2000 warning window connection", x=575, y=361)
    cs2000_warn_tag_before_meas = Point(description="cs2000 warning window before meas", x=620, y=381)
    cs2000_warn_tag_after_meas  = Point(description="cs2000 warning window after meas", x=602, y=381)
    cs2000_sw_path = r"C:\Users\admin panel-Design\Documents\Desktop\MeSpectra for CS2000 β02 - Copy.exe"

    def __init__(self, **kwargs):
        """initialize instance with **kwargs
        """
        self.meas_time           = kwargs['meas_time']
        self.pmod_scc_no         = kwargs['pmod_scc_no']
        self.color               = kwargs['color']
        self.ca_mode             = kwargs['ca_mode']
        self.user_psg500_op_time = kwargs['user_psg500_op_time']
        self.ageing_hour         = kwargs['ageing_hour']

    def __get_rgb_under_mouse_cursor(self, x, y):
        """
        return r,g,b value of pixel under mouse cursor, ie. (x=337, y=191)
        """
        from ctypes import windll
        pyautogui.moveTo(x, y) #color changes when mouse event occurs
        dc = windll.user32.GetDC(0)
        rgb = windll.gdi32.GetPixel(dc, x, y)
        r = rgb & 0xff
        g = (rgb >> 8) & 0xff
        b = (rgb >> 16) & 0xff
        return (r, g, b)

    def __IsDone(self, point_warn_before, point_warn_after):
        """
        return true if the current measurement is done or false. based
        """
        color_flag = 230 # background color of prompt window
        bound = 10 # bound of color difference
        isDone1 = abs(self.__get_rgb_under_mouse_cursor(point_warn_before.x, point_warn_before.y)[0] - color_flag) <= bound
        isDone2 = abs(self.__get_rgb_under_mouse_cursor(point_warn_after.x, point_warn_after.y)[0] - color_flag) <= bound
        if isDone1 or isDone2:
            return True
        return False
    
        
    def __IsConnect(self, point_warn_connect):
        """return true if cs2000 connects with jig pc or false

        Args:
            point_warn_connect (Point object): Point object that contains a piece of information of a screen pixel

        Returns:
            Boolean: true or false
        """
        time.sleep(0.5)
        color_flag = 176 # background color of prompt window
        bound = 10 # bound of color difference
        isConnect = abs(self.__get_rgb_under_mouse_cursor(point_warn_connect.x, point_warn_connect.y)[0] - color_flag) <= bound
        if isConnect:
            return False
        return True

    def __exit(self, point_warn_connect, engine):
        """close cs2000 software and mute pc speaker volume when connection faied

        Args:
            point_warn_connect (Point object): Point of cs2000_warning_connection window
        """
        if not self.__IsConnect(point_warn_connect):
            engine.say("Connection Error")
            engine.runAndWait()
            sound.Sound.mute()
            pyautogui.hotkey('Enter')
            pyautogui.hotkey('Enter')
            sys.exit("Error: Connecting to CS2000 failed")
        return

    def __get_windows_placement(self, window_id):
        """find window placement based on window name"""
        return win32gui.GetWindowPlacement(window_id)[1]

    def __set_active_window(self, window_id):
        """make window active"""
        if self.__get_windows_placement(window_id) == 2:
            win32gui.ShowWindow(window_id, SW_RESTORE)
        else:
            win32gui.ShowWindow(window_id, SW_SHOW)
        win32gui.SetForegroundWindow(window_id)
        win32gui.SetActiveWindow(window_id)
        return 

    def __close_excel_window(self, file_name):
        """terminate excel application"""
        xl = win32com.client.Dispatch("Excel.Application")
        xl.DisplayAlerts = False
        for wb in xl.workbooks:
            if file_name.lower() in wb.name.lower():
                wb.Close(True)
                break
        return

    def __reset_ui_windows(self):
        """windows hotkey to dispay desktop. reset focus
        """
        pyautogui.hotkey('win', 'd')
        return

    def __activate_cs2k(self):
        """activate or start or bring afront cs2000 software
        """
        subprocess.Popen([self.cs2000_sw_path])
        return

    def __find_excel_on_taskbar(self):
        """find excel and bring to top"""
        window_id = find_window(title='Microsoft Excel - Sheet1')
        self.__set_active_window(window_id)
        return

    def __click_cs2000sw_saveToExcel(self):
        """save to excel"""
        pyautogui.moveTo(x=self.cs2000_save2excel_button.x, y=self.cs2000_save2excel_button.y)
        pyautogui.click()
        return

    def __close_cs2000sw(self):
        """close software"""
        self.__reset_ui_windows()
        self.__activate_cs2k()
        pyautogui.moveTo(x=self.cs2000_close_button.x, y=self.cs2000_close_button.y)
        pyautogui.click()
        return

    def __cs2000_sw_meas(self):
        """cs2000 software sets up, measures and logs"""
        engine = pyttsx3.init()
        if not self.ca_mode: # meas specific colore only
            engine.say("Start to measure momentarily. {} pattern".format(self.color))
            engine.runAndWait()

            # make cs2000sw as top window
            self.__reset_ui_windows()
            self.__activate_cs2k()
            # meas
            if self.color.lower() == "white":
                pyautogui.moveTo(x=self.cs2000_white_box.x, y=self.cs2000_white_box.y)  # checkbox white
                pyautogui.click()
            if self.color.lower() == "red":
                pyautogui.moveTo(x=self.cs2000_red_box.x, y=self.cs2000_red_box.y)  # checkbox red
                pyautogui.click()
            if self.color.lower() == "green":
                pyautogui.moveTo(x=self.cs2000_green_box.x, y=self.cs2000_green_box.y)  # checkbox green
                pyautogui.click()
            if self.color.lower() == "blue":
                pyautogui.moveTo(x=self.cs2000_blue_box.x, y=self.cs2000_blue_box.y)  # checkbox blue
                pyautogui.click()
            if self.color.lower() == "black":
                pyautogui.moveTo(x=self.cs2000_black_box.x, y=self.cs2000_black_box.y)  # checkbox black
                pyautogui.click()
            pyautogui.moveTo(x=self.cs2000_remote_button.x, y=self.cs2000_remote_button.y)  # click on button of REMOTE
            pyautogui.click()
            # exits if connection does not establish between cs2000 and jig pc
            self.__exit(self.cs2000_warn_tag_connection, engine)
            time.sleep(self.user_psg500_op_time) #  wait for user to change PSG500
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            pyautogui.moveTo(x=self.cs2000_meas_button.x, y=self.cs2000_meas_button.y)
            pyautogui.click()
            pyautogui.hotkey('Enter')
            # waiting for cs2000 meas, N seconds
            if self.color.lower() == "white":
                time.sleep(self.wait_time_white)
            if self.color.lower() == "red":
                time.sleep(self.wait_time_red)
            if self.color.lower() == "green":
                time.sleep(self.wait_time_green)
            if self.color.lower() == "blue":
                time.sleep(self.wait_time_blue)
            if self.color.lower() == "black":
                time.sleep(self.wait_time_black)
            # double check if measurement finished or wait till finish
            while not self.__IsDone(self.cs2000_warn_tag_before_meas, self.cs2000_warn_tag_after_meas):
                time.sleep(self.wait_time_red)
            pyautogui.hotkey('Enter')
            # save to excel
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            self.__click_cs2000sw_saveToExcel()
            time.sleep(self.wait_time_xl_react) # waiting for excel operation
            pyautogui.hotkey('Enter')
            time.sleep(self.wait_time_xl_react)
            engine.say("Record successfully")
            engine.runAndWait()
        
        if self.ca_mode: # WRGBB
            engine.say("Start to measure momentarily, WRGBB mode")
            engine.runAndWait()
            # make cs2000sw as top window
            self.__reset_ui_windows()
            self.__activate_cs2k()
            # set up
            pyautogui.moveTo(x=self.cs2000_checkall_box.x, y=self.cs2000_checkall_box.y)#click on box of CheckAll
            pyautogui.click()
            pyautogui.moveTo(x=self.cs2000_remote_button.x, y=self.cs2000_remote_button.y)#click on button of REMOTE
            pyautogui.click()
            # exits if connection does not establish between cs2000 and jig pc
            self.__exit(self.cs2000_warn_tag_connection, engine)
            # meas White
            engine.say("Start to measure White, {} seconds later".format(self.user_psg500_op_time))
            engine.runAndWait()
            time.sleep(self.user_psg500_op_time) #  wait for user to change PSG500
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            pyautogui.moveTo(x=self.cs2000_meas_button.x, y=self.cs2000_meas_button.y)
            pyautogui.click() 
            pyautogui.hotkey('Enter')
            time.sleep(self.wait_time_white) # waiting for cs2000 meas, N seconds
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            # double check if measurement finished or wait till finish
            while not self.__IsDone(self.cs2000_warn_tag_before_meas, self.cs2000_warn_tag_after_meas):
                time.sleep(self.wait_time_red)
            # Red
            engine.say("Start to measure Red, {} seconds later".format(self.user_psg500_op_time))
            engine.runAndWait()
            time.sleep(self.user_psg500_op_time + 3) #  wait for user to change PSG500
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            pyautogui.hotkey('Enter')
            time.sleep(self.wait_time_red) # waiting for cs2000 meas, N seconds
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            # double check if measurement finished or wait till finish
            while not self.__IsDone(self.cs2000_warn_tag_before_meas, self.cs2000_warn_tag_after_meas):
                time.sleep(self.wait_time_red)
            # Green
            engine.say("Start to measure Green, {} seconds later".format(self.user_psg500_op_time))
            engine.runAndWait()
            time.sleep(self.user_psg500_op_time + 3) #  wait for user to change PSG500
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            pyautogui.hotkey('Enter') 
            time.sleep(self.wait_time_green) # waiting for cs2000 meas, N seconds
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            # double check if measurement finished or wait till finish
            while not self.__IsDone(self.cs2000_warn_tag_before_meas, self.cs2000_warn_tag_after_meas):
                time.sleep(self.wait_time_red)
            # Blue
            engine.say("Start to measure Blue, {} seconds later".format(self.user_psg500_op_time))
            engine.runAndWait()
            time.sleep(self.user_psg500_op_time + 3) #  wait for user to change PSG500
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            pyautogui.hotkey('Enter')
            time.sleep(self.wait_time_blue) # waiting for cs2000 meas, N seconds
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            # double check if measurement finished or wait till finish
            while not self.__IsDone(self.cs2000_warn_tag_before_meas, self.cs2000_warn_tag_after_meas):
                time.sleep(self.wait_time_red)
            # Black
            engine.say("Start to measure Black, {} seconds later".format(self.user_psg500_op_time))
            engine.runAndWait()
            time.sleep(self.user_psg500_op_time) #  wait for user to change PSG500
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            pyautogui.hotkey('Enter')
            time.sleep(self.wait_time_black) # waiting for cs2000 meas, N seconds
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            # double check if measurement finished or wait till finish
            while not self.__IsDone(self.cs2000_warn_tag_before_meas, self.cs2000_warn_tag_after_meas):
                time.sleep(self.wait_time_red)
            pyautogui.hotkey('Enter')
            
            # save to excel
            self.__activate_cs2k() # make sure to bring cs2000 sw front afer the long waiting time
            self.__click_cs2000sw_saveToExcel()
            time.sleep(self.wait_time_xl_react) # waiting for excel operation
            pyautogui.hotkey('Enter')
            time.sleep(self.wait_time_xl_react)
            engine.say("Record successfully")
            engine.runAndWait()
        engine.stop() 
        return 

    def __save_excel_data_to_local_documents_folder(self, file_name):
        """save cs2000 log data"""
        # make excel as top window
        self.__find_excel_on_taskbar()
        pyautogui.hotkey('ALT','F','S')   
        time.sleep(self.wait_time_xl_rename)
        pyautogui.typewrite(file_name)
        time.sleep(self.wait_time_xl_rename)
        pyautogui.hotkey('Enter')
        time.sleep(self.wait_time_xl_rename)
        self.__close_excel_window(file_name)
        self.__close_cs2000sw()
        return

    def start(self):
        self.__cs2000_sw_meas()
        if self.ca_mode:
            meas_color = "WRGBB"
        else:
            meas_color = self.color
        file_name = datetime.now().strftime('%Y-%m-%d %H%M%S') + " {0}_{1}_{2}H".format(self.pmod_scc_no, meas_color, self.ageing_hour)
        self.__save_excel_data_to_local_documents_folder(file_name.upper())
        return