import tkinter as tk
from tkinter import messagebox
import sys, os
sys.path.append(os.path.dirname(__file__))
from lib.pkg_ca2500_manager.core import CA2500_Extractor, Point

class CA2500_Extractor_Wanted(CA2500_Extractor):
    def verify_data_name_img(self, dn):
        """verify data name: filter ca2500 log file based on wanted and NotWanted conditions"""
        dn = dn.upper()
        if any(i.upper() in dn for i in self.img_notwanted):
            return False
        else:
            return all(i.upper() in dn for i in self.img_wanted)

if __name__=='__main__':
    root = tk.Tk()
    root.withdraw()
    
    # ##CAUTION: user must set up ca2500 first, including reference tab, colorful, range
    # messagebox.showinfo("Grab Img", message="Adjust for 100IRE50IRE Pls") #wait till user adjusts CA2500
    # extractor1 = CA2500_Extractor()
    # extractor1.clear_clipboard()
    # extractor1.main_get_img()

    ## all images
    messagebox.showinfo("Grab User-defined Images", message="Please set up CA2500")
    e1 = CA2500_Extractor_Wanted()
    e1.data_notwanted= []
    e1.data_wanted = ['IRE']
    e1.img_notwanted = []
    e1.img_wanted = ['IRE']
    e1.clear_clipboard()
    e1.main_get_data()
    e1.clear_clipboard()
    e1.main_get_img()

    # messagebox.showinfo(title="Grab Img@ZL", message="Adjust for 100IRE Gray Pls") #wait till user adjusts CA2500
    # extractor2 = CA2500_Extractor_Wanted()
    # extractor2.needs_color_switch = False
    # extractor2.excel_tab_img_ws = Point(description="excel worksheet Img_100IRE_Gray", x=185, y=699)
    # extractor2.img_wanted = ['100IRE']
    # extractor2.img_notwanted = ['SET', 'ZOOM']
    # extractor2.clear_clipboard()
    # extractor2.main_get_img()

    # messagebox.showinfo("Grab Img", message="Adjust for 100IRE ZoomIn Pls") #wait till user adjusts CA2500
    # extractor3 = CA2500_Extractor_Wanted()
    # extractor3.needs_color_switch = False
    # extractor3.excel_img_ws_scroll_right_times = 7 #excel, Img_ZoomIn100IRE_Gray, image occupies 7 cells
    # extractor3.excel_tab_img_ws = Point(description="excel worksheet Img_ZoomIn100IRE_Gray", x=328, y=699)
    # extractor3.img_shooting_coors = (369, 181, 836, 648)#ca25, ZoomIn, Point(x=369, y=181), Point(x=836, y=648)
    # extractor3.img_wanted = ['100IRE', 'ZOOMIN']
    # extractor3.img_notwanted = []
    # extractor3.clear_clipboard()
    # extractor3.main_get_img()

    messagebox.showinfo(title="Grab Img@ZL", message="All Done") #wait till user adjusts CA2500
    root.destroy()


