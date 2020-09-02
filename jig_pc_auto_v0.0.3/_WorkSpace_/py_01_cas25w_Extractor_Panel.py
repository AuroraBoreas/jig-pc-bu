import tkinter
from tkinter import messagebox
import sys, os
sys.path.append(os.path.dirname(__file__))
from lib.pkg_ca2500_manager.core import CA2500_Extractor, Point

class CA2500_Extractor_Wanted(CA2500_Extractor):
    def verify_data_name_img(self, dn):
        """verify data name: filter ca2500 log file based on wanted and NotWanted conditions"""
        dn = dn.upper()
        if any(i in dn for i in self.img_notwanted):
            return False
        else:
            return all(i in dn for i in self.img_wanted)

if __name__=='__main__':
    root = tkinter.Tk()
    root.withdraw()

    user_answer = messagebox.askquestion(title="CA2500 Panel Data Extracter@ZL", message="Switch to CA2500 and Adjust for 100IRE50IRE pls", icon="warning", type="yesno")
    if user_answer == 'yes':
        ## default setting
        e1 = CA2500_Extractor()
        e1.clear_clipboard()
        e1.main_get_data()
        e1.clear_clipboard()
        e1.main_get_img()

        ### get a specific panel
        # e0 = CA2500_Extractor()
        # e0.data_wanted = ['IRE']
        # e0.data_notwanted = ['SET', 'ZOOMIN', 'SCC0000011', 'SCC0000007']
        # e0.img_wanted = ['IRE']
        # e0.img_notwanted = ['SET', 'ZOOMIN', 'SCC0000011', 'SCC0000007']
        # e0.clear_clipboard()
        # e0.main_get_data()
        # e0.clear_clipboard()
        # e0.main_get_img()
    messagebox.showinfo(title="CA2500 Panel Data Extracter@ZL", message="All Done")
    root.destroy()