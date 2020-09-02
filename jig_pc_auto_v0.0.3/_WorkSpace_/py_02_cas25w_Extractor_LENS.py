import tkinter
from tkinter import (
    Checkbutton,
    Frame,
    Label,
    Button,
    Text,
    messagebox,
)
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

class app_wm():
    #unify label widths to make GUI pretty
    label_width = 15
    txtbox_height = 2
    txtbox_width = 40

    #tags: Panel 100IRE, 50IRE
    wanted_10050IRE = "100IRE, 50IRE"
    unwanted_10050IRE = "SET, ZOOM"
    #tags: Panel 100IRE Gray
    wanted_100IREGray = "100IRE"
    unwanted_100IREGray = "SET, ZOOM, 50IRE"
    #tags: Panel 100IRE ZoomIn Gray
    wanted_100ZoomInGray = "100IRE, ZOOM"
    unwanted_100ZoomInGray = ""    

    def __init__(self):
        self.root = tkinter.Tk()
        self.root.title("CA25 Image Extractor v0, @ZL, 2020")

        #get screen width and height
        self.w, self.h = 420, 420
        self.ws = self.root.winfo_screenwidth()
        self.hs = self.root.winfo_screenheight()
        x = int((self.ws - self.w) / 2)
        y = int((self.hs - self.h) / 2)
        self.root.geometry("{}x{}+{}+{}".format(self.w, self.h, x, y))
        
        #check button
        self.var_10050IRE = tkinter.IntVar()
        self.chb_wanted_10050IRE = Checkbutton(self.root, text='Need 100 and 50IRE?', variable=self.var_10050IRE)
        self.chb_wanted_10050IRE.grid(row=0, column=0, sticky="nw")
        ##<~frame1
        self.fm1 = Frame(self.root)
        self.fm1.config(highlightbackground='gray', highlightcolor='gray', highlightthickness=1, bd=0)
        self.fm1.grid(row=1, column=0, sticky='w', padx=10, pady=5)
        #data wanted
        self.lbl_wanted_10050IRE = Label(self.fm1, text="Data Wanted:", width=self.label_width)
        self.lbl_wanted_10050IRE.grid(row=0, column=1, sticky='nw')
        self.txt_wanted_10050IRE = Text(self.fm1, width=self.txtbox_width, height=self.txtbox_height)
        self.txt_wanted_10050IRE.config(font=('Arial', 10))
        self.txt_wanted_10050IRE.grid(row=0, column=2, sticky='w')
        self.txt_wanted_10050IRE.insert(tkinter.INSERT, self.wanted_10050IRE)
        #data unwanted
        self.lbl_unwanted_10050IRE = Label(self.fm1, text="Data unwanted:", width=self.label_width)
        self.lbl_unwanted_10050IRE.grid(row=1, column=1, sticky='nw')
        self.txt_unwanted_10050IRE = Text(self.fm1, width=self.txtbox_width, height=self.txtbox_height)
        self.txt_unwanted_10050IRE.config(font=('Arial', 10))
        self.txt_unwanted_10050IRE.grid(row=1, column=2, sticky='w')
        self.txt_unwanted_10050IRE.insert(tkinter.INSERT, self.unwanted_10050IRE)        

        #check button
        self.var_100IREGray = tkinter.IntVar()
        self.chb_wanted_100IREGray = Checkbutton(self.root, text='Need 100IRE Gray?', variable=self.var_100IREGray)
        self.chb_wanted_100IREGray.grid(row=2, column=0, sticky="nw")
        ##<~frame2
        self.fm2 = Frame(self.root)
        self.fm2.config(highlightbackground='gray', highlightcolor='gray', highlightthickness=1, bd=0)
        self.fm2.grid(row=3, column=0, sticky='w', padx=10, pady=5)
        #data wanted
        self.lbl_wanted_100IREGray = Label(self.fm2, text="Data Wanted:", width=self.label_width)
        self.lbl_wanted_100IREGray.grid(row=0, column=0, sticky='nw')
        self.txt_wanted_100IREGray = Text(self.fm2, width=self.txtbox_width, height=self.txtbox_height)
        self.txt_wanted_100IREGray.config(font=('Arial', 10))
        self.txt_wanted_100IREGray.grid(row=0, column=1, sticky='w')
        self.txt_wanted_100IREGray.insert(tkinter.INSERT, self.wanted_100IREGray)
        #data unwanted
        self.lbl_unwanted_100IREGray = Label(self.fm2, text="Data unwanted:", width=self.label_width)
        self.lbl_unwanted_100IREGray.grid(row=1, column=0, sticky='nw')
        self.txt_unwanted_100IREGray = Text(self.fm2, width=self.txtbox_width, height=self.txtbox_height)
        self.txt_unwanted_100IREGray.config(font=('Arial', 10))
        self.txt_unwanted_100IREGray.grid(row=1, column=1, sticky='w')
        self.txt_unwanted_100IREGray.insert(tkinter.INSERT, self.unwanted_100IREGray)    

        #check button
        self.var_100ZoomInGray = tkinter.IntVar()
        self.chb_wanted_100ZoomInGray = Checkbutton(self.root, text='Need 100IRE Mura Enhancement?(AKA ZoomIn)', variable=self.var_100ZoomInGray)
        self.chb_wanted_100ZoomInGray.grid(row=4, column=0, sticky="nw")
        ##<~frame2
        self.fm3 = Frame(self.root)
        self.fm3.config(highlightbackground='gray', highlightcolor='gray', highlightthickness=1, bd=0)
        self.fm3.grid(row=5, column=0, sticky='w', padx=10, pady=5)
        #data wanted
        self.lbl_wanted_100ZoomInGray = Label(self.fm3, text="Data Wanted:", width=self.label_width)
        self.lbl_wanted_100ZoomInGray.grid(row=0, column=0, sticky='nw')
        self.txt_wanted_100ZoomInGray = Text(self.fm3, width=self.txtbox_width, height=self.txtbox_height)
        self.txt_wanted_100ZoomInGray.config(font=('Arial', 10))
        self.txt_wanted_100ZoomInGray.grid(row=0, column=1, sticky='w')
        self.txt_wanted_100ZoomInGray.insert(tkinter.INSERT, self.wanted_100ZoomInGray)
        #data unwanted
        self.lbl_unwanted_100ZoomInGray = Label(self.fm3, text="Data unwanted:", width=self.label_width)
        self.lbl_unwanted_100ZoomInGray.grid(row=1, column=0, sticky='nw')
        self.txt_unwanted_100ZoomInGray = Text(self.fm3, width=self.txtbox_width, height=self.txtbox_height)
        self.txt_unwanted_100ZoomInGray.config(font=('Arial', 10))
        self.txt_unwanted_100ZoomInGray.grid(row=1, column=1, sticky='w')
        self.txt_unwanted_100ZoomInGray.insert(tkinter.INSERT, self.unwanted_100ZoomInGray)

        #<~ button, click to start ocr
        self.btn_start = Button(self.root, text='START', width=20, height=self.txtbox_height, command=self.start_extractor)
        self.btn_start.grid(row=6, column=0, sticky='nw', padx=120, pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def start_extractor(self):
        self.root.withdraw()
        if self.var_10050IRE.get():
            self.extract_ca25_10050IRE()

        if self.var_100IREGray.get():
            self.extract_ca25_100IREGray()

        if self.var_100ZoomInGray.get():
            self.extract_ca25_100ZoomInGray()
        messagebox.showinfo(title="Grab Img@ZL", message="All Done")
        return 

    def get_wanted10050IRE(self):
        txt_box_content = self.txt_wanted_10050IRE.get("1.0", 'end-1c')
        return [item.strip().upper() for item in txt_box_content.split(',') if item] if txt_box_content else []

    def get_unwanted10050IRE(self):
        txt_box_content = self.txt_unwanted_10050IRE.get("1.0", 'end-1c')
        return [item.strip().upper() for item in txt_box_content.split(',') if item] if txt_box_content else []

    def get_wanted100IREGray(self):
        txt_box_content = self.txt_wanted_100IREGray.get("1.0", 'end-1c')
        return [item.strip().upper() for item in txt_box_content.split(',') if item] if txt_box_content else []

    def get_unwanted100IREGray(self):
        txt_box_content = self.txt_unwanted_100IREGray.get("1.0", 'end-1c')
        return [item.strip().upper() for item in txt_box_content.split(',') if item] if txt_box_content else []

    def get_wanted100ZoomIngGray(self):
        txt_box_content = self.txt_wanted_100ZoomInGray.get("1.0", 'end-1c')
        return [item.strip().upper() for item in txt_box_content.split(',') if item] if txt_box_content else [] 

    def get_unwanted100ZoomIngGray(self):
        txt_box_content = self.txt_unwanted_100ZoomInGray.get("1.0", 'end-1c')
        return [item.strip().upper() for item in txt_box_content.split(',') if item] if txt_box_content else []

    def extract_ca25_10050IRE(self):
        user_answer = messagebox.askquestion(title="Grab Img@ZL", message="Switch to CA2500 and Adjust for 100IRE50IRE Pls", icon="warning", type="yesno")
        if user_answer == 'yes':
            extractor = CA2500_Extractor()
            extractor.img_wanted = self.get_wanted10050IRE()
            extractor.img_notwanted = self.get_unwanted10050IRE()
            extractor.clear_clipboard()
            extractor.main_get_img()
        return

    def extract_ca25_100IREGray(self):
        user_answer = messagebox.askquestion(title="Grab Img@ZL", message="Switch to CA2500 and Adjust for 100IRE Gray Pls", icon="warning", type="yesno")
        if user_answer == 'yes':
            extractor = CA2500_Extractor_Wanted()
            extractor.needs_color_switch = False
            extractor.excel_tab_img_ws = Point(description="excel worksheet Img_100IRE_Gray", x=185, y=699)
            extractor.img_wanted = self.get_wanted100IREGray()
            extractor.img_notwanted = self.get_unwanted100IREGray()
            extractor.clear_clipboard()
            extractor.main_get_img()
        return 

    def extract_ca25_100ZoomInGray(self):
        user_answer = messagebox.askquestion("Grab Img", message="Switch to CA2500 and Adjust for 100IRE ZoomIn Gray Pls", icon="warning", type="yesno")
        if user_answer == 'yes':
            extractor = CA2500_Extractor_Wanted()
            extractor.needs_color_switch = False
            extractor.excel_img_ws_scroll_right_times = 7 #excel, Img_ZoomIn100IRE_Gray, image occupies 7 cells
            extractor.excel_tab_img_ws = Point(description="excel worksheet Img_ZoomIn100IRE_Gray", x=328, y=699)
            extractor.img_shooting_coors = (369, 181, 836, 648)#ca25, ZoomIn, Point(x=369, y=181), Point(x=836, y=648)
            extractor.img_wanted = self.get_wanted100ZoomIngGray()
            extractor.img_notwanted = self.get_unwanted100ZoomIngGray()
            extractor.clear_clipboard()
            extractor.main_get_img()
        return

    def close_app(self):
        self.root.destroy()

    def mainloop(self):
        self.root.mainloop()

def main():
    app = app_wm()
    app.mainloop()


if __name__=='__main__':
    main()


