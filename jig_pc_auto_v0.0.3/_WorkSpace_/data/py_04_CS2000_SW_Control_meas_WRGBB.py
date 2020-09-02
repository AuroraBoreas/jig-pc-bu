"""
Automate CS2000 measurement @ZL, 2019
"""
import time, pyttsx3
import tkinter as tk
from tkinter import messagebox
from lib.pkg_cs2000_manager.core import CS2kMeasurer
from lib.pkg_Windows_Sound_Manager import sound

def cs2k_mease_WRGB(sample_name):
    #mute then toggle speaker on
    sound.Sound.mute()
    sound.Sound.volume_min()
    sound.Sound.volume_max()

    engine = pyttsx3.init()
    engine.say("Program is ready")
    engine.runAndWait()

    dict_setup = dict(meas_time           = 12,
                      pmod_scc_no         = sample_name,
                      color               = "RED",
                      ca_mode             = True,
                      user_psg500_op_time = 3,
                      ageing_hour         = 2)
    measurer1 = CS2kMeasurer(**dict_setup)
    measurer1.start()
    engine.say("job done")
    engine.runAndWait()
    engine.stop()
    #toggle sound off
    sound.Sound.mute()

if __name__=='__main__':
    root = tk.Tk()
    root.withdraw()
    samples = [
                # 'NX75_SCC9300174_SKC_DPOP', 
                # 'NX75_SCC9300174_LGE1_DPOP', 
                'NX75_SCC9300174_LGE2_DPOP', 
                # 'NH55_SCC0011827', 
                # 'NH55_SCC0011827', 
                # 'NH55_SCC0011827', 
                # 'NH55_SCC0011827', 
               ]

    for sample in samples:
        user_answer = messagebox.askquestion("cs2k Meas", message=f"{sample} Ready?", icon="warning", type="yesno")
        if user_answer == 'yes':
            cs2k_mease_WRGB(sample_name=sample)

    root.destroy()