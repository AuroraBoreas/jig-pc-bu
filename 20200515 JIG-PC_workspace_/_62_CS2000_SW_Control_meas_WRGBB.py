"""
=============================
Automate CS2000 measurement @ZL, 2019
=============================
"""
import time, pyttsx3
from CS2000_package.CS2000_SW_Control import main_meas_IRE
from Windows_Sound_Manager_master import sound

##mute then toggle speaker on
sound.Sound.mute()
sound.Sound.volume_min()
sound.Sound.volume_max()

engine = pyttsx3.init()
engine.say("Program is ready")
engine.runAndWait()

meas_time = 12
"""===update the following 3 variables before press F5=="""
SCC_No = 'NX65_SCC0003058_LD34'
IRE = "WRGB"
Ageing_Hour = 2
"""===update the above 3 variables before press F5  ==="""
CA_Mode = True
User_PSG500_op_time = 3


main_meas_IRE( meas_time, SCC_No, IRE, CA_Mode, User_PSG500_op_time, Ageing_Hour)
engine.say("job done")
engine.runAndWait()
engine.stop()

##toggle sound off
sound.Sound.mute()
