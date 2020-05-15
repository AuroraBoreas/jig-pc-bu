import time, pyttsx3
from CS2000_package.CS2000_SW_Control import main_meas_IRE

engine = pyttsx3.init()
engine.say("Program is ready")
engine.runAndWait()

meas_time = 12 #<~ give cs2000_sw 22 seconds to measure
SCC_No = 'NX85_CS_SCC9300101'
IRE = "100"
CA_Mode = False
User_PSG500_op_time = 12
main_meas_IRE(meas_time, SCC_No, IRE, CA_Mode, User_PSG500_op_time)

engine.say("start measure  {} seconds later".format(10))
engine.runAndWait()
time.sleep(10)

meas_time = 20 #<~ give cs2000_sw N seconds to measure
SCC_No = 'NX85_CS_SCC9300101'
IRE = "20"
CA_Mode = False
User_PSG500_op_time = 12
main_meas_IRE(meas_time, SCC_No, IRE, CA_Mode, User_PSG500_op_time)

engine.say("job done")
engine.runAndWait()
engine.stop()
