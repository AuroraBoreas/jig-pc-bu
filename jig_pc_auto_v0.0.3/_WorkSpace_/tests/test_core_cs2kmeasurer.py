import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg_cs2000_manager.core import CS2kMeasurer

if __name__ == '__main__':
    dict_setup = dict(meas_time           = 12,
                      pmod_scc_no         = "NH55_SCC0011827",
                      color               = "BLue",
                      ca_mode             = False,
                      user_psg500_op_time = 3,
                      ageing_hour         = 2)
    measurer = CS2kMeasurer(**dict_setup)
    measurer.wait_time_blue = 50
    measurer.start()