import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg_Windows_Sound_Manager import sound

#mute then toggle speaker on
sound.Sound.mute()