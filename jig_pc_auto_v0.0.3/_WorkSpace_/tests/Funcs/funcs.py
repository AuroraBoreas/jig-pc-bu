"""====================================================================================================================
just like PERSONAL.XLSB in VBA, I wanna collect all general-purpose functions I made during daily Python Programming.
Author: @ZL, 2020
===================================================================================================================="""


import datetime, pathlib
import hashlib
import shutil

import platform, socket, re, uuid, json, psutil, logging, sys, ctypes

def get_weeknum(d):
    """return a string to display a given day. argument d is python datetime obj"""
    return '{:02d}'.format(datetime.date.today().isocalendar()[1])

def get_current_weeknum():
    """return a string to display current weeknum"""
    return '{:02d}'.format(datetime.date.today().isocalendar()[1])

def get_today_date():
    """return a string to display military datetime format"""
    return datetime.datetime.now().strftime('%Y%m%d')

def is_expired(expired_date='Dec 1 2020 8:00AM'):
    """check if current date is explired date"""
    ed = datetime.datetime.strptime(expired_date,'%b %d %Y %I:%M%p')
    now = datetime.datetime.now()
    return (now - ed) > datetime.timedelta(days=1)

def is_outdated(file_path, days=1):
    """check if file modified date is outdated. default is 1 day outdated"""
    mtime = pathlib.Path(file_path).stat().st_mtime
    mtime = datetime.datetime.fromtimestamp(mtime)
    now = datetime.datetime.now()
    return (now - mtime) > datetime.timedelta(days=days)

def hash_file(file_path):
    """display hash strings of a give file path"""
    # BUF_SIZE is totally arbitrary, change for your app!
    BUF_SIZE = 65536  # lets read stuff in 64kb chunks!
    md5 = hashlib.md5()
    sha1 = hashlib.sha1()
    with open(file_path, 'rb') as f:
        while True:
            data = f.read(BUF_SIZE)
            if not data:
                break
            md5.update(data)
            sha1.update(data)
    print("MD5: {0}".format(md5.hexdigest()))
    print("SHA1: {0}".format(sha1.hexdigest()))

def copy_src(filepath, store_fd):
    """copy file into a give folder"""
    shutil.copy(filepath, store_fd)
    return

def getSystemInfo():
    """get development environment"""
    try:
        info = {}
        info['platform']         = platform.system()
        info['platform-release'] = platform.release()
        info['platform-version'] = platform.version()
        info['architecture']     = platform.machine()
        info['hostname']         = socket.gethostname()
        info['ip-address']       = socket.gethostbyname(socket.gethostname())
        info['mac-address']      = ':'.join(re.findall('..', '%012x' % uuid.getnode()))
        info['processor']        = platform.processor()
        info['ram']              = str(round(psutil.virtual_memory().total / (1024.0 **3)))+" GB"
        info['python -V']        = sys.version
        info['screensize']       = ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1)
        return json.dumps(info)
    except Exception as e:
        logging.exception(e)

def change_encoding_for_py():
    import sys, io
    """modify default python i/o setting to display Chinese characters correctly"""
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

if __name__ == '__main__':
    # assert is_expired() == False
    # print(get_current_weeknum())
    # assert get_weeknum(datetime.datetime.strptime('2020-05-29 10:09:30.243860', '%Y-%m-%d %H:%M:%S.%f')) == '23'

    info = json.loads(getSystemInfo())
    for k,v in info.items():
        # Markdown output
        print("* {0:<16} : {1:}".format(k, v))