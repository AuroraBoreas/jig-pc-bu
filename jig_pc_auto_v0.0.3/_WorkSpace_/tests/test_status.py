import pyautogui

def get_rgb_under_mouse_cursor(x, y):
    """
    return r,g,b value of pixel under mouse cursor (x=337, y=191)
    """
    from ctypes import windll
    pyautogui.moveTo(x, y) #color changes when mouse event occurs
    dc = windll.user32.GetDC(0)
    rgb = windll.gdi32.GetPixel(dc, x, y)
    r = rgb & 0xff
    g = (rgb >> 8) & 0xff
    b = (rgb >> 16) & 0xff
    return (r, g, b)

# cs2000_warning_tag_before_meas
x, y = 620, 381
r, g, b = 234, 174, 0
# warning_tag_after_meas
x, y = 602, 381
r, g, b = 236, 175, 0

def IsDone(x, y):
    color_flag = 230 #background color of prompt window
    if abs(get_rgb_under_mouse_cursor(x, y)[0] - color_flag) <= 10:
        print("Done")
        return True
    print("NotDone")
    return False

IsDone(x, y)

def IsConnected(point):
    color_flag = 176
    x, y = 575, 361
    