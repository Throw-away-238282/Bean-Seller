import numpy as np
import pytesseract
import win32gui
import win32ui
import win32con
import win32com.client
import json
from time import sleep

pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'


def list_window_names():
    def win_enum_handler(hwnd):
        if win32gui.IsWindowVisible(hwnd):
            print(hex(hwnd), win32gui.GetWindowText(hwnd))

    win32gui.EnumWindows(win_enum_handler, None)


class WindowCapture:
    # properties
    w = 0
    h = 0
    hwnd = None

    def __init__(self, window_name):
        self.hwnd = win32gui.FindWindow(None, window_name)
        if not self.hwnd:
            list_window_names()
            raise Exception('Window not found: {}'.format(window_name))

        self.left, self.top, self.right, self.bot = win32gui.GetWindowRect(self.hwnd)
        self.w = self.right - self.left
        self.h = self.bot - self.top
        # account for the window border and titlebar and cut them off
        self.border_pixels = 8
        self.titlebar_pixels = 30
        self.w = self.w - (self.border_pixels * 2)
        self.h = self.h - self.titlebar_pixels - self.border_pixels

    def get_screenshot(self):
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(self.hwnd)
        sleep(.3)

        hdesktop = win32gui.GetDesktopWindow()
        hwnd_dc = win32gui.GetWindowDC(hdesktop)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()

        save_bit_map = win32ui.CreateBitmap()
        save_bit_map.CreateCompatibleBitmap(mfc_dc, self.w, self.h)

        save_dc.SelectObject(save_bit_map)

        save_dc.BitBlt((0, 0), (self.w, self.h), mfc_dc, (self.left, self.top), win32con.SRCCOPY)

        # convert the raw data into a format opencv can read
        # save_bit_map.SaveBitmapFile(save_dc, 'debug.bmp')
        signed_ints_array = save_bit_map.GetBitmapBits(True)
        img = np.frombuffer(signed_ints_array, dtype='uint8')
        img.shape = (self.h, self.w, 4)

        # free resources
        mfc_dc.DeleteDC()
        save_dc.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwnd_dc)
        win32gui.DeleteObject(save_bit_map.GetHandle())

        # drop the alpha channel, or cv.matchTemplate() will throw an error like:
        #   error: (-215:Assertion failed) (depth == CV_8U || depth == CV_32F) && type == _templ.type()
        #   && _img.dims() <= 2 in function 'cv::matchTemplate'
        img = img[..., :3]

        # make image C_CONTIGUOUS to avoid errors that look like:
        #   File ... in draw_rectangles
        #   TypeError: an integer is required (got type tuple)
        # see the discussion here:
        # https://github.com/opencv/opencv/issues/14866#issuecomment-580207109
        img = np.ascontiguousarray(img)

        return img


class CaptureItem:

    def __init__(self, screen_shot):
        self.screen_shot = screen_shot
        self.item_x_range_0 = [400, 475]
        self.item_x_range_1 = [475, 580]
        self.item_x_range_2 = [580, 640]
        self.item_x_range_3 = [650, 710]
        self.item_x_range_4 = [730, 790]
        self.item_x_range_5 = [800, 860]
        self.item_x_range_6 = [870, 940]
        self.item_x_range_7 = [960, 1005]
        self.item_x_range_8 = [1030, 1082]
        self.item_x_range_9 = [995, 1040]
        self.item_x_range_10 = [1060, 1120]
        self.item_name = [705, 885]
        self.item_price = [1030, 1090]
        self.item_GS = [1180, 1240]
        self.item_rarity = [1425, 1520]
        self.item_location = [1720, 1810]
        self.items_per_scan = 9

    def image_to_text(self, image):
        return pytesseract.image_to_string(image).strip()

    def get_items(self):
        items_list = []

        for i in range(self.items_per_scan):
            match i:
                case 0:
                    item = self.get_item(self.item_x_range_0)
                    items_list.append(item)
                case 1:
                    item = self.get_item(self.item_x_range_1)
                    items_list.append(item)
                case 2:
                    item = self.get_item(self.item_x_range_2)
                    items_list.append(item)
                case 3:
                    item = self.get_item(self.item_x_range_3)
                    items_list.append(item)
                case 4:
                    item = self.get_item(self.item_x_range_4)
                    items_list.append(item)
                case 5:
                    item = self.get_item(self.item_x_range_5)
                    items_list.append(item)
                case 6:
                    item = self.get_item(self.item_x_range_6)
                    items_list.append(item)
                case 7:
                    item = self.get_item(self.item_x_range_7)
                    items_list.append(item)
                case 8:
                    item = self.get_item(self.item_x_range_8)
                    items_list.append(item)

        return items_list

    def get_last_items(self):
        items_list = []

        item = self.get_item(self.item_x_range_9)
        items_list.append(item)

        item = self.get_item(self.item_x_range_10)
        items_list.append(item)

        return items_list

    def get_item(self, item_range: list):
        item_name_img = self.screen_shot[item_range[0]:item_range[1],
                                         self.item_name[0]:self.item_name[1]]

        item_price_img = self.screen_shot[item_range[0]:item_range[1],
                                          self.item_price[0]:self.item_price[1]]

        item_gs_img = self.screen_shot[item_range[0]:item_range[1],
                                       self.item_GS[0]:self.item_GS[1]]

        item_rarity_img = self.screen_shot[item_range[0]:item_range[1],
                                           self.item_rarity[0]:self.item_rarity[1]]

        item_location_img = self.screen_shot[item_range[0]:item_range[1],
                                             self.item_location[0]:self.item_location[1]]

        item_name = self.image_to_text(item_name_img)
        item_price = self.image_to_text(item_price_img)
        item_gs = self.image_to_text(item_gs_img)
        item_rarity = self.image_to_text(item_rarity_img)
        item_location = self.image_to_text(item_location_img)

        item = {"item_name": item_name,
                "item_price": item_price,
                "item_gs": item_gs,
                "item_rarity": item_rarity,
                "item_location": item_location}
        return item


items = []
wincap = WindowCapture('New World')
while True:
    ans = input("\n Press any key for a full scrape, press w to get the last 2 items, press q to quit. \n")
    if ans.lower() == 'q':
        break
    elif ans.lower() == 'w':
        gameshot = wincap.get_screenshot()
        itemcap = CaptureItem(gameshot)
        items.extend(itemcap.get_last_items())
        continue

    gameshot = wincap.get_screenshot()
    itemcap = CaptureItem(gameshot)

    items.extend(itemcap.get_items())

json = json.dumps(items, indent=2)
with open("market.json", "w") as file1:
    # Writing data to a file
    file1.write(json)
