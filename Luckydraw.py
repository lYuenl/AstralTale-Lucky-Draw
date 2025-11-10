from PyQt5.QtWidgets import QApplication, QLabel, QMainWindow, QWidget
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPixmap, QImage, QIcon
from PIL import ImageGrab
from GUI import Ui_MainWindow
import cv2
import numpy as np
import mouse
import keyboard
import time
import win32api
import win32con
from win32com import client
import win32gui
import pyautogui
import pythoncom
import ddddocr
import threading
import os
import sys

matchList = []

def MouseMove(x, y):
    mouse.move(x, y, absolute = True, duration = 0.002)
    time.sleep(0.01)

def LeftClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.02)

def RightClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
    time.sleep(0.01)

def SetForegroundWindow(hwnd):
    if hwnd:
        try:
            win32gui.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 1|2)
            win32gui.SetWindowPos(hwnd, -2, 0, 0, 0, 0, 1|2)
            pythoncom.CoInitialize()
            shell = client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            print(f"設置窗口錯誤: {str(e)}")

def HideLogOutput():
    Log_hwnd = win32gui.FindWindow(None, "Log Output")
    if Log_hwnd:
        win32gui.ShowWindow(Log_hwnd, win32con.SW_HIDE)

class ToggleBorderLabel(QLabel):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setStyleSheet("border: 1px solid black;")
        self.color = None
        self.clicked = False
        self.rightclicked = False

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked = not self.clicked
            #color = "orange" if self.clicked else "black"
            #self.setStyleSheet(f"border: 2px solid {color};")
            if self.rightclicked:
                self.rightclicked = not self.rightclicked
                matchList.remove((self.objectName(), self.pixmap()))
            if self.clicked:
                self.color = "black"
                self.setStyleSheet(f"border: 4px solid black;")
                matchList.append((self.objectName(), self.pixmap()))
            else:
                self.color = None
                self.setStyleSheet(f"border: 1px solid black;")
                matchList.remove((self.objectName(), self.pixmap()))
                
        elif event.button() == Qt.RightButton:
            self.rightclicked = not self.rightclicked
            if self.clicked:
                self.clicked = not self.clicked
                matchList.remove((self.objectName(), self.pixmap()))
            if self.rightclicked:
                self.color = "red"
                self.setStyleSheet(f"border: 4px solid red;")
                matchList.append((self.objectName(), self.pixmap()))
            else:
                self.color = None
                self.setStyleSheet(f"border: 1px solid black;")
                matchList.remove((self.objectName(), self.pixmap()))

class MainWindow(QMainWindow):
    def __init__(self):
        global AstralTale_window
        super().__init__()
        self.ui = Ui_MainWindow()
        iconPath = os.path.join(os.path.dirname(__file__), "AstralTale.ico")
        self.setWindowIcon(QIcon(iconPath))
        self.ui.setupUi(self)
        self.setFixedSize(351, 454) #351, 435
        self.ui.topCheckBox.stateChanged.connect(self.toggle_on_top)
        self.ui.topCheckBox.setChecked(True)
        self.ui.StartButton.clicked.connect(self.Start)
        self.ocr = ddddocr.DdddOcr(beta = True, show_ad = False)
        self.isInitial = False
        self.isStart = False
        self.current_img_data_position = None
        self.current_img_data = None

        self.img_data_position = []
        self.img_data = []
        try:
            AstralTale_window = pyautogui.getWindowsWithTitle("Astral Realm Online")[0]
            self.AstralTale_hwnd = AstralTale_window._hWnd
        except Exception as e:
            self.ui.StartButton.setText("獲取hWnd失敗")

    def closeEvent(self, event): #關閉視窗時強制退出
        os._exit(1)
        
    def Initial(self):
        try:
            HideLogOutput()
            initial_position = self.FindImg("initial_demo.png") # x+14, y+65
            
            self.ui.StartButton.setText("開始煉金")
            initial_position[0] = initial_position[0] + 14
            initial_position[1] = initial_position[1] + 65

            start_button_pos = self.FindRGBImg("start_button.png")
            self.start_button_X = start_button_pos[0] + 42
            self.start_button_Y = start_button_pos[1] + 15

            get_item_pos = self.FindImg("get_item_disable.png")
            self.get_item_X = get_item_pos[0] + 42
            self.get_item_Y = get_item_pos[1] + 15

            self.AstralShard_Text, self.AstralStone_Text = self.AstralStoneOCR()
            self.UpdateUiText()
            
            self.AddImgPosition(initial_position)

            # 把 42 個 img1 ~ img42 替換成可點擊的版本
            for i in range(1, 43):
                label = getattr(self.ui, f"img{i}")  # 取得 QLabel 物件
                toggle_label = ToggleBorderLabel(self)  # 建立可點擊的 label
                toggle_label.setText(label.text())
                toggle_label.setGeometry(label.geometry())
                toggle_label.setAlignment(label.alignment())
                toggle_label.setPixmap(self.img_data[i - 1])
                toggle_label.setScaledContents(True)

                label.hide()
                toggle_label.setObjectName(f"img{i}")
                setattr(self.ui, f"img{i}", toggle_label)  # 替換成新的物件
                toggle_label.setParent(self)
                toggle_label.show()

            disable_img_names = ["img1", "img2", "img3", "img4", "img5", "img6", "img7", "img13", "img19", "img25", "img31", "img37"]

            for name in disable_img_names:
                label = self.findChild(QLabel, name)
                if label:
                    label.mousePressEvent = lambda event: None

            self.isInitial = True
        except Exception as e:
            self.ui.StartButton.setText("初始化失敗")
            self.isInitial = False

    def Start(self):
        if not self.isInitial:
            self.Initial()
        else:
            if not pause:
                self.isStart = True
                SetForegroundWindow(self.AstralTale_hwnd)
                self.ui.StartButton.setText("自動煉金中")
            if self.isStart:
                time.sleep(0.5)
                while True:
                    if not pause:
                        self.AstralShard_Text, self.AstralStone_Text = self.AstralStoneOCR()
                        self.UpdateUiText()
                        QApplication.processEvents()
                        if self.MatchItem():
                            MouseMove(x = self.get_item_X, y = self.get_item_Y)
                            LeftClick()
                            MouseMove(x = self.get_item_X, y = self.get_item_Y + 40)
                            #print("領取物品")
                            time.sleep(0.5)
                            if self.FindRGBImg("get_item_enable.png"):
                                get_all_item_X, get_all_item_Y = self.FindRGBImg("get_all_item_enable.png")
                                MouseMove(x = get_all_item_X, y = get_all_item_Y)
                                LeftClick()
                                time.sleep(1)
                                MouseMove(x = self.get_item_X, y = self.get_item_Y)
                                LeftClick()
                                MouseMove(x = self.get_item_X, y = self.get_item_Y + 40)

                        elif self.FindRGBImg("continue_button_enable.png") and self.FindRGBImg("get_item_enable.png"):
                            if self.FindRGBImg("discard_msg.png"):
                                discard_ok_X, discard_ok_Y = self.FindRGBImg("discard_ok.png")
                                MouseMove(x = discard_ok_X + 42, y = discard_ok_Y + 16)
                                LeftClick()
                                MouseMove(x = discard_ok_X + 42, y = discard_ok_Y + 200)
                                time.sleep(1.7)
                                continue
                            MouseMove(x = self.start_button_X, y = self.start_button_Y)
                            LeftClick()
                            MouseMove(x = self.start_button_X, y = self.start_button_Y + 40)
                            #print("繼續練金")
                        elif self.FindRGBImg("continue_button_disable.png") and self.FindRGBImg("get_item_enable.png"):
                            MouseMove(x = self.get_item_X, y = self.get_item_Y)
                            LeftClick()
                            MouseMove(x = self.get_item_X, y = self.get_item_Y + 40)
                            time.sleep(0.5)
                            if self.FindRGBImg("get_item_enable.png"):
                                get_all_item_X, get_all_item_Y = self.FindRGBImg("get_all_item_enable.png")
                                MouseMove(x = get_all_item_X, y = get_all_item_Y)
                                LeftClick()
                                time.sleep(1)
                                MouseMove(x = self.get_item_X, y = self.get_item_Y)
                                LeftClick()
                                MouseMove(x = self.get_item_X, y = self.get_item_Y + 40)
                            #print("領取碎塊")
                        elif self.FindRGBImg("get_item_disable.png") and self.FindRGBImg("start_button_disable.png"):
                            pauseProcess(None)
                        else:
                            MouseMove(x = self.start_button_X, y = self.start_button_Y)
                            LeftClick()
                            MouseMove(x = self.start_button_X, y = self.start_button_Y + 40)

                        self.UpdateUiText()
                        QApplication.processEvents()
                        time.sleep(1.7)

                    QApplication.processEvents()
                    time.sleep(0.01)

    def MatchItem(self):
        initial_position = self.FindImg("initial_demo.png")
        initial_position[0] = initial_position[0] + 14
        initial_position[1] = initial_position[1] + 65
        self.current_img_data_position = initial_position[0] - 24, initial_position[1] + 429
        if matchList and initial_position:
            for objectName, matchImg in matchList:
                current_img = ImageGrab.grab(bbox = (self.current_img_data_position[0], self.current_img_data_position[1], self.current_img_data_position[0] + 44, self.current_img_data_position[1] + 45)).convert("RGB")
                width = current_img.width
                height = current_img.height
                rgb_current_img = cv2.cvtColor(np.array(current_img).reshape(height, width, 3), cv2.COLOR_BGR2RGB)
                #cv2.imwrite("output.png", rgb_current_img)
                if self.MatchTemplate(rgb_current_img, self.QPixmapToRGBImage(matchImg)):
                    toggle_label = self.findChild(QLabel, f"{objectName}")
                    if toggle_label.color == "black":
                        toggle_label.clicked = False
                        toggle_label.color = None
                        toggle_label.setStyleSheet(f"border: 1px solid black;")
                        matchList.remove((objectName, matchImg))

                    return True
            return False
        else:
            #print("尚未選擇任何物品")
            return False

    def AddImgPosition(self, initial_position):
        offsetX = [0, 55, 55, 56, 55, 55]
        offsetY = [0, 56, 56, 56, 56, 56, 55]
        initial_position_X = initial_position[0]
        initial_position_Y = initial_position[1]
        for Y in offsetY:
            initial_position[1] = initial_position[1] + Y

            for X in offsetX:
                initial_position[0] = initial_position[0] + X
                img = ImageGrab.grab(bbox = (initial_position[0], initial_position[1], initial_position[0] + 44, initial_position[1] + 45)).convert("RGB")
                w, h = img.size
                qimg = QImage(img.tobytes(), w, h, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(qimg)

                self.img_data_position.append(initial_position)
                self.img_data.append(pixmap)

            initial_position[0] = initial_position_X
        
        self.current_img_data_position = initial_position_X - 24, initial_position_Y + 429

    def QPixmapToRGBImage(self, pixmap):   
        qimg = pixmap.toImage().convertToFormat(QImage.Format_RGB888)
        width = qimg.width()
        height = qimg.height()
        ptr = qimg.bits()
        ptr.setsize(qimg.byteCount())
        img = np.array(ptr).reshape(height, width, 3)
        rgb_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        #cv2.imwrite("1.png", rgb_img)
        return rgb_img

    def MatchTemplate(self, img1, img2):
        res = cv2.matchTemplate(img1, img2, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        precision = int(round(max_val, 2) * 100)

        if precision == 100:
            return True
        else:
            return False
        
    def FindImg(self, imgfile: str):
        try:
            screenshot = ImageGrab.grab()
            gray_screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2GRAY)
            imgfile = os.path.join(os.path.dirname(__file__), imgfile)
            template_img = cv2.imdecode(np.fromfile(imgfile, dtype = np.uint8), -1)
            #template_img = cv2.imread(os.path.join(os.path.dirname(__file__), imgfile))
            gray_template = cv2.cvtColor(template_img, cv2.COLOR_BGR2GRAY)
            res = cv2.matchTemplate(gray_screenshot, gray_template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            top_left = max_loc
            precision = int(round(max_val, 2) * 100)
            if precision >= 90:
                return list(top_left)
            else:
                return None
        except Exception as e:
            os._exit(1)
        
    def FindRGBImg(self, imgfile: str):
        try:
            screenshot = ImageGrab.grab().convert("RGB")
            rgb_screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2RGB)
            imgfile = os.path.join(os.path.dirname(__file__), imgfile)
            template_img = cv2.imdecode(np.fromfile(imgfile, dtype = np.uint8), -1)
            #template_img = cv2.imread(os.path.join(os.path.dirname(__file__), imgfile))
            rgb_template = cv2.cvtColor(template_img, cv2.COLOR_BGR2RGB)
            res = cv2.matchTemplate(rgb_screenshot, rgb_template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            top_left = max_loc
            precision = int(round(max_val, 2) * 100)
            if precision >= 85:
                return list(top_left)
            else:
                return None
        except Exception as e:
            os._exit(1)

    def AstralStoneOCR(self):
        try:
            AstralShard_Pos = self.FindRGBImg("AstralShard.png")
            AstralShard_Image = ImageGrab.grab(bbox = (AstralShard_Pos[0] + 100, AstralShard_Pos[1] + 7, AstralShard_Pos[0] + 157, AstralShard_Pos[1] + 25))
            AstralShard_Text = str(self.ocr.classification(AstralShard_Image)).replace("o", "0").replace("l", "1").replace("I", "1").replace("i", "1").replace("s", "9").replace(">", "7").replace("u", "11").replace("e", "2").replace("口", "0")

            AstralStone_Pos = self.FindRGBImg("AstralStone.png")
            AstralStone_Image = ImageGrab.grab(bbox = (AstralStone_Pos[0] + 100, AstralStone_Pos[1] + 7, AstralStone_Pos[0] + 157, AstralStone_Pos[1] + 25))
            AstralStone_Text = str(self.ocr.classification(AstralStone_Image)).replace("o", "0").replace("l", "1").replace("I", "1").replace("i", "1").replace("s", "9").replace(">", "7").replace("u", "11").replace("e", "2").replace("口", "0")

            return AstralShard_Text, AstralStone_Text
        except Exception as e:
            os._exit(1)
    
    def UpdateUiText(self):
        self.ui.AstralShard.setText(f"星界碎塊: {self.AstralShard_Text}")
        self.ui.AstralStone.setText(f"星界石: {self.AstralStone_Text}")

    def toggle_on_top(self, state):
        is_on_top = (state == Qt.Checked)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, is_on_top)
        self.show()

def pauseProcess(event):
    global pause
    if window.isInitial:
        pause = not pause
        if pause:
            window.ui.StartButton.setText("暫停煉金")
        else:
            if window.isStart:
                SetForegroundWindow(AstralTale_window._hWnd)
                window.ui.StartButton.setText("自動煉金中")
            else:
                window.ui.StartButton.setText("開始煉金")
        
def closeProcess(event):
    os._exit(1)
        
if __name__ == '__main__':
    pause = False
    keyboard.on_release_key("f8", closeProcess)
    keyboard.on_release_key("f9", pauseProcess)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

#common.onnx: ddddocr
#onnxruntime_providers_shared.dll: onnxruntime\capi