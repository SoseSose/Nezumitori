#%%
from PySide6 import QtWidgets, QtGui, QtCore
from PySide6.QtGui import QScreen
from PySide6.QtWidgets import QApplication, QLabel
from PySide6.QtCore import Qt
import win32gui, win32process, win32con
import uiautomation as ui
import mouse
import keyboard
import msaa
from comtypes import COMError
import psutil

dpr = 1.0
#uiが多すぎる時、全てが「.z」になるので対策を
#configファイルを作成できるように
#場所でクリックできる機能を追加できるように


class keyCloseLabel(QLabel):
    def __init__(self, l, t, w, h, key, w_handle, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.key = key
        self.w_handle = w_handle
        labelStyle = """QLabel {
            background-color: #00000000;
            font-size:        12px;     
        }"""

        self.setStyleSheet(labelStyle)
        self.move(l, t)
        self.resize(w, h)
        lineStyle = """QLabel {
            background-color: #AAAA0000; 
        }"""

        def draw_kaku(x, y, w, h):
            kaku = QLabel(self)
            kaku.setGeometry(x, y, w, h)
            kaku.setStyleSheet(lineStyle)

        line_w = 2
        draw_kaku(0, 0, line_w, h / 3)
        draw_kaku(0, 0, w / 3, line_w)
        # 左側、下側はline_wに+1の補正があった方が自然になる。
        draw_kaku(0, h * 2 / 3, line_w, h)
        draw_kaku(0, h - (line_w + 1), w / 3, h)

        draw_kaku(w * 2 / 3, 0, w, line_w)
        draw_kaku(w - (line_w + 1), 0, w, h / 3)

        draw_kaku(w * 2 / 3, h - (line_w + 1), w, h)
        draw_kaku(w - (line_w + 1), h * 2 / 3, w, h)

        text_lbl = QLabel(self)
        text_lbl.move(line_w, line_w)
        text_lbl.setAlignment(Qt.AlignRight | Qt.AlignTop)
        w = w - 2 * line_w if w < 20 else 20
        h = h - 2 * line_w if h < 20 else 20
        text_lbl.resize(w - 2 * line_w, h - 2 * line_w)
        labelStyle = """QLabel {
            background-color: #CC000000;
            color: #FFFA1A;
            font-size:        12px;     
            font-family:MADE Evolve Sans; 
            font-weight:bold;
        }"""
        text_lbl.setStyleSheet(labelStyle)
        text_lbl.setText(key)
        self.text_lbl = text_lbl

    def keyClose(self, key):
        top_key, self.key = self.key[:1], self.key[1:]
        if top_key == key:
            if len(self.key) > 0:
                self.text_lbl.setText(self.key)
                return 1
            else:
                x = (self.x() + int(self.width() / 2)) * dpr
                y = (self.y() + int(self.height() / 2)) * dpr
                win32gui.SetWindowPos(self.w_handle,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                print("click " + top_key, x, y)
                mouse.move(x, y)
                mouse.click("left")
                win32gui.SetWindowPos(self.w_handle, win32con.HWND_NOTOPMOST, 0,
                      0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.deleteLater()
                return 0

        else:
            self.deleteLater()
            return 0


alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
alp_list = [text for text in alphabet]
mod_list = [";", ":", ".", ","]
alp_mod_list = alp_list + [mod + alp for mod in mod_list for alp in alp_list]
alphabet_ind = 0
mod_ind = -1

MAX_HEIGHT = 20
MAX_WIDTH = 20


def scanning(w_handle, display_box):
    # print(w_handle)
    threadid, pid = win32process.GetWindowThreadProcessId(w_handle)
    p_name = psutil.Process(pid).name()
    if p_name in [
        name + ".exe"
        for name in [
            "exploror",
            "lghub",
            "Discord",
        ]
    ]:
        backend = "msaa"
    else:
        backend = "uia"
    print(p_name, backend)
    if backend == "msaa":

        def scanning_ui(obj):
            if obj == None:
                return

            for w_c in obj:
                try:
                    ct = msaa.AccRoleNameMap.get(w_c.accRole(), "Unkown")
                    location = w_c.accLocation()
                    # print(location,ct)
                    if location != (0, 0, 0, 0) and ct in [
                        "PushButton",
                        "MenuItem",
                        "TabItem",
                        "ListItem",
                        "CheckBox",
                        "SplitButton",
                        "TreeItem",
                        "PageTab",
                        "Graphic",
                    ]:
                        display_box(location)
                        pass
                    else:
                        pass
                except COMError as ce:
                    pass
                except:
                    pass
                scanning_ui(w_c)

        ww = msaa.window(w_handle)
        scanning_ui(ww)

    elif backend == "uia":

        def scanning_ui(obj):
            if obj == None:
                return
            ct = obj.ControlTypeName

            rec = obj.BoundingRectangle
            ltwh = (
                rec.left,
                rec.top,
                rec.right - rec.left,
                rec.bottom - rec.top,
            )
            if ct in [
                c_type + "Control"
                for c_type in [
                    "Button",
                    "MenuItem",
                    "ToolItem",
                    "TabItem",
                    "ListItem",
                    "ListViewItem",
                    "ListBoxItem",
                    "CheckBox",
                    "SplitButton",
                    "TreeItem",
                    "TreeViewItem",
                ]
            ]:
                display_box(ltwh)
            # ["Button,SplitButton,MenuBar,MenuItem,ToolBar,ToolItem,ListItem,TreeItem,Thumb,CheckBox,ToolItem,Tab, TabItem, Custom, List,ProgressBar,Image,start,Text"]

            if ct in [c_type + "Control" for c_type in ["Edit"]]:
                display_box(ltwh)

            objs = obj.GetChildren()
            for x in objs:
                scanning_ui(x)

        w_ctrl = ui.ControlFromHandle(w_handle)
        scanning_ui(w_ctrl)


class Window(QtWidgets.QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        print("start")

        title = "nezumitori"
        self.setWindowTitle(title)
        self.setStyleSheet("background-color:white;")
        ps = QApplication.primaryScreen()
        global dpr
        dpr = QScreen.devicePixelRatio(ps)
        ww = QScreen.availableVirtualGeometry(ps).width()
        wh = QScreen.availableVirtualGeometry(ps).height()
        self.setGeometry(0, 0, ww, wh)
        self.setFixedSize(ww, wh)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        # self.setWindowFlags(QtCore.Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)

    def show_boxes(self):
        self.keyIndex = 0
        w_handle = ui.GetForegroundWindow()

        def show_box(ltwh):
            l, t, w, h = ltwh
            # なぜかww,whがおかしくて、dprで割って補正しないといけない。
            l /= dpr
            t /= dpr
            w /= dpr
            h /= dpr
            if w < MAX_WIDTH:
                w = MAX_WIDTH
            if h < MAX_HEIGHT:
                h = MAX_HEIGHT

            key = alp_mod_list[self.keyIndex]
            if self.keyIndex < len(alp_mod_list) - 1:
                self.keyIndex += 1
            keyCloseLabel(l, t, w, h, key=key, w_handle = w_handle, parent=self)

        scanning(w_handle, show_box)


    def hideEvent(self, event):
        keyboard.wait("ctrl + shift + l", suppress=True, trigger_on_release=True)
        self.show_boxes()
        print("show")
        self.show()
        # win32gui.SetForegroundWindow(self.winId())
        win32gui.SetWindowPos(self.winId(),win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        self.setFocus()

    # def keyReleaseEvent(self, event):
    def keyPressEvent(self, event):

        if event.isAutoRepeat():
            return
        key = QtGui.QKeySequence(event.key()).toString(QtGui.QKeySequence.NativeText)
        print("Released %s" % key)
        dont_stop_flag = 0
        for box in self.children():
            dont_stop_flag += box.keyClose(key)

        if dont_stop_flag == 0:
            print("hide")
            self.hide()


if __name__ == "__main__":

    app = QtWidgets.QApplication()

    window = Window()
    window.show()
    window.hide()
    app.exec()
