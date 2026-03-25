import numpy as np
import cv2
import board_recognition as br
import pyperclip
import time
import sys
from typing import Optional, Any, cast, TYPE_CHECKING

from PySide6.QtWidgets import (QApplication, QFileDialog, QDialog, QVBoxLayout, 
                             QPushButton, QLabel, QWidget, QRadioButton, QGroupBox, QHBoxLayout)
from PySide6.QtCore import Qt, QRect, QPoint, Signal, QObject
from PySide6.QtGui import QPainter, QColor, QScreen, QPixmap, QImage
import threading
import ctypes
from ctypes import wintypes


class HotkeyEmitter(QObject):
    hotkey_pressed = Signal()


def _start_windows_hotkey_listener(emitter: HotkeyEmitter, holder: dict):
    """スレッド内で Windows のメッセージループを開始し、F4 押下時に emitter.hotkey_pressed を emit する。"""
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    # F4 の仮想キーコードは 0x73
    VK_F4 = 0x73
    HOTKEY_ID = 1
    WM_HOTKEY = 0x0312

    # スレッド ID を共有オブジェクトに保存
    thread_id = kernel32.GetCurrentThreadId()
    holder["thread_id"] = thread_id

    if not user32.RegisterHotKey(None, HOTKEY_ID, 0, VK_F4):
        err = kernel32.GetLastError()
        print(f"RegisterHotKey failed (VK={hex(VK_F4)}), GetLastError={err}")
        return

    try:
        msg = wintypes.MSG()
        while True:
            res = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if res == 0:
                break
            if msg.message == WM_HOTKEY:
                try:
                    emitter.hotkey_pressed.emit()
                except Exception:
                    pass
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
    finally:
        user32.UnregisterHotKey(None, HOTKEY_ID)


def is_libreoffice_calc_focused() -> bool:
    """前面ウィンドウのタイトルに LibreOffice Calc が含まれるか簡易判定する。"""
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return False
    buf = ctypes.create_unicode_buffer(512)
    length = user32.GetWindowTextW(hwnd, buf, 512)
    title = buf.value[:length]
    if not title:
        return False
    t = title.lower()
    return ("libreoffice" in t and "calc" in t) or ("calc" in t and "libreoffice" in t) or ("libreoffice calc" in t)


def find_egaroucid_window() -> Optional[int]:
    """Egaroucid アプリのウィンドウハンドルを探す。"""
    user32 = ctypes.windll.user32
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    found_hwnd = []

    def callback(hwnd, lparam):
        if user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(hwnd, buf, 512)
            if "egaroucid" in buf.value.lower():
                found_hwnd.append(hwnd)
                return False
        return True

    user32.EnumWindows(WNDENUMPROC(callback), 0)
    return found_hwnd[0] if found_hwnd else None


def activate_window(hwnd: int):
    """指定したハンドルのウィンドウを最前面に持ってくる。"""
    user32 = ctypes.windll.user32
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    user32.SetForegroundWindow(hwnd)


def send_ctrl_v():
    """Send Ctrl+V to the active window using keybd_event (simple, reliable).

    Note: keybd_event is deprecated but acceptable for this small helper.
    """
    user32 = ctypes.windll.user32
    VK_CONTROL = 0x11
    VK_V = 0x56
    KEYEVENTF_KEYUP = 0x0002

    # press ctrl
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    # press v
    user32.keybd_event(VK_V, 0, 0, 0)
    # small delay to ensure target app receives sequence
    time.sleep(0.05)
    # release v
    user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
    # release ctrl
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


def send_down_key():
    """Send a single Down-arrow key press using keybd_event."""
    user32 = ctypes.windll.user32
    VK_DOWN = 0x28
    KEYEVENTF_KEYUP = 0x0002

    # press Down
    user32.keybd_event(VK_DOWN, 0, 0, 0)
    # short delay
    time.sleep(0.02)
    # release Down
    user32.keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)

if TYPE_CHECKING:
    # Type-checker-only fallbacks for attributes that some PySide6 stubs may miss
    class _QtStub:  # pragma: no cover
        WindowStaysOnTopHint: int
        FramelessWindowHint: int
        Tool: int
        CrossCursor: int
        LeftButton: int

    class _QDialogStub:  # pragma: no cover
        Accepted: int

    class _QImageStub:  # pragma: no cover
        Format_RGBA8888: int

    Qt = _QtStub()  # type: ignore
    QDialog = _QDialogStub  # type: ignore
    QImage = _QImageStub  # type: ignore

class SelectionDialog(QDialog):
    def __init__(self, initial_source="capture", initial_format="text", initial_turn="auto", last_rect=None):
        super().__init__()
        self.setWindowTitle("Reversi to Clipboard")  # type: ignore[attr-defined]
        self.source = initial_source
        self.format = initial_format
        self.turn = initial_turn
        self.last_rect = last_rect
        
        layout = QVBoxLayout()
        
        # Source Selection
        source_group = QGroupBox("Image Source")
        source_layout = QVBoxLayout()
        self.cap_radio = QRadioButton("Screen Capture (Rectangle)")
        self.last_cap_radio = QRadioButton("Screen Capture (the last rectangle)")
        if self.last_rect is None:
            self.last_cap_radio.setEnabled(False)
        self.file_radio = QRadioButton("Open Image File")

        if self.source == "file":
            self.file_radio.setChecked(True)
        elif self.source == "last_capture" and self.last_rect is not None:
            self.last_cap_radio.setChecked(True)
        else:
            self.cap_radio.setChecked(True)
            
        source_layout.addWidget(self.cap_radio)
        source_layout.addWidget(self.last_cap_radio)
        source_layout.addWidget(self.file_radio)
        source_group.setLayout(source_layout)
        layout.addWidget(source_group)
        
        # Format Selection
        format_group = QGroupBox("Output Format")
        format_layout = QVBoxLayout()
        self.text_radio = QRadioButton("Text Format (-XO)")
        self.sgf_radio = QRadioButton("SGF Format")
        self.ggf_radio = QRadioButton("GGF Format")

        if self.format == "text":
            self.text_radio.setChecked(True)
        elif self.format == "sgf":
            self.sgf_radio.setChecked(True)
        elif self.format == "ggf":
            self.ggf_radio.setChecked(True)

        format_layout.addWidget(self.text_radio)
        format_layout.addWidget(self.sgf_radio)
        format_layout.addWidget(self.ggf_radio)
        format_group.setLayout(format_layout)
        layout.addWidget(format_group)

        # Turn Selection
        turn_group = QGroupBox("Next Turn")
        turn_layout = QVBoxLayout()
        self.auto_turn_radio = QRadioButton("Auto")
        self.auto_turn_radio.setShortcut("A")
        self.black_turn_radio = QRadioButton("Black (X)")
        self.black_turn_radio.setShortcut("B")
        self.white_turn_radio = QRadioButton("White (O)")
        self.white_turn_radio.setShortcut("W")

        if self.turn == "black":
            self.black_turn_radio.setChecked(True)
        elif self.turn == "white":
            self.white_turn_radio.setChecked(True)
        else:
            self.auto_turn_radio.setChecked(True)

        turn_layout.addWidget(self.auto_turn_radio)
        turn_layout.addWidget(self.black_turn_radio)
        turn_layout.addWidget(self.white_turn_radio)
        turn_group.setLayout(turn_layout)
        layout.addWidget(turn_group)
        
        # Buttons
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("OK (F4)")
        ok_btn.setFixedHeight(ok_btn.sizeHint().height() * 3)
        ok_btn.clicked.connect(self.accept_settings)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)  # type: ignore[attr-defined]
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)  # type: ignore[attr-defined]

    def accept_settings(self):
        if self.file_radio.isChecked():
            self.source = "file"
        elif self.last_cap_radio.isChecked():
            self.source = "last_capture"
        else:
            self.source = "capture"
            
        if self.text_radio.isChecked():
            self.format = "text"
        elif self.sgf_radio.isChecked():
            self.format = "sgf"
        elif self.ggf_radio.isChecked():
            self.format = "ggf"

        if self.black_turn_radio.isChecked():
            self.turn = "black"
        elif self.white_turn_radio.isChecked():
            self.turn = "white"
        else:
            self.turn = "auto"
        self.accept()  # type: ignore[attr-defined]

class CaptureOverlay(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)  # type: ignore[attr-defined]
        self.setWindowOpacity(0.3)
        self.setStyleSheet("background-color: black;")
        self.setCursor(Qt.CrossCursor)  # type: ignore[attr-defined]
        
        # 全画面表示
        screen = QApplication.primaryScreen()
        self.setGeometry(screen.geometry())
        
        self.start_pos: Optional[QPoint] = None
        self.end_pos: Optional[QPoint] = None
        self.capture_pixmap: Optional[QPixmap] = None
        self.is_selecting = False

    def paintEvent(self, event):
        if self.is_selecting and self.start_pos and self.end_pos:
            painter = QPainter(self)
            painter.setPen(QColor(255, 255, 255))
            painter.setBrush(QColor(255, 255, 255, 50))
            rect = QRect(self.start_pos, self.end_pos).normalized()
            painter.drawRect(rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:  # type: ignore[attr-defined]
            self.start_pos = event.position().toPoint()
            self.is_selecting = True

    def mouseMoveEvent(self, event):
        if self.is_selecting:
            self.end_pos = event.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.is_selecting:  # type: ignore[attr-defined]
            self.end_pos = event.position().toPoint()
            self.is_selecting = False
            self.capture_area()
            self.close()

    def capture_area(self):
        # start_pos/end_pos が None の場合は何もしない
        if self.start_pos is None or self.end_pos is None:
            return

        # 選択領域を別名で保持（QWidget.rect と衝突しないようにする）
        self.selection_rect: QRect = QRect(self.start_pos, self.end_pos).normalized()
        if self.selection_rect.width() < 10 or self.selection_rect.height() < 10:
            return

        screen = QApplication.primaryScreen()
        self.hide() # 自分自身を隠してキャプチャ
        QApplication.processEvents() # 再描画を待機
        self.capture_pixmap = screen.grabWindow(0, self.selection_rect.x(), self.selection_rect.y(), self.selection_rect.width(), self.selection_rect.height())

# QApplication インスタンスを作成
app = QApplication(sys.argv)

def show_dialog(default_source="capture", default_format="text", default_turn="auto", last_rect=None):
    # 設定の選択
    dialog = SelectionDialog(default_source, default_format, default_turn, last_rect)

    # F4 をグローバルホットキーとして登録し、押下時にダイアログの OK 相当処理を呼ぶ
    emitter = HotkeyEmitter()
    emitter.hotkey_pressed.connect(dialog.accept_settings)
    holder: dict = {}
    listener_thread = threading.Thread(target=_start_windows_hotkey_listener, args=(emitter, holder), daemon=True)
    listener_thread.start()

    if dialog.exec() != QDialog.Accepted:  # type: ignore[attr-defined]
        # ダイアログ終了時はホットキースレッドを停止させる
        if "thread_id" in holder:
            ctypes.windll.user32.PostThreadMessageW(holder["thread_id"], 0x0012, 0, 0)  # WM_QUIT
        listener_thread.join(timeout=0.5)
        sys.exit()

    image = None
    if dialog.source == "file":
        # ダイアログボックスを表示し、画像ファイルを読み込む
        file_name = QFileDialog.getOpenFileName(None, "Open Image", "", "Image Files (*.png *.jpg *.jpeg *.bmp *.gif)")[0]
        if not file_name:
            print("No file selected.")
            sys.exit()
        image = cv2.imread(file_name)
    elif dialog.source == "last_capture" and last_rect is not None:
        # 前回のキャプチャ座標で直接取得
        screen = QApplication.primaryScreen()
        pixmap = screen.grabWindow(0, last_rect.x(), last_rect.y(), last_rect.width(), last_rect.height())
        # QPixmap を numpy 形式 (OpenCV) に変換
        from PySide6.QtGui import QImage
        qimage = pixmap.toImage().convertToFormat(QImage.Format_RGBA8888)  # type: ignore[attr-defined]
        width = qimage.width()
        height = qimage.height()
        ptr = qimage.bits()
        arr = np.frombuffer(ptr, np.uint8).reshape((height, width, 4))
        image = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR)
    else:
        # スクリーンキャプチャ
        overlay = CaptureOverlay()
        overlay.showFullScreen()
        # イベントループを回して終了を待つ
        while overlay.isVisible():
            app.processEvents()
        
        if overlay.capture_pixmap:
            last_rect = overlay.selection_rect
            # QPixmap を numpy 形式 (OpenCV) に変換
            from PySide6.QtGui import QImage
            qimage = overlay.capture_pixmap.toImage().convertToFormat(QImage.Format_RGBA8888)  # type: ignore[attr-defined]
            width = qimage.width()
            height = qimage.height()
            ptr = qimage.bits()
            # ptr (memoryview) を numpy 配列に変換
            arr = np.frombuffer(ptr, np.uint8).reshape((height, width, 4))
            # RGBA -> BGR
            image = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR)
        else:
            print("No area selected.")
            sys.exit()

    # ダイアログ終了後はホットキーリスナーを停止
    if "thread_id" in holder:
        ctypes.windll.user32.PostThreadMessageW(holder["thread_id"], 0x0012, 0, 0)  # WM_QUIT
        listener_thread.join(timeout=0.5)

    if image is None or image.size == 0:
        print("Failed to load/capture image.")
        sys.exit()


    # 解析に当たってのヒント情報
    hint = br.Hint()
    hint.mode = br.Mode.PHOTO # 写真モード  # type: ignore[attr-defined]

    # 認識に使用するクラス。ここでは自動認識を使用する
    # 自動認識の場合、実盤/スクリーンショット/白黒書籍を自動判断する
    recognizer = br.AutomaticRecognizer()

    # ※うまく行かない場合は、個別のクラスを明示的に使用することも可能
    # recognizer = br.RealBoardRecognizer()  実盤用
    # recognizer = br.ScreenshotRecognizer()  スクリーンショット用
    # recognizer = br.PrintedBoardRecognizer()  白黒書籍用

    # 認識実行
    try:
        ret, result = recognizer.analyzeBoard(image, hint)
    except Exception as e:
        print(f"Error: {e}")
        # sys.exit()
        ret = False
        result = None

    if ret:
        assert result is not None
        # 明示的に Any にキャストして型チェックの警告を抑制
        result = cast(Any, result)
        # 成功した場合は結果を描画
        CELL = 40
        SIZE = CELL * 8
        # 盤の部分を切り出し
        board = recognizer.extractBoard(image, result.vertex, (SIZE, SIZE))

        # 結果を配列に格納する。-2:不明、-1:空き、0:黒、1:白
        bd = np.ones((8, 8), dtype=np.int8) * -1
        bd[result.isUnknown == True] = -2
        for d in result.disc:
            # 配列を更新しつつ、石の場所に円を描画
            if d.color == br.DiscColor.BLACK:
                color = (0, 0, 0)
                line = (255, 255, 255)
            else:
                color = (255, 255, 255)
                line = (0, 0, 0)
            bd[d.cell[0], d.cell[1]] = int(d.color)
            x = int(d.position[1] * SIZE)
            y = int(d.position[0] * SIZE)
            cv2.circle(board, (x, y), 8, line, -1)
            cv2.circle(board, (x, y), 7, color, -1)

        # 空きマス・不明マスの描画
        for j in range(0, 8):
            for i in range(0, 8):
                x = int((i + 0.5) * CELL)
                y = int((j + 0.5) * CELL)
                if bd[j, i] == -1:
                    # 空きマス
                    cv2.rectangle(board, (x - 4, y - 4), (x + 4, y + 4), (0, 255, 0), -1)
                elif bd[j, i] == -2:
                    # 不明マス
                    cv2.rectangle(board, (x - 4, y - 4), (x + 4, y + 4), (128, 128, 128), -1)
        # 出力
        output_str = ""
        if dialog.format == "sgf":
            # SGF形式での出力
            # AB: Black, AW: White
            sgf_black = ""
            sgf_white = ""
            for d in result.disc:
                # d.cell は (y, x) の順
                col = chr(ord('a') + d.cell[1])
                row = chr(ord('a') + d.cell[0])
                pos = f"[{col}{row}]"
                if d.color == br.DiscColor.BLACK:
                    sgf_black += pos
                else:
                    sgf_white += pos
            
            output_str = "(GM[2];SZ[8]"
            if sgf_black:
                output_str += ";AB" + sgf_black
            if sgf_white:
                output_str += ";AW" + sgf_white
            if dialog.turn == "black":
                output_str += ";PL[B]"
            elif dialog.turn == "white":
                output_str += ";PL[W]"
            else:
                # 手番を PL[B] or PL[W] で追加. 全石数が偶数なら黒番とする
                if len(result.disc) % 2 == 0:
                    output_str += ";PL[B]"
                else:
                    output_str += ";PL[W]"
            
            output_str += ")"
        elif dialog.format == "ggf":
            # GGF形式での出力
            # output_str += "(;GM[Othello]PC[]PB[b_player]PW[w_player]BO[8 "
            output_str += "(;GM[Othello]BO[8 "
            # bd[y, x] -1:空き, 0:黒, 1:白
            for j in range(8):
                for i in range(8):
                    if bd[j, i] == 0:
                        output_str += "*"
                    elif bd[j, i] == 1:
                        output_str += "O"
                    else:
                        output_str += "-"
            
            # 手番
            if dialog.turn == "black":
                output_str += " *]"
            elif dialog.turn == "white":
                output_str += " O]"
            else:
                if len(result.disc) % 2 == 0:
                    output_str += " *]"
                else:
                    output_str += " O]"
            output_str += ";)"
        else:
            # テキスト形式での出力
            # bd[y, x] -1:空き, 0:黒, 1:白
            for j in range(8):
                for i in range(8):
                    if bd[j, i] == 0:
                        output_str += "X"
                    elif bd[j, i] == 1:
                        output_str += "O"
                    else:
                        output_str += "-"
            
            output_str += " "
            # 手番
            if dialog.turn == "black":
                output_str += "X"
            elif dialog.turn == "white":
                output_str += "O"
            else:
                if len(result.disc) % 2 == 0:
                    output_str += "X"
                else:
                    output_str += "O"

    else:
        output_str = "Failed to analyze board"

    print(output_str)           # consoleに出力
    pyperclip.copy(output_str)  # clipboardにコピー

    # LibreOffice Calc にフォーカスがある場合は Ctrl+V で貼り付け
    try:
        if is_libreoffice_calc_focused():
            send_ctrl_v()
            # 小さな遅延の後に Down キーを送信して次の行へ移動
            time.sleep(0.05)
            send_down_key()
    except Exception as e:
        print(f"Auto-paste failed: {e}")

    # Egaroucid が開いている場合はアクティブにして Ctrl+V を送信
    try:
        egaroucid_hwnd = find_egaroucid_window()
        if egaroucid_hwnd:
            activate_window(egaroucid_hwnd)
            time.sleep(0.1)  # アクティブ化を待つ
            send_ctrl_v()
    except Exception as e:
        print(f"Egaroucid auto-paste failed: {e}")

    return dialog, last_rect

last_source = "capture"
last_format = "text"
last_turn = "auto"
last_rect = None
while True:
    dialog, last_rect = show_dialog(last_source, last_format, last_turn, last_rect)
    # 設定を保存
    last_source = dialog.source
    last_format = dialog.format
    last_turn = dialog.turn
    time.sleep(0.5)
