import numpy as np
import cv2
import board_recognition as br
import pyperclip
import time

import sys
from PySide6.QtWidgets import (QApplication, QFileDialog, QDialog, QVBoxLayout, 
                             QPushButton, QLabel, QWidget)
from PySide6.QtCore import Qt, QRect, QPoint
from PySide6.QtGui import QPainter, QColor, QScreen, QPixmap

from PySide6.QtWidgets import (QApplication, QFileDialog, QDialog, QVBoxLayout, 
                             QPushButton, QLabel, QWidget, QRadioButton, QGroupBox, QHBoxLayout)
from PySide6.QtCore import Qt, QRect, QPoint
from PySide6.QtGui import QPainter, QColor, QScreen, QPixmap

class SelectionDialog(QDialog):
    def __init__(self, initial_source="capture", initial_format="text", last_rect=None):
        super().__init__()
        self.setWindowTitle("Reversi to Clipboard")
        self.source = initial_source
        self.format = initial_format
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
        
        # Buttons
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept_settings)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)

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
        self.accept()

class CaptureOverlay(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.setWindowOpacity(0.3)
        self.setStyleSheet("background-color: black;")
        self.setCursor(Qt.CrossCursor)
        
        # 全画面表示
        screen = QApplication.primaryScreen()
        self.setGeometry(screen.geometry())
        
        self.start_pos = None
        self.end_pos = None
        self.capture_pixmap = None
        self.is_selecting = False

    def paintEvent(self, event):
        if self.is_selecting and self.start_pos and self.end_pos:
            painter = QPainter(self)
            painter.setPen(QColor(255, 255, 255))
            painter.setBrush(QColor(255, 255, 255, 50))
            rect = QRect(self.start_pos, self.end_pos).normalized()
            painter.drawRect(rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_pos = event.position().toPoint()
            self.is_selecting = True

    def mouseMoveEvent(self, event):
        if self.is_selecting:
            self.end_pos = event.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.is_selecting:
            self.end_pos = event.position().toPoint()
            self.is_selecting = False
            self.capture_area()
            self.close()

    def capture_area(self):
        self.rect = QRect(self.start_pos, self.end_pos).normalized()
        if self.rect.width() < 10 or self.rect.height() < 10:
            return
            
        screen = QApplication.primaryScreen()
        self.hide() # 自分自身を隠してキャプチャ
        QApplication.processEvents() # 再描画を待機
        self.capture_pixmap = screen.grabWindow(0, self.rect.x(), self.rect.y(), self.rect.width(), self.rect.height())

# QApplication インスタンスを作成
app = QApplication(sys.argv)

def show_dialog(default_source="capture", default_format="text", last_rect=None):
    # 設定の選択
    dialog = SelectionDialog(default_source, default_format, last_rect)
    if dialog.exec() != QDialog.Accepted:
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
        qimage = pixmap.toImage().convertToFormat(QImage.Format_RGBA8888)
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
            last_rect = overlay.rect
            # QPixmap を numpy 形式 (OpenCV) に変換
            from PySide6.QtGui import QImage
            qimage = overlay.capture_pixmap.toImage().convertToFormat(QImage.Format_RGBA8888)
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

    if image is None or image.size == 0:
        print("Failed to load/capture image.")
        sys.exit()


    # 解析に当たってのヒント情報
    hint = br.Hint()
    hint.mode = br.Mode.PHOTO # 写真モード

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
            if len(result.disc) % 2 == 0:
                output_str += "X"
            else:
                output_str += "O"

    else:
        output_str = "Failed to analyze board"

    print(output_str)           # consoleに出力
    pyperclip.copy(output_str)  # clipboardにコピー
    return dialog, last_rect

last_source = "capture"
last_format = "text"
last_rect = None
while True:
    dialog, last_rect = show_dialog(last_source, last_format, last_rect)
    # 設定を保存
    last_source = dialog.source
    last_format = dialog.format
    time.sleep(2)
