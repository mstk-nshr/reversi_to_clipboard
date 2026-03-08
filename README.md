# reversi_to_clipboard

A Python program that recognizes the Reversi board state from an image file or screenshot and outputs it to the clipboard in SGF or text format.  

画像ファイルまたはスクリーンショットからリバーシの盤面を認識し、SGF形式またはテキスト形式でクリップボードに出力するpythonプログラム

## Usage

Run the following command to copy the reversi game record to the clipboard:

```bash / cmd / pwsh
python reversi_to_clipboard.py
```

Then, type CTRL+V on the "Egaroucid" reversi application window.  
(*) Egaroucid does not support SGF format. Use text format.

## License

MIT License

## Note

- forked from <https://github.com/lavox/reversi_recognition.git>
- cv2 module : $ pip install opencv-python
- PySide6 module : $ pip install PySide6
- pyperclip module : $ pip install pyperclip
