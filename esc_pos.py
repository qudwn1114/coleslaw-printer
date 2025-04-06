import os
import platform
import threading
import socket
import serial
import sys
import winreg
import webbrowser

from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QGraphicsDropShadowEffect
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QColor

app = Flask(__name__)
CORS(app, resources={r"*": {"origins": "*"}})
PORT = 5050

class SplashScreen(QWidget):
    def __init__(self, message="Coleslaw Printer 서버를 준비 중입니다...", timeout=3000):
        super().__init__()
        self.setWindowTitle("로딩 중")
        self.setFixedSize(400, 150)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet("background-color: white; border-radius: 15px;")

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 150))
        shadow.setOffset(0, 5)
        self.setGraphicsEffect(shadow)

        self.label = QLabel(message, self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.label.setGeometry(0, 0, 400, 150)

        screen = QApplication.primaryScreen().availableGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

        QTimer.singleShot(timeout, self.close)

def show_splash_message(message, timeout=3000):
    app = QApplication(sys.argv)
    splash = SplashScreen(message=message, timeout=timeout)
    splash.show()
    QTimer.singleShot(timeout, app.quit)
    app.exec()

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def show_notification(title, message):
    current_os = platform.system()
    if current_os == "Darwin":
        os.system(f"osascript -e 'display notification \"{message}\" with title \"{title}\"'")
    elif current_os == "Windows":
        try:
            from plyer import notification
            notification.notify(title=title, message=message, app_name='ESC/POS Printer', timeout=5)
        except:
            print("plyer 알림 실패")
    else:
        print("지원하지 않는 OS")

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def add_to_startup(app_name="PrintServer", exe_path=None):
    if exe_path is None:
        exe_path = sys.executable
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(registry_key, app_name, 0, winreg.REG_SZ, exe_path)
        winreg.CloseKey(registry_key)
        return True
    except Exception as e:
        print(f"등록 실패: {e}")
        return False

def remove_from_startup(app_name="PrintServer"):
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(registry_key, app_name)
        winreg.CloseKey(registry_key)
        return True
    except Exception as e:
        print(f"제거 실패: {e}")
        return False

def is_in_startup(app_name="PrintServer"):
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_READ)
        value, _ = winreg.QueryValueEx(registry_key, app_name)
        winreg.CloseKey(registry_key)
        return value == sys.executable
    except FileNotFoundError:
        return False

def cleanup_old_startup_entry(app_name="PrintServer"):
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    current_path = sys.executable
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_READ)
        registered_path, _ = winreg.QueryValueEx(registry_key, app_name)
        winreg.CloseKey(registry_key)
        if registered_path != current_path:
            print(f"[\u26a0\ufe0f] 등록된 경로가 현재 실행 경로와 다르네요. 기존 경로 제거: {registered_path}")
            remove_from_startup(app_name)
    except FileNotFoundError:
        pass

@app.route('/')
def index():
    return render_template('test_print.html')

@app.route('/print', methods=['POST'])
def print_receipt():
    if request.is_json:
        data = request.json
    else:
        data = request.form.to_dict()

    port = data.get("port", "COM2")
    message = data.get("message", "")
    baud_rate = data.get("baud_rate", 9600)

    try:
        if not message:
            return jsonify({"status": "fail", "message": "no message..."}), 400

        printer = serial.Serial(port, baud_rate, timeout=1)

        ESC = b'\x1B'
        INIT = ESC + b'@'
        CUT = b'\x1D\x56\x00'

        printer.write(INIT + message.encode('cp949') + CUT)
        printer.close()
        return jsonify({"status": "success", "message": "Printed successfully."}), 200
    except Exception as e:
        print(e)
        return jsonify({"status": "error", "message": str(e)}), 500

def create_tray():
    import pystray
    from pystray import MenuItem as item
    from PIL import Image

    def on_exit(icon, item):
        icon.stop()
        os._exit(0)

    def toggle_startup(icon, item):
        if is_in_startup():
            remove_from_startup()
        else:
            add_to_startup()
        icon.update_menu()

    def open_web(icon, item):
        webbrowser.open(f"http://127.0.0.1:{PORT}/")

    image = Image.open(resource_path("printer.ico"))

    tray_menu = pystray.Menu(
        item('테스트 페이지 열기', open_web),
        item('시작 프로그램 등록', toggle_startup, checked=lambda _: is_in_startup()),
        item('종료', on_exit)
    )

    tray_icon = pystray.Icon("print_server", image, "Coleslaw Printer", tray_menu)
    tray_icon.run()

if __name__ == '__main__':

    if is_port_in_use(PORT):
        # 이미 실행 중이면 안내 후 종료
        show_splash_message("Coleslaw Printer는 이미 실행 중입니다.", timeout=3000)
        sys.exit(0)
    else:
        # 서버 실행 준비 완료 메시지 후 실행
        show_splash_message("Coleslaw Printer 서버를 준비 중입니다...", timeout=3000)

        if platform.system() == "Windows":
            cleanup_old_startup_entry()
            threading.Thread(target=create_tray, daemon=True).start()
        else:
            print("macOS는 트레이 미지원")

        app.run(host='127.0.0.1', port=PORT)