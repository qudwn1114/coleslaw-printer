import os
import platform
import threading
import socket
import serial
import sys
import winreg
import webbrowser
import requests, math, time, base64

from queue import Queue
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QGraphicsDropShadowEffect
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QColor

app = Flask(__name__)
CORS(app, resources={r"*": {"origins": "*"}})
PORT = 5050

print_queue = Queue()

def printer_worker():
    while True:
        job = print_queue.get()
        try:
            job()
        except Exception as e:
            print("Print job error:", e)
        finally:
            print_queue.task_done()

threading.Thread(target=printer_worker, daemon=True).start()


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

ESC = b'\x1B'
INIT = ESC + b'@'
ALIGN_CENTER = b'\x1B\x61\x01'  # ESC a 1
ALIGN_LEFT   = b'\x1B\x61\x00'  # ESC a 0 (원래대로)
BLANK_SPACE = b'\x1B\x64\x06' # 6줄 아래
CUT = b'\x1D\x56\x00'
BARCODE_HEIGHT = b'\x1D\x68\x64' # height 100
BARCODE_WIDTH = b'\x1D\x77\x02' # width 2
HRI_POSITION = b'\x1D\x48\x02' #숫자 아래 출력
BARCODE_CODE128 = b'\x1D\x6B\x49' # CODE128

QR_MODEL_2   = b'\x1D\x28\x6B\x04\x00\x31\x41\x32\x00'  # Model 2
QR_SIZE_4 = b'\x1D\x28\x6B\x03\x00\x31\x43\x04'
QR_SIZE_6    = b'\x1D\x28\x6B\x03\x00\x31\x43\x06'
QR_SIZE_8 = b'\x1D\x28\x6B\x03\x00\x31\x43\x08'
QR_ERROR_M  = b'\x1D\x28\x6B\x03\x00\x31\x45\x31'      # error correction M
QR_PRINT    = b'\x1D\x28\x6B\x03\x00\x31\x51\x30'      # print

@app.route('/print', methods=['POST'])
def print_receipt():
    if request.is_json:
        data = request.json
    else:
        data = request.form.to_dict()
    connection_type = data.get("connection_type", "serial") #serial, network
    locale = data.get("locale", "ko_KR")
    message = data.get("message")
    barcode = data.get("barcode")
    qrcode = data.get("qrcode")
    if not message:
        return jsonify({"status": "error","message": "message가 없습니다."}), 400
    if locale == 'ko_KR':
        encoding = 'cp949'
    elif locale == 'ja_JP':
        encoding = 'shift_jis'
    else:
        return jsonify({"status": "error", "message": '지원하지 않는 locale 입니다.'}), 400
    print_bytes = INIT + message.encode(encoding, errors='replace')  
    if barcode:
        barcode_bytes = barcode.encode('ascii')
        print_bytes += (
            b'\n\n' +
            ALIGN_CENTER +
            BARCODE_HEIGHT +
            BARCODE_WIDTH +
            HRI_POSITION +
            BARCODE_CODE128 +
            bytes([len(barcode_bytes)]) +
            barcode_bytes +
            ALIGN_LEFT 
        )
    if qrcode:
        qr_bytes = qrcode.encode('utf-8')
        print_bytes += (
            b'\n\n' +
            ALIGN_CENTER +
            QR_MODEL_2 +
            QR_SIZE_8 +
            QR_ERROR_M +
            qr_store(qr_bytes) +
            QR_PRINT +
            ALIGN_LEFT
        )
    print_bytes += BLANK_SPACE + CUT
    
    if connection_type == 'serial':
        port = data.get("port")
        baud_rate = data.get("baud_rate")
        if not port or not baud_rate:
            return jsonify({"status": "error", "message": "Port 또는 Baud Rate가 없습니다."}), 400
        baud_rate = int(baud_rate)
        def job():
            printer = serial.Serial(port, baud_rate, timeout=1)
            printer.write(print_bytes)
            printer.close()

        print_queue.put(job)
        return jsonify({"status": "queued", "message": "Print job queued."}), 200
    elif connection_type == 'network':
        ip = data.get("ip")
        port = data.get("port")
        if not ip or not port:
            return jsonify({"status": "error", "message": "IP 또는 Port가 없습니다."}), 400
        port = int(port)
        def job():
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(3)
                s.connect((ip, port))
                s.sendall(print_bytes)

        print_queue.put(job)
        return jsonify({"status": "queued", "message": "Print job queued."}), 200
    else:
        return jsonify({"status": "error", "message": '지원하지 않는 connection_type 입니다.'}), 400


def qr_store(data: bytes) -> bytes:
    length = len(data) + 3
    pL = length & 0xFF
    pH = (length >> 8) & 0xFF

    return (
        b'\x1D\x28\x6B' +
        bytes([pL, pH]) +
        b'\x31\x50\x30' +
        data
    )


def create_tray():
    import pystray
    from pystray import MenuItem as item

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

PRINTER_WIDTH = 512
def create_test_image():
    width = PRINTER_WIDTH
    height = 300

    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    # 상단/하단 가이드
    draw.rectangle((0, 0, width-1, height-1), outline="black")
    
    # 중앙 가이드
    draw.line((width//2, 0, width//2, height), fill="black")

    # 테스트 박스
    draw.rectangle((10, 10, width-10, 80), outline="black")
    draw.rectangle((10, 100, width-10, 170), outline="black")

    font_path = "ARIAL.TTF"
    font_size = 20
    font = ImageFont.truetype(font_path, font_size)
    draw.text((20, 20), f"TEST IMAGE {width}px", fill="black", font=font)
    draw.text((20, 110), "WIDTH CHECK", fill="black", font=font)

    img.save("test_image.png")
    return img


def image_to_esc_bytes(img):
    THRESHOLD = 200

    # 그레이스케일 → 1비트
    img = img.convert("L")
    img = img.point(lambda x: 0 if x < THRESHOLD else 255, '1')

    # 사이즈 맞춤
    if img.width != PRINTER_WIDTH:
        img = img.resize(
            (PRINTER_WIDTH, int(img.height * (PRINTER_WIDTH / img.width))), Image.NEAREST
        )

    width, height = img.size
    bytes_per_row = math.ceil(width / 8)

    data = bytearray()
    data += b'\x1B\x40'  # 초기화
    data += b'\x1D\x76\x30\x00'
    data += bytes([
        bytes_per_row & 0xFF,
        (bytes_per_row >> 8) & 0xFF,
        height & 0xFF,
        (height >> 8) & 0xFF
    ])

    pixels = img.load()
    for y in range(height):
        row = bytearray()
        for x in range(0, width, 8):
            byte = 0
            for bit in range(8):
                if x + bit < width and pixels[x + bit, y] == 0:
                    byte |= 1 << (7 - bit)
            row.append(byte)
        data += row

    data += b'\x1B\x64\x06' # 6줄 아래
    data += b'\x1D\x56\x00'  # 커터
    return data

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