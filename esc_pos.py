import os
import platform
import threading
import socket
import serial
import sys
import winreg

from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

app = Flask(__name__)
CORS(app, resources={r"*": {"origins": "*"}})

# 🔔 OS별 알림 함수
def show_notification(title, message):
    current_os = platform.system()
    if current_os == "Darwin":  # macOS
        os.system(f"osascript -e 'display notification \"{message}\" with title \"{title}\"'")
    elif current_os == "Windows":
        try:
            from plyer import notification
            notification.notify(title=title, message=message, app_name='ESC/POS Printer', timeout=5)
        except:
            print("plyer 알림 실패")
    else:
        print("지원하지 않는 OS")

# 🖠 포트 중복 검사
def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

# ✅ 시작 프로그램 등록/해제 함수 (Windows 전용)
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
        winreg.QueryValueEx(registry_key, app_name)
        winreg.CloseKey(registry_key)
        return True
    except FileNotFoundError:
        return False

# 🌐 라우트
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

# 🪟 트레이 아이콘 (Windows 전용)
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
            item.checked = False
        else:
            add_to_startup()
            item.checked = True
        icon.update_menu()

    image = Image.open("icon.ico")  # ico 파일 위치

    startup_item = item(
        '시작 프로그램 등록', toggle_startup, checked=lambda item: is_in_startup()
    )

    menu = pystray.Menu(
        item('종료', on_exit),
        startup_item
    )

    tray_icon = pystray.Icon("print_server", image, "ESC/POS Print Server", menu)
    tray_icon.run()

# 🏁 메인 실행
if __name__ == '__main__':
    PORT = 5050
    if is_port_in_use(PORT):
        show_notification("포트 충돌", f"포트 {PORT}가 이미 사용 중 입니다.!")
    else:
        show_notification("실행 알림", "프린트 서버 실행 완료!")

        if platform.system() == "Windows":
            threading.Thread(target=create_tray, daemon=True).start()
        else:
            print("macOS는 트레이 미지원")

        app.run(host='127.0.0.1', port=PORT)
