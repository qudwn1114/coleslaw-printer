import os
import platform
import threading
import socket
import serial
import sys
import winreg
import webbrowser
import requests, math, time, base64
import uuid
import sqlite3

from serial.tools import list_ports
from queue import Queue
from datetime import datetime, timedelta

from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

from pathlib import Path
import ctypes

app = Flask(__name__)
CORS(app, resources={r"*": {"origins": "*"}})
PORT = 5050

MAX_RETRY = 5

print_queue = Queue()

APP_NAME = "ColeslawPrinter"
APP_DIR = Path(os.getenv("LOCALAPPDATA")) / "ColeslawPrinter"
APP_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = APP_DIR / "print_jobs.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS print_jobs (
            job_id TEXT PRIMARY KEY,
            created_at TEXT,
            printed_at TEXT,
            connection_type TEXT,
            serial_port TEXT,
            baud_rate INTEGER,
            network_ip TEXT,
            network_port INTEGER,
            locale TEXT,
            full_message TEXT,
            barcode TEXT,
            qrcode TEXT,
            status TEXT,
            error_message TEXT,
            retry_count INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()

def insert_job(job):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO print_jobs (
            job_id, created_at, printed_at, connection_type,
            serial_port, baud_rate, network_ip, network_port,
            locale, full_message, barcode, qrcode,
            status, error_message, retry_count
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        job["job_id"], job["created_at"], job.get("printed_at"),
        job["connection_type"], job.get("serial_port"), job.get("baud_rate"),
        job.get("network_ip"), job.get("network_port"),
        job["locale"], job["full_message"], job.get("barcode"), job.get("qrcode"),
        job["status"], job.get("error_message"), job.get("retry_count", 0)
    ))
    conn.commit()
    conn.close()

def get_job(job_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM print_jobs WHERE job_id=?", (job_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return None
    columns = [col[0] for col in cursor.description]
    return dict(zip(columns, row))

def update_job(job_id, status, error_message=None, printed_at=None):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE print_jobs
        SET status = ?, error_message = ?, printed_at = ?
        WHERE job_id = ?
    """, (status, error_message, printed_at, job_id))
    conn.commit()
    conn.close()

def cleanup_old_jobs(days=7):
    cutoff = datetime.now() - timedelta(days=days)
    cutoff_iso = cutoff.isoformat()

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        DELETE FROM print_jobs
        WHERE created_at < ?
    """, (cutoff_iso,))
    conn.commit()
    conn.close()

def list_jobs_by_date(target_date):
    # target_date: "YYYY-MM-DD"
    start = datetime.fromisoformat(target_date + "T00:00:00")
    end = start + timedelta(days=1)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT job_id, created_at, printed_at, connection_type, status, retry_count
        FROM print_jobs
        WHERE created_at >= ? AND created_at < ?
        ORDER BY created_at DESC
    """, (start.isoformat(), end.isoformat()))
    rows = cursor.fetchall()
    conn.close()
    return rows

def increment_retry_count(job_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE print_jobs
        SET retry_count = retry_count + 1
        WHERE job_id = ?
    """, (job_id,))
    conn.commit()
    conn.close()

def worker():
    while True:
        job = print_queue.get()
        print_job(job)
        print_queue.task_done()

threading.Thread(target=worker, daemon=True).start()

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
            notification.notify(title=title, message=message, app_name=APP_NAME, timeout=5)
        except:
            print("plyer 알림 실패")
    else:
        print("지원하지 않는 OS")

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def add_to_startup(exe_path=None):
    if exe_path is None:
        exe_path = sys.executable
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(registry_key, APP_NAME, 0, winreg.REG_SZ, exe_path)
        winreg.CloseKey(registry_key)
        return True
    except Exception as e:
        print(f"등록 실패: {e}")
        return False

def remove_from_startup():
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(registry_key, APP_NAME)
        winreg.CloseKey(registry_key)
        return True
    except Exception as e:
        print(f"제거 실패: {e}")
        return False

def is_in_startup():
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_READ)
        value, _ = winreg.QueryValueEx(registry_key, APP_NAME)
        winreg.CloseKey(registry_key)
        return value == sys.executable
    except FileNotFoundError:
        return False

def cleanup_old_startup_entry():
    key = winreg.HKEY_CURRENT_USER
    reg_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    current_path = sys.executable
    try:
        registry_key = winreg.OpenKey(key, reg_path, 0, winreg.KEY_READ)
        registered_path, _ = winreg.QueryValueEx(registry_key, APP_NAME)
        winreg.CloseKey(registry_key)
        if registered_path != current_path:
            print(f"[\u26a0\ufe0f] 등록된 경로가 현재 실행 경로와 다르네요. 기존 경로 제거: {registered_path}")
            remove_from_startup(APP_NAME)
    except FileNotFoundError:
        pass

@app.route('/')
def index():
    return render_template('test_print.html')

@app.route('/log')
def log_page():
    return render_template('jobs.html')

@app.route('/jobs/<job_id>')
def job_detail_page(job_id):
    return render_template("job_detail.html", job_id=job_id)

@app.route('/jobs', methods=['GET'])
def list_jobs():
    date_str = request.args.get("date")
    if not date_str:
        date_str = datetime.now().strftime("%Y-%m-%d")

    rows = list_jobs_by_date(date_str)

    jobs = []
    for row in rows:
        jobs.append({
            "job_id": row[0],
            "created_at": row[1],
            "printed_at": row[2],
            "connection_type": row[3],
            "status": row[4],
            "retry_count": row[5]
        })
    return jsonify(jobs)

@app.route('/api/jobs/<job_id>')
def api_job_detail(job_id):
    job = get_job(job_id)
    if not job:
        return jsonify({"error": "not found"}), 404
    return jsonify(job)

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


def build_print_bytes(job):
    if job["locale"] == 'ko_KR':
        encoding = 'cp949'
    elif job["locale"] == 'ja_JP':
        encoding = 'shift_jis'
    else:
        raise ValueError("지원하지 않는 locale")

    b = INIT + job["full_message"].encode(encoding, errors='replace')

    if job["barcode"]:
        barcode_bytes = job["barcode"].encode('ascii')
        b += (
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

    if job["qrcode"]:
        qr_bytes = job["qrcode"].encode('utf-8')
        b += (
            b'\n\n' +
            ALIGN_CENTER +
            QR_MODEL_2 +
            QR_SIZE_8 +
            QR_ERROR_M +
            qr_store(qr_bytes) +
            QR_PRINT +
            ALIGN_LEFT
        )

    b += BLANK_SPACE + CUT
    return b

def print_job(job):
    try:
        print_bytes = build_print_bytes(job)

        if job["connection_type"] == "serial":
            available = [p.device for p in list_ports.comports()]
            if job["serial_port"] not in available:
                raise Exception("Serial port not found")
            with serial.Serial(job["serial_port"], job["baud_rate"], timeout=1) as printer:
                if not printer.is_open:
                    raise Exception("Serial port not open")
                printer.write(print_bytes)
                printer.flush()

        else:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(3)
                s.connect((job["network_ip"], job["network_port"]))
                s.sendall(print_bytes)

        update_job(job["job_id"], status="printed", printed_at=datetime.now().isoformat())

    except Exception as e:
        update_job(job["job_id"], status="failed", error_message=str(e))



@app.route('/print', methods=['POST'])
def print_receipt():
    data = request.get_json() if request.is_json else request.form.to_dict()

    connection_type = data.get("connection_type", "serial")
    locale = data.get("locale", "ko_KR")
    message = data.get("message")
    barcode = data.get("barcode")
    qrcode = data.get("qrcode")

    if connection_type == "serial":
        serial_port = data.get("serial_port")
        baud_rate = data.get("baud_rate")
        if not serial_port or not baud_rate:
            return jsonify({"status":"error","message":"Port 또는 Baud Rate가 없습니다."}), 400
        baud_rate = int(baud_rate)
        network_ip = None
        network_port = None
    else:
        network_ip = data.get("network_ip")
        network_port = data.get("network_port", 0)
        if not network_ip or not network_port:
            return jsonify({"status":"error","message":"IP 또는 Port가 없습니다."}), 400
        network_port = int(network_port)
        serial_port = None
        baud_rate = None

    job_id = str(uuid.uuid4())
    job = {
        "job_id": job_id,
        "created_at": datetime.now().isoformat(),
        "printed_at": None,
        "connection_type": connection_type,
        "serial_port": serial_port,
        "baud_rate": baud_rate,
        "network_ip": network_ip,
        "network_port": network_port,
        "locale": locale,
        "full_message": message,
        "barcode": barcode,
        "qrcode": qrcode,
        "status": "queued",
        "error_message": None,
        "retry_count": 0
    }

    insert_job(job)
    print_queue.put(job)

    return jsonify({"status":"queued", "message":"Print job queued.", "job_id": job_id}), 200

@app.route('/reprint', methods=['POST'])
def reprint_job():
    data = request.get_json() if request.is_json else request.form.to_dict()
    old_job_id = data.get("job_id")

    old = get_job(old_job_id)
    if not old:
        return jsonify({"status":"error","message":"job not found"}), 404
    
    if old.get("retry_count", 0) >= MAX_RETRY:
        return jsonify({"status":"error", "message":"재출력 최대 횟수 초과"}), 400

    new_job = dict(old)
    new_job["job_id"] = str(uuid.uuid4())
    new_job["created_at"] = datetime.now().isoformat()
    new_job["printed_at"] = None
    new_job["status"] = "queued"
    new_job["error_message"] = None
    new_job["retry_count"] = 0

    insert_job(new_job)
    increment_retry_count(old_job_id)
    print_queue.put(new_job)

    return jsonify({"status":"queued", "message":"Print job queued.", "job_id": new_job["job_id"]}), 200


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

    def open_log(icon, item):
        webbrowser.open(f"http://127.0.0.1:{PORT}/log")

    image = Image.open(resource_path("printer.ico"))

    tray_menu = pystray.Menu(
        item('Test Page', open_web),
        item('Log', open_log),
        item('Run at Startup ', toggle_startup, checked=lambda _: is_in_startup()),
        item('Exit', on_exit)
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
    init_db()
    cleanup_old_jobs(7)
    if is_port_in_use(PORT):
        ctypes.windll.user32.MessageBoxW(
            0, 
            "Coleslaw Printer Server is already running.", 
            "Coleslaw Printer", 
            0x30
        )
        sys.exit(0)
    else:
        print("Starting Coleslaw Printer Server...")

        if platform.system() == "Windows":
            cleanup_old_startup_entry()
            threading.Thread(target=create_tray, daemon=True).start()
        else:
            print("macOS는 트레이 미지원")

        app.run(host='127.0.0.1', port=PORT)