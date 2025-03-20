from flask import Flask, request, jsonify
from flask_cors import CORS
from plyer import notification
import os
import platform
import serial
import socket

app = Flask(__name__)

# CORS 전역 설정 (모든 엔드포인트에서 허용)
CORS(app, supports_credentials=True)

def show_notification(title, message):
    current_os = platform.system()
    if current_os == "Windows":
        try:
            from plyer import notification
            notification.notify(title=title, message=message, app_name='ESC/POS Printer', timeout=5)
        except ImportError:
            print("plyer 모듈이 설치되지 않았습니다. Windows에서 알림이 작동하지 않을 수 있습니다.")
    elif current_os == "Darwin":  # macOS
        os.system(f"osascript -e 'display notification \"{message}\" with title \"{title}\"'")
    else:
        print(f"알림 기능이 {current_os}에서 지원되지 않습니다.")

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

@app.route('/print', methods=['POST'])
def print_receipt():
    if request.is_json:
        data = request.json
    else:
        data = request.form.to_dict()  # 폼 데이터를 딕셔너리로 변환

    port = data.get("port", "COM2")
    message = data.get("message", "")
    baud_rate = data.get("baud_rate", 9600)
    # 프린터 연결 및 출력
    try:
        if not message:
            return jsonify({"status": "fail", "message": "no message..."}), 400

        printer = serial.Serial(port, baud_rate, timeout=1)

        ESC = b'\x1B'   # ESC
        INIT = ESC + b'@'  # 프린터 초기화 명령
        CUT = b'\x1D\x56\x00'  # 용지 자르기 명령

        # ESC/POS 명령어로 출력
        printer.write(INIT + message.encode('cp949') + CUT)
        printer.close()
        return jsonify({"status": "success", "message": "Printed successfully."}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    
# 모든 응답에 CORS 헤더 추가 (Preflight 문제 해결)
@app.after_request
def after_request(response):
    response.headers.add("Access-Control-Allow-Origin", "*")
    response.headers.add("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
    response.headers.add("Access-Control-Allow-Headers", "Content-Type, Authorization")
    return response


if __name__ == '__main__':
    PORT=5050
    if is_port_in_use(PORT):
        show_notification("포트 충돌", f"포트 {PORT}가 이미 사용 중 입니다.!")
    else:
        show_notification("실행 알림", "프린트 서버 실행 완료!")
        app.run(host='127.0.0.1', port=PORT)