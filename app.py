from flask import Flask, request, jsonify
import serial

app = Flask(__name__)

@app.route('/print', methods=['POST'])
def print_receipt():
    data = request.json
    port = data.get("port", "COM3")
    message = data.get("message", "")
    baud_rate = data.get("baud_rate", 9600)

    # 프린터 연결 및 출력
    try:
        printer = serial.Serial(port, baud_rate, timeout=1)

        ESC = b'\x1B'   # ESC
        INIT = ESC + b'@'  # 프린터 초기화 명령
        CUT = b'\x1D\x56\x00'  # 용지 자르기 명령

        # ESC/POS 명령어로 출력
        printer.write(INIT + message.encode('utf-8') + CUT)
        printer.close()
        return jsonify({"status": "success", "message": "Printed successfully."}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5050)