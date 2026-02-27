import serial
import threading
import time

class SerialManager:
    def __init__(self, baudrate=38400, timeout=1):
        self.baudrate = baudrate
        self.timeout = timeout
        self.ser = None
        self.lock = threading.Lock()  # 排他制御

    def connect(self, port):
        try:
            self.ser = serial.Serial(
                port=port,
                baudrate=self.baudrate,
                timeout=self.timeout,
                write_timeout=self.timeout
            )
            return True
        except Exception as e:
            return (False, str(e))

    def disconnect(self):
        with self.lock:
            if self.ser and self.ser.is_open:
                self.ser.close()

    def is_connected(self):
        return self.ser and self.ser.is_open

    def write(self, data):
        """バイナリで送信（例：b'test\r'）"""
        with self.lock:
            if self.is_connected():
                self.ser.write(data)

    def write_line(self, text, end="\r"):
        """文字列で送信し、終端記号付き（デフォルト: CR）"""
        self.write((text + end).encode("utf-8"))

    def read_line(self):
        """改行までのレスポンスを受信（strで返す）"""
        with self.lock:
            if self.is_connected():
                try:
                    return self.ser.readline().decode("utf-8", errors="ignore").strip()
                except Exception:
                    return ""
            return ""

    def read_all(self):
        """受信バッファをすべて読み取る（str）"""
        with self.lock:
            if self.is_connected():
                try:
                    return self.ser.read_all().decode("utf-8", errors="ignore")
                except Exception:
                    return ""
            return ""

    def flush_input(self):
        """受信バッファをクリア"""
        with self.lock:
            if self.is_connected():
                self.ser.reset_input_buffer()

    def read(self, size=1):
        """指定バイト数のデータを読み取り（デフォルト: 1バイト）"""
        with self.lock:
            if self.is_connected() and self.ser.in_waiting > 0:
                try:
                    data = self.ser.read(size).decode("utf-8", errors="ignore")
                    return data
                except Exception as e:
                    print(f"[ERROR] read() failed: {e}")
                    return ""
            return ""

    def send_command(self, cmd):
        """コマンド送信（終端CR付き、バッファクリア後に送信）"""
        with self.lock:
            if self.is_connected():
                try:
                    self.ser.reset_input_buffer()
                    self.ser.write((cmd + '\r').encode('utf-8'))
                    self.ser.flush()
                    return True
                except Exception as e:
                    print(f"[ERROR] send_command() failed: {e}")
                    return False
            return False

    def send_command_with_response(self, cmd, wait_sec=0.001, read_timeout=0.015, prompt=">"):
        """コマンド送信後、応答を受信して返す（プロンプト検出で終了）"""
        with self.lock:
            if not self.is_connected():
                return None
            try:
                self.ser.reset_input_buffer()
                self.ser.write((cmd + '\r').encode('utf-8'))
                self.ser.flush()
                time.sleep(wait_sec)

                old_timeout = self.ser.timeout
                self.ser.timeout = read_timeout
                response_lines = []
                while True:
                    line = self.ser.readline().decode('utf-8', errors='ignore').strip()
                    if not line:
                        break
                    if prompt and line == prompt:
                        break
                    response_lines.append(line)
                self.ser.timeout = old_timeout

                result = '\n'.join(response_lines) if response_lines else ""
                if prompt:
                    result = result.rstrip(prompt).rstrip('\r')
                result = result.replace('\r', '\n')
                return result
            except Exception as e:
                print(f"[ERROR] send_command_with_response() failed: {e}")
                return None