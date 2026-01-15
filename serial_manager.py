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
                    #print(f"[DEBUG] read(): {repr(data)}")  # ← ここ追加
                    return data
                except Exception as e:
                    print(f"[ERROR] read() failed: {e}")
                    return ""
            return ""