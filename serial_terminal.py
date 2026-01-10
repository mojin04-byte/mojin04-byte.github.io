import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
from datetime import datetime

# 외부 라이브러리 로딩 (pyserial)
try:
    import serial
    from serial.tools import list_ports
except ImportError:
    serial = None
    print("오류: pyserial 라이브러리가 필요합니다.")

# [여기 아래에 원본의 class SerialTerminalGUI: 전체 코드를 붙여넣으세요]
class SerialTerminalGUI:
    def __init__(self, parent):
        self.parent = parent
        self.ser = None
        self.is_open = False
        self.is_auto_sending = False
        
        settings_frame = ttk.LabelFrame(parent, text="포트 설정", padding="10")
        settings_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(settings_frame, text="COM:").pack(side="left")
        self.com_entry = ttk.Entry(settings_frame, width=8); self.com_entry.insert(0, "COM3"); self.com_entry.pack(side="left", padx=5)
        ttk.Label(settings_frame, text="Baud:").pack(side="left")
        self.baud_combo = ttk.Combobox(settings_frame, values=["9600", "19200", "38400", "115200"], width=8)
        self.baud_combo.current(0); self.baud_combo.pack(side="left", padx=5)
        
        self.btn_open = ttk.Button(settings_frame, text="열기", command=self.toggle_serial)
        self.btn_open.pack(side="left", padx=10)

        send_frame = ttk.LabelFrame(parent, text="전송", padding="10")
        send_frame.pack(fill="x", padx=10, pady=5)
        
        self.input_format = tk.StringVar(value="HEX")
        ttk.Radiobutton(send_frame, text="HEX", variable=self.input_format, value="HEX").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(send_frame, text="ASCII", variable=self.input_format, value="ASCII").grid(row=0, column=1, padx=5)
        
        self.txt_send = ttk.Entry(send_frame, width=50)
        self.txt_send.grid(row=0, column=2, padx=5)
        ttk.Button(send_frame, text="전송", command=self.send_data).grid(row=0, column=3, padx=5)
        
        auto_frame = ttk.Frame(send_frame)
        auto_frame.grid(row=1, column=0, columnspan=4, sticky="w", pady=5)
        self.chk_auto_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(auto_frame, text="자동 전송", variable=self.chk_auto_var, command=self.toggle_auto).pack(side="left", padx=5)
        ttk.Label(auto_frame, text="주기(초):").pack(side="left")
        self.entry_interval = ttk.Entry(auto_frame, width=5); self.entry_interval.insert(0, "1.0"); self.entry_interval.pack(side="left", padx=2)

        log_frame = ttk.LabelFrame(parent, text="로그", padding="5")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        view_frame = ttk.Frame(log_frame)
        view_frame.pack(fill="x")
        ttk.Label(view_frame, text="보기 방식:").pack(side="left")
        self.view_format = tk.StringVar(value="HEX")
        ttk.Radiobutton(view_frame, text="HEX", variable=self.view_format, value="HEX").pack(side="left")
        ttk.Radiobutton(view_frame, text="ASCII", variable=self.view_format, value="ASCII").pack(side="left")
        ttk.Button(view_frame, text="지우기", command=self.clear_log).pack(side="right")

        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=10)
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config("TX", foreground="blue"); self.log_text.tag_config("RX", foreground="green")

    def toggle_serial(self):
        if not self.is_open:
            try:
                self.ser = serial.Serial(self.com_entry.get(), int(self.baud_combo.get()), timeout=0.1)
                self.is_open = True
                self.btn_open.config(text="닫기")
                self.log("포트 열림", "INFO")
                threading.Thread(target=self.read_loop, daemon=True).start()
            except Exception as e: messagebox.showerror("오류", f"열기 실패: {e}")
        else:
            self.is_open = False
            self.chk_auto_var.set(False)
            if self.ser: self.ser.close()
            self.btn_open.config(text="열기")
            self.log("포트 닫힘", "INFO")

    def send_data(self):
        if not self.is_open: return
        raw = self.txt_send.get().strip()
        if not raw: return
        try:
            data = bytes.fromhex(raw.replace(" ", "")) if self.input_format.get() == "HEX" else raw.encode()
            self.ser.write(data)
            self.display(data, "TX")
        except: self.log("형식 오류", "ERR")

    def toggle_auto(self):
        if self.chk_auto_var.get():
            self.is_auto_sending = True
            threading.Thread(target=self.auto_loop, daemon=True).start()
        else: self.is_auto_sending = False

    def auto_loop(self):
        while self.is_auto_sending and self.is_open:
            try: interval = float(self.entry_interval.get())
            except: interval = 1.0
            self.parent.after(0, self.send_data)
            time.sleep(max(0.1, interval))

    def read_loop(self):
        while self.is_open and self.ser.is_open:
            try:
                if self.ser.in_waiting:
                    data = self.ser.read(self.ser.in_waiting)
                    self.display(data, "RX")
                time.sleep(0.05)
            except: break

    def display(self, data, tag):
        text = " ".join([f"{b:02X}" for b in data]) if self.view_format.get() == "HEX" else data.decode(errors='replace')
        self.log(f"[{tag}] {text}", tag)

    def log(self, msg, tag="INFO"):
        t = datetime.now().strftime("[%H:%M:%S] ")
        def _add():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, t + msg + "\n", tag)
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.parent.after(0, _add)

    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
