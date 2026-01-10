import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import time
from datetime import datetime

# --- 외부 라이브러리 로딩 (상세 에러 출력 기능 추가) ---
try:
    # pymodbus 3.x 버전 호환
    from pymodbus.client import ModbusTcpClient, ModbusSerialClient
    MODBUS_AVAILABLE = True
except ImportError as e:
    ModbusTcpClient = None
    ModbusSerialClient = None
    MODBUS_AVAILABLE = False
    print(f"[Modbus Master] 라이브러리 로드 실패 (상세): {e}")

class PLCLoggerGUI:
    def __init__(self, parent):
        self.parent = parent
        self.is_running = False
        self.client = None
        self.thread = None
        
        style = ttk.Style()
        style.configure("TLabel", font=("맑은 고딕", 10))
        style.configure("TButton", font=("맑은 고딕", 10, "bold"))

        # 설정
        settings_frame = ttk.LabelFrame(parent, text="통신 설정", padding="10")
        settings_frame.pack(fill="x", padx=10, pady=5)

        self.mode_var = tk.StringVar(value="TCP")
        ttk.Radiobutton(settings_frame, text="Ethernet (TCP)", variable=self.mode_var, value="TCP", command=self.toggle_mode).grid(row=0, column=0, padx=5)
        ttk.Radiobutton(settings_frame, text="Serial (RTU)", variable=self.mode_var, value="RTU", command=self.toggle_mode).grid(row=0, column=1, padx=5)

        self.tcp_frame = ttk.Frame(settings_frame)
        self.tcp_frame.grid(row=1, column=0, columnspan=4, sticky="w", pady=5)
        ttk.Label(self.tcp_frame, text="IP 주소:").pack(side="left", padx=5)
        self.ip_entry = ttk.Entry(self.tcp_frame, width=15); self.ip_entry.insert(0, "192.168.0.10"); self.ip_entry.pack(side="left")
        ttk.Label(self.tcp_frame, text="Port:").pack(side="left", padx=5)
        self.port_entry = ttk.Entry(self.tcp_frame, width=6); self.port_entry.insert(0, "502"); self.port_entry.pack(side="left")

        self.rtu_frame = ttk.Frame(settings_frame)
        ttk.Label(self.rtu_frame, text="COM:").pack(side="left", padx=5)
        self.com_entry = ttk.Entry(self.rtu_frame, width=8); self.com_entry.insert(0, "COM3"); self.com_entry.pack(side="left")
        ttk.Label(self.rtu_frame, text="Baud:").pack(side="left", padx=5)
        self.baud_combo = ttk.Combobox(self.rtu_frame, values=["9600", "19200", "38400", "115200"], width=8)
        self.baud_combo.current(0); self.baud_combo.pack(side="left")

        # 파라미터
        read_frame = ttk.LabelFrame(parent, text="읽기 옵션", padding="10")
        read_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(read_frame, text="영역 선택:").grid(row=0, column=0, padx=5)
        self.reg_type = ttk.Combobox(read_frame, values=["Holding Reg (4x)", "Input Reg (3x)", "Coils (0x)", "Discrete In (1x)"], width=15, state="readonly")
        self.reg_type.current(0)
        self.reg_type.grid(row=0, column=1, padx=5)

        ttk.Label(read_frame, text="국번(ID):").grid(row=0, column=2, padx=5)
        self.unit_id_entry = ttk.Entry(read_frame, width=5); self.unit_id_entry.insert(0, "1"); self.unit_id_entry.grid(row=0, column=3)
        ttk.Label(read_frame, text="시작 주소:").grid(row=0, column=4, padx=5)
        self.addr_entry = ttk.Entry(read_frame, width=6); self.addr_entry.insert(0, "0"); self.addr_entry.grid(row=0, column=5)
        ttk.Label(read_frame, text="개수:").grid(row=0, column=6, padx=5)
        self.count_entry = ttk.Entry(read_frame, width=5); self.count_entry.insert(0, "10"); self.count_entry.grid(row=0, column=7)
        ttk.Label(read_frame, text="주기(초):").grid(row=0, column=8, padx=5)
        self.interval_entry = ttk.Entry(read_frame, width=5); self.interval_entry.insert(0, "1.0"); self.interval_entry.grid(row=0, column=9)

        # 버튼
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=10, pady=10)
        self.start_btn = tk.Button(btn_frame, text="수집 시작 (Start)", bg="#4CAF50", fg="white", font=("맑은 고딕", 11, "bold"), command=self.start_logging)
        self.start_btn.pack(side="left", fill="x", expand=True, padx=5)
        self.stop_btn = tk.Button(btn_frame, text="정지 (Stop)", bg="#F44336", fg="white", font=("맑은 고딕", 11, "bold"), command=self.stop_logging, state="disabled")
        self.stop_btn.pack(side="left", fill="x", expand=True, padx=5)
        self.save_btn = tk.Button(btn_frame, text="로그 저장 (Save)", bg="#2196F3", fg="white", font=("맑은 고딕", 11, "bold"), command=self.save_log_to_file)
        self.save_btn.pack(side="left", fill="x", expand=True, padx=5)

        # 로그
        log_frame = ttk.LabelFrame(parent, text="실시간 모니터링 로그", padding="5")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=15, font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config("TX", foreground="blue")
        self.log_text.tag_config("RX", foreground="green")
        self.log_text.tag_config("ERR", foreground="red")
        self.log_text.tag_config("INFO", foreground="black")

        self.status_var = tk.StringVar(value="대기 중...")
        ttk.Label(parent, textvariable=self.status_var, relief="sunken", anchor="w").pack(side="bottom", fill="x")
        
        if not MODBUS_AVAILABLE:
            self.log_msg("라이브러리 로드 실패. pip install pymodbus 필요", "ERR")
            self.start_btn.config(state="disabled")

    def toggle_mode(self):
        if self.mode_var.get() == "TCP":
            self.rtu_frame.grid_forget()
            self.tcp_frame.grid(row=1, column=0, columnspan=4, sticky="w", pady=5)
        else:
            self.tcp_frame.grid_forget()
            self.rtu_frame.grid(row=1, column=0, columnspan=4, sticky="w", pady=5)

    def log_msg(self, msg, tag="INFO"):
        t = datetime.now().strftime("[%H:%M:%S] ")
        def _append():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, t + msg + "\n", tag)
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.parent.after(0, _append)

    def save_log_to_file(self):
        filename = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if filename:
            try:
                content = self.log_text.get("1.0", tk.END)
                with open(filename, 'w', encoding='utf-8') as f: f.write(content)
                messagebox.showinfo("저장 완료", "로그가 저장되었습니다.")
            except Exception as e: messagebox.showerror("오류", f"저장 실패: {e}")

    def start_logging(self):
        try:
            config = {
                'mode': self.mode_var.get(),
                'ip': self.ip_entry.get(),
                'port': int(self.port_entry.get()),
                'com': self.com_entry.get(),
                'baud': int(self.baud_combo.get()),
                'uid': int(self.unit_id_entry.get()),
                'addr': int(self.addr_entry.get()),
                'cnt': int(self.count_entry.get()),
                'sec': float(self.interval_entry.get()),
                'type': self.reg_type.get()
            }
        except:
            messagebox.showerror("입력 오류", "설정값을 확인해주세요.")
            return

        self.is_running = True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.status_var.set("상태: 실행 중...")
        
        self.thread = threading.Thread(target=self.scan_loop, args=(config,), daemon=True)
        self.thread.start()

    def stop_logging(self):
        self.is_running = False
        self.log_msg("정지 요청됨.")

    def scan_loop(self, config):
        try:
            if config['mode'] == "TCP":
                self.client = ModbusTcpClient(config['ip'], port=config['port'])
            else:
                # pymodbus 3.x 호환 (method='rtu' 대신 명시적 파라미터 사용 권장되나 호환성 유지 시도)
                self.client = ModbusSerialClient(port=config['com'], baudrate=config['baud'], bytesize=8, parity='N', stopbits=1)
            
            if not self.client.connect():
                self.log_msg(f"연결 실패: {config['mode']}", "ERR")
                self.reset_ui()
                return
            
            self.log_msg(f"연결 성공. ({config['type']} 수집)", "INFO")
            
            while self.is_running:
                rr = None
                t_str = config['type']
                
                # pymodbus 3.x에서는 slave=... 인자가 필수일 수 있음
                if "Holding" in t_str:
                    rr = self.client.read_holding_registers(config['addr'], count=config['cnt'], slave=config['uid'])
                elif "Input Reg" in t_str:
                    rr = self.client.read_input_registers(config['addr'], count=config['cnt'], slave=config['uid'])
                elif "Coils" in t_str:
                    rr = self.client.read_coils(config['addr'], count=config['cnt'], slave=config['uid'])
                elif "Discrete" in t_str:
                    rr = self.client.read_discrete_inputs(config['addr'], count=config['cnt'], slave=config['uid'])
                
                if not rr.isError():
                    val = rr.bits if ("Coils" in t_str or "Discrete" in t_str) else rr.registers
                    self.log_msg(f"주소 {config['addr']} ~ : {val}", "RX")
                else:
                    self.log_msg(f"읽기 실패: {rr}", "ERR")
                
                time.sleep(config['sec'])
                
        except Exception as e:
            self.log_msg(f"시스템 에러: {e}", "ERR")
        finally:
            if self.client: self.client.close()
            self.reset_ui()

    def reset_ui(self):
        self.parent.after(0, lambda: self._reset())
    def _reset(self):
        self.is_running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.status_var.set("상태: 대기 중")