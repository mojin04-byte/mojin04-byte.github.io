import sys
import tkinter as tk
from tkinter import ttk, messagebox
from ctypes import windll

# 모듈 폴더에서 각 클래스 불러오기
# (modules 폴더 안에 해당 파일들이 있어야 합니다)
from modules.ip_scanner import IPScannerGUI
from modules.modbus_master import PLCLoggerGUI
from modules.modbus_slave import ModbusSlaveGUI
from modules.serial_terminal import SerialTerminalGUI
from modules.address_calc import AddressCalculatorGUI
from modules.report_v1 import CimonReportMaker
from modules.report_v2 import CimonReportMakerV2

class EngineeringToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엔지니어링 툴 ver3.4 (분할 버전)")
        self.root.geometry("1100x850") 
        
        # 상단 정보 바
        top_bar = tk.Frame(root, bg="#eee", height=30)
        top_bar.pack(fill="x", side="top")
        tk.Label(top_bar, text="Engineering Tool v3.4", bg="#eee", fg="#555").pack(side="left", padx=10)
        tk.Button(top_bar, text="정보 (About)", command=self.show_about, bg="white", relief="flat").pack(side="right", padx=5, pady=2)
        
        # 탭 컨트롤 생성
        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 각 탭에 기능 연결
        t1 = ttk.Frame(nb); nb.add(t1, text="  IP 스캐너  "); self.scanner = IPScannerGUI(t1)
        t2 = ttk.Frame(nb); nb.add(t2, text="  Modbus 마스터  "); self.logger = PLCLoggerGUI(t2)
        t3 = ttk.Frame(nb); nb.add(t3, text="  Modbus 슬레이브  "); self.slave = ModbusSlaveGUI(t3)
        t4 = ttk.Frame(nb); nb.add(t4, text="  시리얼 터미널  "); self.serial = SerialTerminalGUI(t4)
        t5 = ttk.Frame(nb); nb.add(t5, text="  주소 계산기  "); self.calc = AddressCalculatorGUI(t5)
        t6 = ttk.Frame(nb); nb.add(t6, text="  리포트 생성(v1)  "); self.cimon = CimonReportMaker(t6)
        t7 = ttk.Frame(nb); nb.add(t7, text="  리포트 생성(v2)  "); self.cimon2 = CimonReportMakerV2(t7)
        
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def show_about(self):
        msg = (
            "엔지니어링 툴 ver 3.4 (모듈 분리형)\n\n"
            "제작자: 서용승\n"
            "기능: IP스캐너, Modbus M/S, 시리얼, 주소변환, 리포트생성기"
        )
        messagebox.showinfo("프로그램 정보", msg)

    def on_close(self):
        # 종료 시 각 모듈의 스레드/포트 정리
        try: self.scanner.is_scanning = False
        except: pass
        try: self.logger.is_running = False
        except: pass
        try: 
            if self.slave.server: 
                self.slave.server.shutdown()
                self.slave.server.server_close()
        except: pass
        try: 
            if self.serial.ser: self.serial.is_open = False; self.serial.ser.close()
        except: pass
        self.root.destroy(); sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    try:
        # 고해상도 모니터(DPI) 대응
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    app = EngineeringToolApp(root)
    root.mainloop()