import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading

# --- 외부 라이브러리 로딩 (pymodbus 3.x 및 하위 호환성 강화) ---
MODBUS_SERVER_AVAILABLE = False
try:
    # 1. 공통 데이터 모듈 임포트 시도
    from pymodbus.server import StartTcpServer, StartSerialServer, ServerStop
    from pymodbus.datastore import ModbusSequentialDataBlock
    from pymodbus.datastore.context import ModbusServerContext, ModbusSlaveContext
    MODBUS_SERVER_AVAILABLE = True
except ImportError:
    try:
        # 2. 구버전 또는 경로가 다를 경우 재시도 (pymodbus 구조가 자주 바뀜)
        from pymodbus.server.sync import StartTcpServer, StartSerialServer, ServerStop
        from pymodbus.datastore import ModbusSequentialDataBlock
        from pymodbus.datastore import ModbusServerContext, ModbusSlaveContext
        MODBUS_SERVER_AVAILABLE = True
    except ImportError as e:
        print(f"[Modbus Slave] 라이브러리 로드 실패: {e}")
        # 에러 방지용 더미 클래스/함수 정의
        StartTcpServer = None
        StartSerialServer = None
        ServerStop = None
        ModbusSequentialDataBlock = None
        ModbusServerContext = None
        ModbusSlaveContext = None

class ModbusSlaveGUI:
    def __init__(self, parent):
        self.parent = parent
        self.parent.title("가상 PLC 시뮬레이터 (Modbus Slave)")
        self.parent.geometry("600x750")
        
        self.server_thread = None
        self.is_server_running = False
        self.context = None
        self.auto_sim_active = False

        # --- 설정 UI ---
        settings_frame = ttk.LabelFrame(parent, text="슬레이브 설정", padding="10")
        settings_frame.pack(fill="x", padx=10, pady=5)

        self.mode_var = tk.StringVar(value="TCP")
        ttk.Radiobutton(settings_frame, text="TCP 서버", variable=self.mode_var, value="TCP", command=self.update_ui_state).grid(row=0, column=0, padx=5, sticky="w")
        ttk.Radiobutton(settings_frame, text="시리얼 슬레이브 (RTU)", variable=self.mode_var, value="RTU", command=self.update_ui_state).grid(row=0, column=1, padx=5, sticky="w")

        # 연결 파라미터
        self.conn_frame = ttk.Frame(settings_frame)
        self.conn_frame.grid(row=1, column=0, columnspan=6, sticky="w", pady=5)
        
        self.lbl_port = ttk.Label(self.conn_frame, text="포트:")
        self.ent_port = ttk.Entry(self.conn_frame, width=6); self.ent_port.insert(0, "502")
        
        self.lbl_com = ttk.Label(self.conn_frame, text="COM:")
        self.ent_com = ttk.Entry(self.conn_frame, width=8); self.ent_com.insert(0, "COM1")
        self.lbl_baud = ttk.Label(self.conn_frame, text="Baud:")
        self.cmb_baud = ttk.Combobox(self.conn_frame, values=["9600","19200","38400","115200"], width=8)
        self.cmb_baud.current(0)

        # 메모리 파라미터
        self.param_frame = ttk.Frame(settings_frame)
        self.param_frame.grid(row=2, column=0, columnspan=6, sticky="w", pady=5)
        
        ttk.Label(self.param_frame, text="ID:").pack(side="left", padx=5)
        self.ent_uid = ttk.Entry(self.param_frame, width=3); self.ent_uid.insert(0, "1"); self.ent_uid.pack(side="left")

        ttk.Label(self.param_frame, text="메모리 영역:").pack(side="left", padx=5)
        self.mem_type = ttk.Combobox(self.param_frame, values=["Holding Reg (4x)", "Input Reg (3x)", "Coils (0x)", "Discrete In (1x)"], width=15, state="readonly")
        self.mem_type.current(0); self.mem_type.pack(side="left")

        ttk.Label(self.param_frame, text="시작 주소:").pack(side="left", padx=5)
        self.ent_addr = ttk.Entry(self.param_frame, width=6); self.ent_addr.insert(0, "0"); self.ent_addr.pack(side="left")

        ttk.Label(self.param_frame, text="개수:").pack(side="left", padx=5)
        self.ent_qty = ttk.Entry(self.param_frame, width=5); self.ent_qty.insert(0, "10"); self.ent_qty.pack(side="left")

        # 버튼
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=10)
        self.btn_start = tk.Button(btn_frame, text="서버 시작", bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), command=self.start_server)
        self.btn_start.pack(side="left", fill="x", expand=True, padx=5)
        self.btn_stop = tk.Button(btn_frame, text="서버 중지", bg="#F44336", fg="white", font=("Arial", 10, "bold"), command=self.stop_server, state="disabled")
        self.btn_stop.pack(side="left", fill="x", expand=True, padx=5)

        # 시뮬레이션 설정
        sim_frame = ttk.LabelFrame(parent, text="자동 시뮬레이션 (값 자동 증가)", padding="5")
        sim_frame.pack(fill="x", padx=10, pady=5)
        self.chk_auto = tk.BooleanVar(value=False)
        ttk.Checkbutton(sim_frame, text="동작", variable=self.chk_auto, command=self.toggle_auto_sim).pack(side="left", padx=10)
        ttk.Label(sim_frame, text="주기(초):").pack(side="left")
        self.ent_sim_int = ttk.Entry(sim_frame, width=5); self.ent_sim_int.insert(0, "1.0"); self.ent_sim_int.pack(side="left", padx=5)

        # 데이터 그리드
        grid_frame = ttk.LabelFrame(parent, text="메모리 모니터링 (더블 클릭하여 수정)", padding="5")
        grid_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.tree = ttk.Treeview(grid_frame, columns=("addr", "val", "hex"), show="headings", height=10)
        self.tree.heading("addr", text="주소")
        self.tree.heading("val", text="값 (DEC / Bool)")
        self.tree.heading("hex", text="값 (HEX)")
        self.tree.column("addr", width=100, anchor="center")
        self.tree.column("val", width=150, anchor="center")
        self.tree.column("hex", width=150, anchor="center")
        
        vsb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        
        self.tree.bind("<Double-1>", self.on_double_click)

        self.update_ui_state()
        
        if not MODBUS_SERVER_AVAILABLE:
            self.btn_start.config(state="disabled")
            messagebox.showerror("오류", "pymodbus 라이브러리 로드 실패.\n'pip install pymodbus pyserial'을 확인해주세요.")

    def update_ui_state(self):
        for widget in self.conn_frame.winfo_children():
            widget.pack_forget()
        if self.mode_var.get() == "TCP":
            self.lbl_port.pack(side="left", padx=5)
            self.ent_port.pack(side="left")
        else:
            self.lbl_com.pack(side="left", padx=5)
            self.ent_com.pack(side="left")
            self.lbl_baud.pack(side="left", padx=5)
            self.cmb_baud.pack(side="left")

    def init_datastore(self):
        if not MODBUS_SERVER_AVAILABLE:
            messagebox.showerror("오류", "라이브러리가 로드되지 않았습니다.")
            return False

        try:
            raw_addr = int(self.ent_addr.get())
            qty = int(self.ent_qty.get())
            t_str = self.mem_type.get()
            
            # 데이터 블록 생성
            block = ModbusSequentialDataBlock(raw_addr, [0]*qty)

            # 슬레이브 컨텍스트 생성 (선택된 타입만 블록 할당)
            store = ModbusSlaveContext(
                hr=block if "Holding" in t_str else None,
                ir=block if "Input Reg" in t_str else None,
                co=block if "Coils" in t_str else None,
                di=block if "Discrete" in t_str else None,
            )

            # [수정] GUI에서 직접 데이터에 접근할 수 있도록 슬레이브 컨텍스트를 인스턴스 변수에 저장
            self.slave_context = store

            # [핵심] 단일 컨텍스트 모드 (single=True)
            self.context = ModbusServerContext(slaves=self.slave_context, single=True)
            
            # 그리드 초기화
            self.tree.delete(*self.tree.get_children())
            for i in range(qty):
                ui_addr = raw_addr + i
                self.tree.insert("", "end", iid=str(i), values=(ui_addr, 0, "0x0000"))
            return True
        except Exception as e:
            messagebox.showerror("설정 오류", f"초기화 실패:\n{str(e)}")
            return False

    def start_server(self):
        if not self.init_datastore(): return

        mode = self.mode_var.get()
        self.is_server_running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        
        self.server_thread = threading.Thread(target=self.run_server_thread, args=(mode,), daemon=True)
        self.server_thread.start()
        
        self.parent.after(1000, self.refresh_ui)
        self.toggle_auto_sim()

    def run_server_thread(self, mode):
        try:
            if mode == "TCP":
                port = int(self.ent_port.get())
                StartTcpServer(context=self.context, address=("0.0.0.0", port))
            else:
                com = self.ent_com.get()
                baud = int(self.cmb_baud.get())
                StartSerialServer(context=self.context, port=com, baudrate=baud, timeout=1)
        except Exception as e:
            print(f"서버 스레드 오류: {e}")
        finally:
            self.is_server_running = False

    def stop_server(self):
        if self.is_server_running:
            try:
                ServerStop()
            except Exception as e:
                print(f"서버 중지 오류: {e}")
        
        self.is_server_running = False
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.chk_auto.set(False)
        messagebox.showinfo("알림", "서버가 중지되었습니다.")

    def get_slave_context(self):
        """ [수정] 명시적으로 저장된 슬레이브 컨텍스트를 반환하는 함수 """
        return self.slave_context

    def refresh_ui(self):
        if not self.is_server_running: return
        
        try:
            # [수정] 컨텍스트 접근 방식 개선
            slave = self.get_slave_context()
            if not slave: return

            t_str = self.mem_type.get()
            addr = int(self.ent_addr.get())
            qty = int(self.ent_qty.get())
            
            fc = 3
            if "Holding" in t_str: fc = 3
            elif "Input Reg" in t_str: fc = 4
            elif "Coils" in t_str: fc = 1
            elif "Discrete" in t_str: fc = 2
            
            values = slave.getValues(fc, addr, count=qty)
            
            for i, val in enumerate(values):
                if self.tree.exists(str(i)):
                    current = self.tree.item(str(i))['values']
                    if current[1] != val:
                        ui_addr = current[0]
                        val_int = int(val)
                        self.tree.item(str(i), values=(ui_addr, val_int, f"0x{val_int:04X}"))
        except: pass
        
        self.parent.after(500, self.refresh_ui)

    def toggle_auto_sim(self):
        if self.chk_auto.get() and self.is_server_running:
            if not self.auto_sim_active:
                self.auto_sim_active = True
                self.sim_loop()
        else:
            self.auto_sim_active = False

    def sim_loop(self):
        if not self.auto_sim_active or not self.is_server_running: return
        try:
            slave = self.get_slave_context()
            if not slave: return
            
            t_str = self.mem_type.get()
            addr = int(self.ent_addr.get())
            qty = int(self.ent_qty.get())
            
            fc = 3
            if "Holding" in t_str: fc = 3
            elif "Input Reg" in t_str: fc = 4
            elif "Coils" in t_str: fc = 1
            elif "Discrete" in t_str: fc = 2

            current_vals = slave.getValues(fc, addr, count=qty)
            new_vals = [(v + 1) % 65535 for v in current_vals]
            slave.setValues(fc, addr, new_vals)
            
            interval = float(self.ent_sim_int.get())
            self.parent.after(int(interval * 1000), self.sim_loop)
        except:
            self.auto_sim_active = False

    def on_double_click(self, event):
        if not self.is_server_running: return
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        current_vals = self.tree.item(item_id)['values']
        try:
            addr = int(current_vals[0])
            val = int(current_vals[1])
        except ValueError:
            return
        
        new_val = simpledialog.askinteger("값 수정", f"{addr}번지의 새 값:", initialvalue=val, minvalue=0, maxvalue=65535)
        
        if new_val is not None:
            try:
                slave = self.get_slave_context()
                t_str = self.mem_type.get()
                
                fc = 3
                if "Holding" in t_str: fc = 3
                elif "Input Reg" in t_str: fc = 4
                elif "Coils" in t_str: fc = 1
                elif "Discrete" in t_str: fc = 2 
                
                slave.setValues(fc, addr, [new_val])
                self.tree.item(item_id, values=(addr, new_val, f"0x{new_val:04X}"))
            except Exception as e:
                messagebox.showerror("오류", str(e))
