import tkinter as tk
from tkinter import ttk
import struct
import re

class AddressCalculatorGUI:
    def __init__(self, parent):
        self.parent = parent
        
        paned = ttk.PanedWindow(parent, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)
        
        # [왼쪽] 기능 패널
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=3)
        
        base_frame = ttk.LabelFrame(left_frame, text="데이터 변환기", padding="15")
        base_frame.pack(fill="x", padx=10, pady=10)
        
        self.is_signed = tk.BooleanVar(value=False)
        self.dec_var = tk.StringVar(); self.dec_var.trace_add('write', self.on_dec)
        self.hex_var = tk.StringVar(); self.hex_var.trace_add('write', self.on_hex)
        self.bin_var = tk.StringVar(); self.bin_var.trace_add('write', self.on_bin)
        self.float_var = tk.StringVar(); self.float_var.trace_add('write', self.on_float)
        self.ascii_var = tk.StringVar()
        self.updating = False

        self.create_row(base_frame, 0, "DEC:", self.dec_var)
        ttk.Checkbutton(base_frame, text="부호 있음(Signed)", variable=self.is_signed, command=lambda: self.on_hex()).grid(row=0, column=2, sticky="w")
        
        self.create_row(base_frame, 1, "HEX:", self.hex_var)
        f_swap = ttk.Frame(base_frame); f_swap.grid(row=1, column=2, sticky="w")
        ttk.Button(f_swap, text="워드 스왑", command=self.swap_word, width=10).pack(side="left")
        ttk.Button(f_swap, text="바이트 스왑", command=self.swap_byte, width=10).pack(side="left")

        self.create_row(base_frame, 2, "BIN:", self.bin_var)
        self.create_row(base_frame, 3, "FLOAT:", self.float_var, color="blue")
        ttk.Label(base_frame, text="(IEEE 754)", foreground="gray").grid(row=3, column=2, sticky="w")
        self.create_row(base_frame, 4, "ASCII:", self.ascii_var, readonly=True)

        plc_frame = ttk.LabelFrame(left_frame, text="PLC 주소 변환", padding="15")
        plc_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(plc_frame, text="입력 (예: MW23, MB46.1):").pack(anchor="w")
        self.plc_input = ttk.Entry(plc_frame, font=("Consolas", 11))
        self.plc_input.pack(fill="x", pady=5)
        
        # 중요: 키 입력이 떼질 때마다 calc_plc 함수를 실행합니다.
        self.plc_input.bind("<KeyRelease>", self.calc_plc)

        self.res_frame = ttk.Frame(plc_frame); self.res_frame.pack(fill="x", pady=5)
        self.addr_vars = {}
        for idx, t in enumerate(["MX (Bit)", "MB (Byte)", "MW (Word)", "MD (DWord)", "ML (LWord)"]):
            ttk.Label(self.res_frame, text=t+":", width=12, anchor="e").grid(row=idx, column=0, pady=2)
            v = tk.StringVar(value="-")
            self.addr_vars[t] = v
            lbl = ttk.Label(self.res_frame, textvariable=v, width=30, foreground="blue", relief="solid")
            lbl.grid(row=idx, column=1, padx=5, pady=2)
            self.bind_copy(lbl, v)

        # [오른쪽] 설명 패널
        right_frame = ttk.LabelFrame(paned, text="도움말 및 가이드", padding="10")
        paned.add(right_frame, weight=1)
        
        lbl1 = ttk.Label(right_frame, text="[데이터 변환기 사용법]\n1. DEC/HEX/BIN/FLOAT 칸에\n   값을 입력하면 자동 변환됩니다.\n2. Word Swap: 상/하위 2바이트 교환\n3. Byte Swap: 바이트 순서 역순\n\n(클릭 시 클립보드 복사)", font=("맑은 고딕", 9), justify="left")
        lbl1.pack(anchor="nw", pady=10)
        
        ttk.Separator(right_frame, orient="horizontal").pack(fill="x", pady=10)
        
        lbl2 = ttk.Label(right_frame, text="[PLC 주소 변환기]\n1. 미쓰비시/LS 스타일 주소 입력\n   (예: MW100, MB10.5)\n2. 입력 시 Bit/Byte/Word/DWord\n   단위 주소를 자동 계산합니다.\n3. 결과 라벨 클릭 시 복사됩니다.", font=("맑은 고딕", 9), justify="left")
        lbl2.pack(anchor="nw", pady=10)
        
        self.always_top = tk.BooleanVar(value=False)
        ttk.Checkbutton(right_frame, text="항상 위에 표시", variable=self.always_top, command=self.toggle_top).pack(side="bottom", anchor="se")

    def create_row(self, parent, row, label, var, color="black", readonly=False):
        ttk.Label(parent, text=label, font=("Arial", 10, "bold"), foreground=color).grid(row=row, column=0, sticky="e", pady=5)
        state = "readonly" if readonly else "normal"
        entry = ttk.Entry(parent, textvariable=var, width=25, font=("Consolas", 11), state=state)
        entry.grid(row=row, column=1, padx=5, pady=5)
        self.bind_copy(entry)

    def bind_copy(self, widget, var=None):
        def copy(e):
            txt = widget.get() if isinstance(widget, ttk.Entry) else var.get()
            if txt and txt != "-":
                self.parent.clipboard_clear(); self.parent.clipboard_append(txt); self.parent.update()
                self.toast("복사되었습니다!")
        widget.bind("<Button-1>", copy)

    def toast(self, msg):
        top = tk.Toplevel(self.parent); top.overrideredirect(True); top.attributes("-topmost", True)
        x = self.parent.winfo_rootx() + self.parent.winfo_width()//2 - 50
        y = self.parent.winfo_rooty() + self.parent.winfo_height()//2
        top.geometry(f"+{x}+{y}")
        tk.Label(top, text=msg, bg="black", fg="white", padx=10, pady=5).pack()
        top.after(800, top.destroy)

    def toggle_top(self):
        self.parent.winfo_toplevel().attributes('-topmost', self.always_top.get())

    def on_dec(self, *a):
        if self.updating: return
        self.updating = True
        try:
            v = int(self.dec_var.get())
            self.hex_var.set(f"{v & 0xFFFFFFFF:X}")
            self.bin_var.set(f"{v & 0xFFFFFFFF:b}")
            self.float_var.set(str(float(v)))
            self.ascii_var.set(self.to_ascii(v))
        except: pass
        self.updating = False

    def on_hex(self, *a):
        if self.updating: return
        self.updating = True
        try:
            v = int(self.hex_var.get(), 16)
            if self.is_signed.get():
                if v <= 0xFFFF: v = struct.unpack('>h', struct.pack('>H', v))[0]
                else: v = struct.unpack('>i', struct.pack('>I', v & 0xFFFFFFFF))[0]
            self.dec_var.set(str(v))
            self.bin_var.set(f"{int(self.hex_var.get(), 16):b}")
            if v <= 0xFFFFFFFF:
                f = struct.unpack('!f', struct.pack('!I', int(self.hex_var.get(), 16)))[0]
                self.float_var.set(f"{f:.4f}")
            self.ascii_var.set(self.to_ascii(int(self.hex_var.get(), 16)))
        except: pass
        self.updating = False

    def on_bin(self, *a):
        if self.updating: return
        self.updating = True
        try:
            v = int(self.bin_var.get(), 2)
            self.hex_var.set(f"{v:X}"); self.dec_var.set(str(v))
        except: pass
        self.updating = False

    def on_float(self, *a):
        if self.updating: return
        self.updating = True
        try:
            f = float(self.float_var.get())
            i = struct.unpack('!I', struct.pack('!f', f))[0]
            self.hex_var.set(f"{i:X}"); self.bin_var.set(f"{i:b}"); self.dec_var.set(str(i))
        except: pass
        self.updating = False

    def to_ascii(self, v):
        try:
            b = v.to_bytes((v.bit_length()+7)//8, 'big')
            return "".join([chr(x) if 32<=x<=126 else '.' for x in b])
        except: return ""

    def swap_word(self):
        try:
            v = int(self.hex_var.get(), 16)
            if v > 0xFFFF: new = ((v & 0xFFFF) << 16) | ((v >> 16) & 0xFFFF)
            else: new = ((v & 0xFF) << 8) | ((v >> 8) & 0xFF)
            self.hex_var.set(f"{new:X}")
        except: pass

    def swap_byte(self):
        try:
            v = int(self.hex_var.get(), 16)
            b = bytearray(v.to_bytes(4, 'big'))
            b[0], b[1] = b[1], b[0]; b[2], b[3] = b[3], b[2]
            self.hex_var.set(f"{int.from_bytes(b, 'big'):X}")
        except: pass

    def calc_plc(self, e):
        raw = self.plc_input.get().upper().strip()
        match = re.search(r'^([A-Z]+)(\d+)(\.(\d+))?$', raw)
        if not match: return
        
        t, idx = match.group(1), int(match.group(2))
        bit = int(match.group(4)) if match.group(4) else 0
        
        g_bit = 0
        if t == "MX": g_bit = idx
        elif t == "MB": g_bit = idx * 8 + bit
        elif t == "MW": g_bit = idx * 16 + bit
        elif t == "MD": g_bit = idx * 32 + bit
        elif t == "ML": g_bit = idx * 64 + bit
        
        self.addr_vars["MX (Bit)"].set(f"%MX{g_bit}")
        self.addr_vars["MB (Byte)"].set(f"%MB{g_bit//8}" + (f".{g_bit%8}" if g_bit%8 else ""))
        self.addr_vars["MW (Word)"].set(f"%MW{g_bit//16}" + (f".{g_bit%16}" if g_bit%16 else ""))
        self.addr_vars["MD (DWord)"].set(f"%MD{g_bit//32}" + (f".{g_bit%32}" if g_bit%32 else ""))
        self.addr_vars["ML (LWord)"].set(f"%ML{g_bit//64}" + (f".{g_bit%64}" if g_bit%64 else ""))