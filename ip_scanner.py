import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import ipaddress
import threading
import re
import socket
import csv
import webbrowser
import os
from queue import Queue

# 외부 라이브러리 로딩
try:
    from mac_vendor_lookup import MacLookup
except ImportError:
    MacLookup = None

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None


class IPScannerGUI:
    def __init__(self, parent):
        self.parent = parent
        self.is_scanning = False
        self.mac_lookup = None
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)

        # 내 PC 정보
        info_frame = ttk.LabelFrame(parent, text="내 PC 정보 (Localhost)", padding="10")
        info_frame.pack(fill="x", padx=10, pady=5)
        
        self.lbl_my_ip = ttk.Label(info_frame, text="IP: 확인 중...", font=("맑은 고딕", 10, "bold"), foreground="blue")
        self.lbl_my_ip.pack(side="left", padx=10)
        self.lbl_my_mac = ttk.Label(info_frame, text="MAC: 확인 중...")
        self.lbl_my_mac.pack(side="left", padx=10)
        self.lbl_my_host = ttk.Label(info_frame, text="Host: 확인 중...")
        self.lbl_my_host.pack(side="left", padx=10)
        
        # 스캔 설정
        top_frame = ttk.LabelFrame(parent, text="스캔 설정", padding="10")
        top_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(top_frame, text="시작 IP:").grid(row=0, column=0, padx=5, sticky="w")
        self.start_ip_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.start_ip_var, width=15).grid(row=0, column=1, padx=5)
        
        ttk.Label(top_frame, text="~").grid(row=0, column=2, padx=5)
        
        ttk.Label(top_frame, text="종료 IP:").grid(row=0, column=3, padx=5, sticky="w")
        self.end_ip_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.end_ip_var, width=15).grid(row=0, column=4, padx=5)
        
        self.detail_scan_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top_frame, text="상세 스캔 (호스트명/포트)", variable=self.detail_scan_var).grid(row=0, column=5, padx=15)

        btn_frame = ttk.Frame(top_frame)
        btn_frame.grid(row=0, column=6, padx=10)
        self.btn_scan = ttk.Button(btn_frame, text="스캔 시작", command=self.start_scan)
        self.btn_scan.pack(side="left", padx=2)
        self.btn_export = ttk.Button(btn_frame, text="CSV 저장", command=self.export_csv)
        self.btn_export.pack(side="left", padx=2)
        self.btn_export_excel = ttk.Button(btn_frame, text="Excel 저장", command=self.export_excel)
        self.btn_export_excel.pack(side="left", padx=2)
        
        # 결과 리스트
        tree_frame = ttk.LabelFrame(parent, text="스캔 결과", padding="10")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        columns = ("ip", "latency", "hostname", "mac", "vendor", "ports")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        self.tree.heading("ip", text="IP 주소", command=lambda: self.sort_tree("ip", False))
        self.tree.heading("latency", text="응답(ms)", command=lambda: self.sort_tree("latency", False))
        self.tree.heading("hostname", text="호스트명", command=lambda: self.sort_tree("hostname", False))
        self.tree.heading("mac", text="MAC 주소", command=lambda: self.sort_tree("mac", False))
        self.tree.heading("vendor", text="제조사", command=lambda: self.sort_tree("vendor", False))
        self.tree.heading("ports", text="열린 포트 (주요)", command=lambda: self.sort_tree("ports", False))
        
        self.tree.column("ip", width=110, anchor="center")
        self.tree.column("latency", width=70, anchor="center")
        self.tree.column("hostname", width=150, anchor="w")
        self.tree.column("mac", width=130, anchor="center")
        self.tree.column("vendor", width=180, anchor="w")
        self.tree.column("ports", width=150, anchor="center")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.tree.tag_configure("odd", background="white")
        self.tree.tag_configure("even", background="#f0f0f0")
        self.tree.tag_configure("slow", foreground="red")
        
        # --- 이벤트 바인딩 ---
        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="CMD Ping 실행", command=self.cmd_ping_target)
        self.context_menu.add_command(label="웹 접속 (HTTP)", command=self.open_web)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="IP 복사", command=self.copy_ip)

        self.tree.bind("<Button-1>", self.on_single_click)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)

        # --- 하단 설명 및 상태바 ---
        help_frame = ttk.Frame(parent, padding=5)
        help_frame.pack(side="bottom", fill="x")
        ttk.Label(help_frame, text="[사용법] 좌클릭: IP 복사 | 더블클릭: PING 테스트 | 우클릭: 추가 메뉴", foreground="gray").pack(pady=2)

        self.status_var = tk.StringVar(value="대기 중...")
        self.lbl_status = ttk.Label(parent, textvariable=self.status_var, relief="sunken", anchor="w")
        self.lbl_status.pack(side="bottom", fill="x")
        
        self.auto_fill_ip()
        self.init_mac_lookup()
        self.parent.after(1000, self.update_local_host_info)

    def on_single_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            ip = self.get_selected_ip()
            if ip:
                self.copy_ip()
                original_status = self.status_var.get()
                self.status_var.set(f"IP: {ip} 가 클립보드에 복사되었습니다.")
                self.parent.after(2000, lambda: self.status_var.get() == f"IP: {ip} 가 클립보드에 복사되었습니다." and self.status_var.set(original_status))

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.cmd_ping_target()

    def get_local_ip(self):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            local_ip = s.getsockname()[0]
            s.close()
            return local_ip
        except:
            return "192.168.0.1"

    def auto_fill_ip(self):
        local_ip = self.get_local_ip()
        parts = local_ip.split('.')
        base_ip = ".".join(parts[:3])
        self.start_ip_var.set(f"{base_ip}.1")
        self.end_ip_var.set(f"{base_ip}.254")

    def init_mac_lookup(self):
        if MacLookup is None:
            self.status_var.set("MacLookup 모듈 없음. 제조사 정보가 표시되지 않습니다.")
            return
        def _load():
            try:
                self.status_var.set("제조사 DB 로딩 중...")
                self.mac_lookup = MacLookup()
                self.status_var.set("대기 중")
            except:
                self.status_var.set("제조사 DB 로드 실패")
                self.mac_lookup = None
        threading.Thread(target=_load, daemon=True).start()

    def update_local_host_info(self):
        try:
            my_ip = self.get_local_ip()
            my_host = socket.gethostname()
            
            import uuid
            mac_num = uuid.getnode()
            my_mac = ':'.join(['{:02x}'.format((mac_num >> elements) & 0xff) for elements in range(0,48,8)][::-1])
            
            self.lbl_my_ip.config(text=f"IP: {my_ip}")
            self.lbl_my_mac.config(text=f"MAC: {my_mac}")
            self.lbl_my_host.config(text=f"Host: {my_host}")
        except Exception as e:
            self.lbl_my_ip.config(text="IP 확인 실패")

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def get_selected_ip(self):
        selected = self.tree.selection()
        if selected: return self.tree.item(selected[0])['values'][0]
        return None

    def cmd_ping_target(self):
        ip = self.get_selected_ip()
        if ip: os.system(f"start cmd /k ping {ip} -t")

    def open_web(self):
        ip = self.get_selected_ip()
        if ip: webbrowser.open(f"http://{ip}")

    def copy_ip(self):
        ip = self.get_selected_ip()
        if ip:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(ip)
            self.parent.update()

    def export_csv(self):
        if not self.tree.get_children():
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV 파일", "*.csv")])
        if not filename: return
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["IP", "응답속도", "호스트명", "MAC", "제조사", "포트"])
                for item in self.tree.get_children():
                    row = self.tree.item(item)['values']
                    writer.writerow(row)
            messagebox.showinfo("성공", "파일이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    def export_excel(self):
        if Workbook is None:
            messagebox.showerror("오류", "Excel 저장을 위해 'openpyxl' 라이브러리가 필요합니다.\n\n터미널에서 'pip install openpyxl' 명령을 실행하세요.")
            return
        if not self.tree.get_children():
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
        if not filename: return
        
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "IP Scan Results"
            
            # 헤더 추가
            headers = ["IP", "응답속도", "호스트명", "MAC", "제조사", "포트"]
            ws.append(headers)
            
            # 데이터 추가
            for item in self.tree.get_children():
                row = self.tree.item(item)['values']
                ws.append(row)
            
            wb.save(filename)
            messagebox.showinfo("성공", "파일이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    def get_mac_address_arp(self, ip):
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW 
            output = subprocess.check_output(f"arp -a {ip}", startupinfo=startupinfo, shell=True).decode('cp949', errors='ignore')
            mac_regex = re.search(r"([0-9a-fA-F]{2}-){5}[0-9a-fA-F]{2}", output)
            if mac_regex: return mac_regex.group(0).replace('-', ':').lower()
        except: pass
        return None

    def check_ports(self, ip, ports=[502, 80, 443, 21, 23, 3389]):
        open_ports = []
        for port in ports:
            if not self.is_scanning: break
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(0.3)
                if s.connect_ex((str(ip), port)) == 0: open_ports.append(str(port))
                s.close()
            except: pass
        return ",".join(open_ports) if open_ports else ""

    def get_hostname(self, ip):
        try: return socket.gethostbyaddr(str(ip))[0]
        except: return ""

    def worker(self, queue, detailed):
        while self.is_scanning:
            try:
                ip = queue.get_nowait()
            except:
                break # 큐가 비었으면 종료
            
            try:
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                proc = subprocess.Popen(['ping', '-n', '1', '-w', '200', str(ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo)
                stdout, _ = proc.communicate()

                if proc.returncode == 0 and self.is_scanning:
                    output = stdout.decode('cp949', errors='ignore')
                    latency = 1
                    lat_match = re.search(r'(time|시간)[=<](\d+)ms', output)
                    if lat_match: latency = int(lat_match.group(2))
                    
                    mac = self.get_mac_address_arp(str(ip))
                    vendor = "" if not mac else "알수없음"
                    if mac and self.mac_lookup:
                        try: vendor = self.mac_lookup.lookup(mac)
                        except: pass
                    
                    hostname, ports = ("", "")
                    if detailed and self.is_scanning:
                        hostname = self.get_hostname(ip)
                        ports = self.check_ports(ip)
                    
                    if self.is_scanning:
                        self.parent.after(0, self.insert_result, str(ip), latency, hostname, mac, vendor, ports)
            except Exception as e:
                print(f"Error scanning {ip}: {e}")
            finally:
                queue.task_done()

    def insert_result(self, ip, latency, hostname, mac, vendor, ports):
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][0] == ip:
                return
        count = len(self.tree.get_children())
        tag = "even" if count % 2 == 0 else "odd"
        tags = [tag, "slow"] if latency >= 100 else [tag]
        self.tree.insert("", "end", values=(ip, f"{latency}ms", hostname, mac, vendor, ports), tags=tags)

    def run_scan_thread(self):
        s_ip = self.start_ip_var.get().strip()
        e_ip = self.end_ip_var.get().strip()
        is_detail = self.detail_scan_var.get()
        try:
            start, end = ipaddress.IPv4Address(s_ip), ipaddress.IPv4Address(e_ip)
            if start > end: start, end = end, start
            ip_list = [ipaddress.IPv4Address(ip) for ip in range(int(start), int(end) + 1)]
        except ValueError:
            # IP 주소 형식이 잘못되었을 경우, GUI 스레드에서 에러 메시지를 표시하고 종료합니다.
            self.parent.after(0, messagebox.showerror, "오류", "IP 형식이 잘못되었습니다.")
            self.parent.after(0, self.finish_scan, True) # 스캔이 중지된 것으로 처리
            return

        self.parent.after(0, lambda: self.status_var.set(f"스캔 중... (대상: {len(ip_list)}개)"))
        scan_queue = Queue()
        for ip in ip_list: scan_queue.put(ip)
        
        threads = []
        num_threads = 30 if is_detail else 50
        for _ in range(num_threads):
            t = threading.Thread(target=self.worker, args=(scan_queue, is_detail))
            t.daemon = True
            t.start()
            threads.append(t)
        
        # 모든 워커 스레드가 작업을 마칠 때까지 백그라운드에서 대기합니다.
        for t in threads:
            t.join()

        # 모든 스레드 작업이 끝나면, GUI 스레드에서 finish_scan을 호출합니다.
        # self.is_scanning이 False이면 사용자가 중지 버튼을 누른 것입니다.
        is_stopped_by_user = not self.is_scanning
        self.parent.after(0, self.finish_scan, is_stopped_by_user)

    def start_scan(self):
        if self.is_scanning: return
        for item in self.tree.get_children(): self.tree.delete(item)
        
        self.is_scanning = True
        self.btn_scan.config(text="스캔 중지", command=self.stop_scan)
        
        threading.Thread(target=self.run_scan_thread, daemon=True).start()

    def stop_scan(self):
        if self.is_scanning:
            # 사용자에게 즉시 피드백을 주고, 중복 클릭을 방지합니다.
            self.btn_scan.config(text="중지 중...", state="disabled")
            self.status_var.set("스캔 중지 중...")
            self.is_scanning = False

    def finish_scan(self, is_stopped):
        # 스캔 완료 또는 중지 시, is_scanning 상태를 확실히 False로 변경합니다.
        was_scanning_before_finish = self.is_scanning
        self.is_scanning = False 
        
        # 스캔 버튼을 다시 '스캔 시작' 상태로 완전히 복구합니다.
        self.btn_scan.config(text="스캔 시작", command=self.start_scan, state="normal")
        
        # 결과가 있을 경우에만 IP 주소로 기본 정렬을 수행합니다.
        if self.tree.get_children():
            self.sort_tree("ip", False) 
        
        if is_stopped:
            self.status_var.set(f"스캔이 중지되었습니다. {len(self.tree.get_children())}개 장치 발견.")
        else:
            self.status_var.set(f"스캔 완료. {len(self.tree.get_children())}개 장치 발견.")
            # 사용자가 '중지'를 누른게 아니라, 정상적으로 끝났을 때만 완료 메시지를 표시합니다.
            if was_scanning_before_finish:
                messagebox.showinfo("완료", "스캔이 완료되었습니다.")

    def sort_tree(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        try:
            if col == "ip": l.sort(key=lambda t: ipaddress.IPv4Address(t[0]), reverse=reverse)
            elif col == "latency": l.sort(key=lambda t: int(str(t[0]).replace('ms','').replace('(나)','0').strip()), reverse=reverse)
            else: l.sort(reverse=reverse)
        except: l.sort(reverse=reverse)
        
        for i, (val, k) in enumerate(l):
            self.tree.move(k, '', i)
            tag = "even" if i % 2 == 0 else "odd"
            latency_val = self.tree.item(k, 'values')[1]
            try:
                latency_int = int(latency_val.replace('ms','').strip())
                tags = [tag, "slow"] if latency_int >= 100 else [tag]
                self.tree.item(k, tags=tags)
            except:
                self.tree.item(k, tags=[tag])

        self.tree.heading(col, command=lambda: self.sort_tree(col, not reverse))
