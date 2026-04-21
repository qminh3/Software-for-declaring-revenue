"""
Quan Ly Doanh Thu

"""
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date
import sys, os

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from database import Database
from tabs.tab_banhang import BanHangTab
from tabs.tab_chiphi import ChiPhiTab
from tabs.tab_baocao import BaoCaoTab
from tabs.tab_bangke_thumua import BangKeThuMuaTab

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Quản lý Doanh Thu")
        self.geometry("1400x820")
        self.minsize(1200, 700)
        self.configure(bg="#F8FAFC")
        
        # Icon
        _ico = resource_path("icon.ico")
        _png = resource_path("icon.png")
        try:
            if os.path.exists(_ico): self.iconbitmap(_ico)
            elif os.path.exists(_png):
                img = tk.PhotoImage(file=_png); self.iconphoto(True, img)
        except: pass

        self.db = Database()
        self._setup_styles()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_styles(self):
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("Treeview", background="white", foreground="#334155",
                    rowheight=26, fieldbackground="white", font=("Arial",9), borderwidth=0)
        s.configure("Treeview.Heading", background="#F8FAFC", foreground="#475569",
                    font=("Arial",9,"bold"), relief="flat", borderwidth=1, bordercolor="#E2E8F0")
        s.map("Treeview", background=[("selected","#E0F2FE")], foreground=[("selected","#0369A1")])
        s.map("Treeview.Heading", background=[("active","#F1F5F9")])

    def _build_ui(self):
        self.status_var = tk.StringVar(value="Sẵn sàng")
        sb = tk.Frame(self, bg="#E2E8F0", height=28)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        tk.Label(sb, textvariable=self.status_var, font=("Arial",9), bg="#E2E8F0", fg="#475569").pack(side="left", padx=10, pady=4)
        tk.Label(sb, text="  VAT Khấu Trừ  ",
                 font=("Arial",9), bg="#E2E8F0", fg="#64748B").pack(side="right", padx=10)

        main_container = tk.Frame(self, bg="#F8FAFC")
        main_container.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(main_container, bg="#1E293B", width=220)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        tk.Label(self.sidebar, text="QUẢN LÝ DOANH THU", font=("Arial",13,"bold"), bg="#1E293B", fg="white", pady=16).pack()
        tk.Frame(self.sidebar, bg="#334155", height=1).pack(fill="x", padx=15, pady=(0, 16))

        # Content Area (Right)
        self.content_area = tk.Frame(main_container, bg="#F1F5F9")
        self.content_area.pack(side="right", fill="both", expand=True)
        
        # Header in Content Area
        hdr = tk.Frame(self.content_area, bg="#FFFFFF", height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Frame(self.content_area, bg="#E2E8F0", height=1).pack(fill="x") # Shadow line

        self.lbl_title = tk.Label(hdr, text="Dashboard", font=("Arial",14,"bold"), bg="#FFFFFF", fg="#0F172A")
        self.lbl_title.pack(side="left", padx=24)

        info = tk.Frame(hdr, bg="#FFFFFF"); info.pack(side="right", padx=24)
        cfg = self.db.get_config()
        tk.Label(info, text=f"{cfg.get('ten_hkd','')}", font=("Arial",10,"bold"), bg="#FFFFFF", fg="#2563EB").pack(anchor="e")
        tk.Label(info, text=f"MST: {cfg.get('mst','')} | {cfg.get('dia_chi','')}", font=("Arial",9), bg="#FFFFFF", fg="#64748B").pack(anchor="e")

        # Container for screens
        self.screen_container = tk.Frame(self.content_area, bg="#F1F5F9")
        self.screen_container.pack(fill="both", expand=True, padx=20, pady=16)

        self.screens = {}
        self.screens['banhang'] = {"widget": BanHangTab(self.screen_container, self.db), "title": "Sổ Doanh Thu (S2b)"}
        self.screens['chiphi'] = {"widget": ChiPhiTab(self.screen_container, self.db), "title": "Sổ Chi Phí (S2c)"}
        self.screens['bangke'] = {"widget": BangKeThuMuaTab(self.screen_container, self.db), "title": "Bảng Kê NV (02/TNDN)"}
        self.screens['baocao'] = {"widget": BaoCaoTab(self.screen_container, self.db), "title": "Báo Cáo Thuế & Sổ Quỹ"}
        self.screens['caidat'] = {"widget": self._build_settings_screen(), "title": "Cài Đặt Hệ Thống"}

        # Sidebar Buttons
        self.nav_btns = {}
        menus = [
            ("banhang", "Sổ Doanh Thu "),
            ("chiphi", "Sổ Chi Phí "), 
            ("bangke", "Bảng Kê Đầu vào "),
            ("baocao", "Báo Cáo "),
            ("caidat", "Cài Đặt Hệ Thống")
        ]
        
        for key, text in menus:
            btn = tk.Button(self.sidebar, text=f"   {text}", font=("Arial", 11), bg="#1E293B", fg="#94A3B8",
                            anchor="w", relief="flat", padx=16, pady=12,
                            activebackground="#334155", activeforeground="white",
                            command=lambda k=key: self._switch_screen(k))
            btn.pack(fill="x", padx=8, pady=2)
            
            # Hover effects
            def on_enter(e, b=btn, k=key):
                if self.current_screen != k: b.config(bg="#334155", fg="white")
            def on_leave(e, b=btn, k=key):
                if self.current_screen != k: b.config(bg="#1E293B", fg="#94A3B8")
                
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            self.nav_btns[key] = btn

        self.current_screen = None
        self._switch_screen('banhang')

    def _switch_screen(self, screen_key):
        if self.current_screen == screen_key: return
        
        # Hide old
        if self.current_screen:
            self.screens[self.current_screen]["widget"].pack_forget()
            old_btn = self.nav_btns[self.current_screen]
            old_btn.config(bg="#1E293B", fg="#94A3B8", font=("Arial", 11))
            
        # Show new
        self.current_screen = screen_key
        self.lbl_title.config(text=self.screens[screen_key]["title"])
        
        widget = self.screens[screen_key]["widget"]
        widget.pack(fill="both", expand=True)
        
        # Trigger refresh if needed
        if screen_key == "baocao" and hasattr(widget, "refresh"): widget.refresh()
        if screen_key == "chiphi" and hasattr(widget, "refresh"): widget.refresh()
        
        # Highlight new button
        new_btn = self.nav_btns[screen_key]
        new_btn.config(bg="#2563EB", fg="white", font=("Arial", 11, "bold"))
        self.set_status(f"Đang xem: {self.screens[screen_key]['title']}")

    def _build_settings_screen(self):
        frm = tk.Frame(self.screen_container, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1)
        
        content = tk.Frame(frm, bg="#FFFFFF")
        content.place(relx=0.5, rely=0.4, anchor="center")
        
        tk.Label(content, text="⚙️  Thông tin Hộ Kinh Doanh",
                 font=("Arial",16,"bold"), fg="#0F172A", bg="#FFFFFF").grid(row=0, column=0, columnspan=2, pady=(0,20))
                 
        # Info box
        info_box = tk.Frame(content, bg="#EFF6FF", highlightbackground="#BFDBFE", highlightthickness=1)
        info_box.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0,24), ipady=12)
        tk.Label(info_box, text="📋 HỘ KINH DOANH NHÓM 3 (theo TT152/2025/TT-BTC):",
                 font=("Arial",10,"bold"), bg="#EFF6FF", fg="#1E3A8A").pack(anchor="w", padx=20, pady=(8,4))
        for line in ["• Kê khai phương pháp KHẤU TRỪ",
                     "• Thuế TNCN: 17% trên Thu nhập tính thuế",
                     "• Sử dụng Hóa đơn điện tử có mã CQT",
                     "• Bắt buộc báo cáo S2b, S2c, S2d, S2e"]:
            tk.Label(info_box, text=line, font=("Arial",9), bg="#EFF6FF", fg="#1E40AF").pack(anchor="w", padx=36, pady=2)
            
        cfg = self.db.get_config()
        fields = [("Tên Hộ Kinh Doanh:","ten_hkd"),("Mã số thuế (MST):","mst"),("Địa chỉ:","dia_chi")]
        self._entries = {}
        for i,(lbl,key) in enumerate(fields, 2):
            tk.Label(content, text=lbl, font=("Arial",10,"bold"), bg="#FFFFFF", fg="#475569", anchor="w"
                     ).grid(row=i*2, column=0, sticky="w", padx=4, pady=(8,2))
            e = tk.Entry(content, font=("Arial",11), width=44, relief="solid", bd=1, bg="#F8FAFC")
            e.insert(0, cfg.get(key,""))
            e.grid(row=i*2+1, column=0, sticky="ew", padx=4, pady=(0,8), ipady=4)
            self._entries[key] = e
            
        tk.Button(content, text="Lưu Các Thay Đổi",
                  font=("Arial",11,"bold"), bg="#2563EB", fg="white",
                  relief="flat", padx=24, pady=8, command=self._save_settings
                  ).grid(row=len(fields)*2+2, column=0, pady=24)
        return frm

    def _save_settings(self):
        cfg = {k: e.get().strip() for k,e in self._entries.items()}
        self.db.save_config(cfg)
        messagebox.showinfo("OK", "Đã lưu thông tin HKD!")
        self.set_status("Đã cập nhật thông tin HKD")

    def set_status(self, msg):
        self.status_var.set(f"{datetime.now().strftime('%H:%M:%S')}  |  {msg}")

    def _on_close(self):
        self.db.close(); self.destroy()

if __name__ == "__main__":
    app = App(); app.mainloop()
