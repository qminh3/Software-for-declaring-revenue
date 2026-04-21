"""Tab Bao Cao - VAT khau tru, TNCN, NXT, Quy (Nhom 3)"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, datetime
from database import THUE_TNCN_SUAT

def fmt(v):
    try: f=float(v or 0); return f"{f:,.0f}" if f else "0"
    except: return "0"

def pct(v): return f"{float(v or 0)*100:.1f}%"


class BaoCaoTab(tk.Frame):
    def __init__(self,parent,db):
        super().__init__(parent,bg="#F8FAFC")
        self.db=db; self._build(); self.refresh()

    def _build(self):
        # Filter bar
        ctrl=tk.Frame(self,bg="#FFFFFF",highlightbackground="#CBD5E1",highlightthickness=1); ctrl.pack(fill="x",padx=0,pady=(0,12))
        tk.Label(ctrl,text="📊  BÁO CÁO TỔNG HỢP - HỘ KINH DOANH NHÓM 3",
                 font=("Arial",11,"bold"),bg="#FFFFFF",fg="#0F172A").pack(side="left",padx=12,pady=8)
        rf=tk.Frame(ctrl,bg="#FFFFFF"); rf.pack(side="right",padx=8)
        tk.Label(rf,text="Năm:",font=("Arial",10),bg="#FFFFFF").pack(side="left")
        self.var_y=tk.StringVar(value=str(date.today().year))
        ttk.Spinbox(rf,textvariable=self.var_y,from_=2020,to=2035,width=6).pack(side="left",padx=4)
        tk.Label(rf,text="Tháng:",font=("Arial",10),bg="#FFFFFF").pack(side="left",padx=(8,0))
        self.var_mf=tk.StringVar(value="1"); self.var_mt=tk.StringVar(value=str(date.today().month))
        ttk.Combobox(rf,textvariable=self.var_mf,values=[str(i) for i in range(1,13)],width=4,state="readonly").pack(side="left",padx=2)
        tk.Label(rf,text="→",bg="#FFFFFF").pack(side="left")
        ttk.Combobox(rf,textvariable=self.var_mt,values=[str(i) for i in range(1,13)],width=4,state="readonly").pack(side="left",padx=2)
        tk.Button(rf,text="🔍 Xem",font=("Arial",10,"bold"),bg="#2563EB",fg="white",relief="flat",padx=12,pady=4,command=self.refresh).pack(side="left",padx=8)
        tk.Button(rf,text="📥 Excel",font=("Arial",10,"bold"),bg="#14B8A6",fg="white",relief="flat",padx=10,pady=4,command=self._export_all).pack(side="left",padx=4)

        # Notebook bên trong tab
        nb=ttk.Notebook(self); nb.pack(fill="both",expand=True,padx=8,pady=6)

        # --- Panel 1: VAT khấu trừ ---
        self.p_vat=tk.Frame(nb,bg="#FFFFFF"); nb.add(self.p_vat,text="🧮 VAT Khấu Trừ")
        self._build_vat_panel(self.p_vat)

        # --- Panel 2: TNCN (S2c) ---
        self.p_tncn=tk.Frame(nb,bg="#FFFFFF"); nb.add(self.p_tncn,text="💼 TNCN Lợi Nhuận")
        self._build_tncn_panel(self.p_tncn)

        # --- Panel 3: Kho (S2d) ---
        self.p_kho=tk.Frame(nb,bg="#FFFFFF"); nb.add(self.p_kho,text="📦 Kho (S2d)")
        self._build_kho_panel(self.p_kho)

        # --- Panel 4: Quỹ (S2e) ---
        self.p_quy=tk.Frame(nb,bg="#FFFFFF"); nb.add(self.p_quy,text="💰 Quỹ/Ngân hàng (S2e)")
        self._build_quy_panel(self.p_quy)

    # ── VAT panel ────────────────────────────────────────────
    def _build_vat_panel(self,p):
        tk.Label(p,text="THUẾ GTGT - PHƯƠNG PHÁP KHẤU TRỪ (Đầu ra - Đầu vào)",
                 font=("Arial",11,"bold"),bg="#2563EB",fg="white").pack(fill="x",ipady=8)
        self.vat_frame=tk.Frame(p,bg="#FFFFFF",padx=30,pady=20); self.vat_frame.pack(fill="both",expand=True)

    def _refresh_vat(self,y,mf,mt):
        for w in self.vat_frame.winfo_children(): w.destroy()
        vat=self.db.get_vat_khau_tru(y,mf,mt)
        data=[("VAT ĐẦU RA (doanh thu bán hàng × thuế suất)",vat["vat_dau_ra"],"#7C3AED","white"),
              ("VAT ĐẦU VÀO (chi phí mua hàng có hóa đơn)",  vat["vat_dau_vao"],"#5B21B6","white"),
              ("═══ VAT PHẢI NỘP NHÀ NƯỚC = Đầu ra - Đầu vào",vat["vat_phai_nop"],"#EDE9FE","#4C1D95"),
              ("(Nếu âm: VAT còn được khấu trừ kỳ sau)",       vat["vat_con_khau_tru"],"#F5F3FF","#7C3AED")]
        for lbl,val,bg,fg in data:
            row=tk.Frame(self.vat_frame,bg=bg,highlightbackground="#CBD5E1",highlightthickness=1); row.pack(fill="x",pady=4)
            tk.Label(row,text=lbl,font=("Arial",11),bg=bg,fg=fg,anchor="w").pack(side="left",padx=20,pady=12)
            tk.Label(row,text=f"{fmt(val)} đ",font=("Arial",13,"bold"),bg=bg,fg=fg).pack(side="right",padx=24,pady=12)
        tk.Label(self.vat_frame,
                 text="📌 Công thức: VAT phải nộp = Σ VAT đầu ra (cột 'VAT đầu ra' trong tab Bán Hàng)\n"
                      "                       − Σ VAT đầu vào (cột 'VAT đầu vào' trong tab Chi Phí)",
                 font=("Arial",10),bg="#FFFFFF",fg="#1E40AF",justify="left",
                 relief="solid",bd=1).pack(fill="x",pady=16,padx=4,ipady=8)

    # ── TNCN panel ───────────────────────────────────────────
    def _build_tncn_panel(self,p):
        tk.Label(p,text="THUẾ TNCN - TÍNH TRÊN THU NHẬP THỰC TẾ (Doanh thu - Chi phí hợp lý)",
                 font=("Arial",11,"bold"),bg="#2563EB",fg="white").pack(fill="x",ipady=8)
        self.tncn_frame=tk.Frame(p,bg="#FFFFFF",padx=30,pady=20); self.tncn_frame.pack(fill="both",expand=True)

    def _refresh_tncn(self,y,mf,mt):
        for w in self.tncn_frame.winfo_children(): w.destroy()
        s=self.db.get_s2c_summary(y,mf,mt)
        data=[("(A) Tổng Doanh Thu (chưa VAT)",s["tong_doanh_thu"],"#FFFBEB","#92400E"),
              ("(B) Tổng Chi Phí Hợp Lý",       s["tong_chi_phi"],  "#FEF3C7","#92400E"),
              ("(C) Thu nhập tính thuế = A − B", s["thu_nhap_tinh_thue"],"#FDE68A","#78350F"),
              (f"(D) Thuế TNCN = C × {THUE_TNCN_SUAT*100:.0f}%", s["thue_tncn"],"#92400E","white")]
        for lbl,val,bg,fg in data:
            row=tk.Frame(self.tncn_frame,bg=bg,highlightbackground="#CBD5E1",highlightthickness=1); row.pack(fill="x",pady=4)
            tk.Label(row,text=lbl,font=("Arial",11),bg=bg,fg=fg,anchor="w").pack(side="left",padx=20,pady=12)
            tk.Label(row,text=f"{fmt(val)} đ",font=("Arial",13,"bold"),bg=bg,fg=fg).pack(side="right",padx=24,pady=12)
        # Chi phí breakdown
        tk.Label(self.tncn_frame,text="Chi tiết Chi phí theo loại:",
                 font=("Arial",10,"bold"),bg="#FFFFFF",fg="#0F172A").pack(anchor="w",pady=(16,4))
        cps=self.db.get_chi_phi(y,mf,mt)
        by_loai={}
        for r in cps:
            loai=r["loai_cp"] or "Khác"
            by_loai[loai]=by_loai.get(loai,0)+float(r["so_tien"] or 0)
        for loai,tien in sorted(by_loai.items(),key=lambda x:-x[1]):
            row=tk.Frame(self.tncn_frame,bg="#EFF6FF"); row.pack(fill="x",pady=1)
            tk.Label(row,text=f"  • {loai}",font=("Arial",9),bg="#EFF6FF",fg="#0F172A",anchor="w").pack(side="left",padx=8,pady=3)
            tk.Label(row,text=f"{fmt(tien)} đ",font=("Arial",9,"bold"),bg="#EFF6FF",fg="#0F172A").pack(side="right",padx=12,pady=3)

    # ── Kho panel (S2d) ──────────────────────────────────────
    def _build_kho_panel(self,p):
        # Manager button
        brow=tk.Frame(p,bg="#FFFFFF"); brow.pack(fill="x",padx=8,pady=4)
        tk.Button(brow,text="⚙️ Quản lý Hàng hóa & Giá vốn",font=("Arial",10),bg="#10B981",fg="white",
                  relief="flat",padx=10,pady=4,command=self._manage_hang).pack(side="left",padx=4)
        cols=("stt","ten_hang","dvt","ton_dau_nam","ton_hien_tai","sl_xuat","doanh_thu")
        hdrs=("STT","Tên hàng","ĐVT","Tồn đầu năm","Tồn hiện tại","SL xuất kỳ","Doanh thu kỳ")
        ws=(40,140,60,110,110,100,130)
        tf=tk.Frame(p,bg="#FFFFFF"); tf.pack(fill="both",expand=True,padx=8,pady=4)
        vsb=ttk.Scrollbar(tf,orient="vertical")
        self.tree_kho=ttk.Treeview(tf,columns=cols,show="headings",yscrollcommand=vsb.set)
        vsb.config(command=self.tree_kho.yview)
        for c,h,w in zip(cols,hdrs,ws):
            self.tree_kho.heading(c,text=h)
            self.tree_kho.column(c,width=w,anchor="e" if c not in ("stt","ten_hang","dvt") else ("center" if c=="stt" else "w"))
        self.tree_kho.pack(side="left",fill="both",expand=True); vsb.pack(side="right",fill="y")
        self.tree_kho.tag_configure("ok",background="#F8FAFC"); self.tree_kho.tag_configure("low",background="#FEE2E2")
        self.tree_kho.tag_configure("tot",background="#0F172A",foreground="white",font=("Arial",10,"bold"))
        self.lbl_kho=tk.Label(p,text="",font=("Arial",10,"bold"),bg="#2563EB",fg="white"); self.lbl_kho.pack(fill="x",padx=8,pady=(0,4),ipady=6)

    def _refresh_kho(self,y,mf,mt):
        data=self.db.get_nxt_data(y,mf,mt); self.tree_kho.delete(*self.tree_kho.get_children())
        tot_xuat=tot_dt=0
        for i,r in enumerate(data):
            tag="low" if float(r["ton_hien_tai"] or 0)<0 else "ok"
            self.tree_kho.insert("","end",tags=(tag,),values=(i+1,r["ten_hang"],r["dvt"] or "",
                fmt(r["ton_dau_nam"]),fmt(r["ton_hien_tai"]),fmt(r["sl_xuat"]),fmt(r["doanh_thu"])))
            tot_xuat+=float(r["sl_xuat"] or 0); tot_dt+=float(r["doanh_thu"] or 0)
        self.tree_kho.insert("","end",tags=("tot",),values=("","TỔNG","","","",fmt(tot_xuat),fmt(tot_dt)))
        self.lbl_kho.config(text=f"   Tổng xuất kỳ: {fmt(tot_xuat)} sl  |  Doanh thu kỳ: {fmt(tot_dt)} đ")

    def _manage_hang(self):
        from tabs.tab_banhang import fmt
        HangManager(self,self.db,on_close=self.refresh)

    # ── Quỹ panel (S2e) ──────────────────────────────────────
    def _build_quy_panel(self,p):
        brow=tk.Frame(p,bg="#FFFFFF"); brow.pack(fill="x",padx=8,pady=4)
        tk.Button(brow,text="⚙️ Quản lý Quỹ/Ngân hàng",font=("Arial",10),bg="#2563EB",fg="white",
                  relief="flat",padx=10,pady=4,command=self._manage_quy).pack(side="left",padx=4)
        cols=("stt","tk","du_dn","du_ht","du_dk","thu","chi","du_ck")
        hdrs=("STT","Tài khoản","Dư đầu năm","Dư hiện tại","Dư đầu kỳ","Tổng thu","Tổng chi","Dư cuối kỳ")
        ws=(40,120,130,130,130,130,130,130)
        tf=tk.Frame(p,bg="#FFFFFF"); tf.pack(fill="both",expand=True,padx=8,pady=4)
        vsb=ttk.Scrollbar(tf,orient="vertical")
        self.tree_quy=ttk.Treeview(tf,columns=cols,show="headings",yscrollcommand=vsb.set)
        vsb.config(command=self.tree_quy.yview)
        for c,h,w in zip(cols,hdrs,ws):
            self.tree_quy.heading(c,text=h)
            self.tree_quy.column(c,width=w,anchor="e" if c not in ("stt","tk") else ("center" if c=="stt" else "w"))
        self.tree_quy.pack(side="left",fill="both",expand=True); vsb.pack(side="right",fill="y")
        self.tree_quy.tag_configure("tm",background="#FEFCE8"); self.tree_quy.tag_configure("bank",background="#EFF6FF")
        self.tree_quy.tag_configure("tot",background="#0F172A",foreground="white",font=("Arial",10,"bold"))
        self.tree_quy.tag_configure("neg",background="#FEE2E2")
        self.lbl_quy=tk.Label(p,text="",font=("Arial",10,"bold"),bg="#2563EB",fg="white"); self.lbl_quy.pack(fill="x",padx=8,pady=(0,4),ipady=6)

    def _refresh_quy(self,y,mf,mt):
        data=self.db.get_s2e_data(y,mf,mt); self.tree_quy.delete(*self.tree_quy.get_children())
        tot_thu=tot_chi=tot_dc=0
        for i,r in enumerate(data):
            dc=float(r["du_cuoi_ky"] or 0); tag="neg" if dc<0 else ("tm" if r["tai_khoan"]=="TM" else "bank")
            self.tree_quy.insert("","end",tags=(tag,),values=(i+1,r["tai_khoan"],fmt(r["du_dau_nam"]),
                fmt(r["du_hien_tai"]),fmt(r["du_dau_ky"]),fmt(r["tong_thu"]),fmt(r["tong_chi"]),fmt(dc)))
            tot_thu+=float(r["tong_thu"] or 0); tot_chi+=float(r["tong_chi"] or 0); tot_dc+=dc
        self.tree_quy.insert("","end",tags=("tot",),values=("","TỔNG","","","",fmt(tot_thu),fmt(tot_chi),fmt(tot_dc)))
        self.lbl_quy.config(text=f"   Tổng thu kỳ: {fmt(tot_thu)} đ  |  Tổng chi kỳ: {fmt(tot_chi)} đ  |  Dư cuối kỳ: {fmt(tot_dc)} đ")

    def _manage_quy(self): QuyManager(self,self.db,on_close=self.refresh)

    # ── Main refresh ─────────────────────────────────────────
    def refresh(self):
        try: y=int(self.var_y.get()); mf=int(self.var_mf.get()); mt=int(self.var_mt.get())
        except: return
        self._refresh_vat(y,mf,mt); self._refresh_tncn(y,mf,mt)
        self._refresh_kho(y,mf,mt); self._refresh_quy(y,mf,mt)

    def _export_all(self):
        from excel_export_nhom3 import export_baocao
        try: y=int(self.var_y.get()); mf=int(self.var_mf.get()); mt=int(self.var_mt.get())
        except: return
        fp=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],
           initialfile=f"BaoCao_Nhom3_{y}_T{mf}-T{mt}.xlsx")
        if not fp: return
        try: export_baocao(self.db,y,mf,mt,fp); messagebox.showinfo("OK",f"Đã xuất:\n{fp}")
        except Exception as ex: messagebox.showerror("Lỗi",str(ex))


class HangManager(tk.Toplevel):
    def __init__(self,parent,db,on_close=None):
        super().__init__(parent); self.db=db; self.on_close=on_close
        self.title("Quản lý Hàng hóa & Giá vốn")
        self.geometry("620x450"); self.configure(bg="#F8FAFC"); self.grab_set()
        self._build(); self._load()
        self.protocol("WM_DELETE_WINDOW",self._close)
    def _build(self):
        tk.Label(self,text="⚙️  Quản lý Hàng hóa & Tồn đầu năm",
                 font=("Arial",12,"bold"),bg="#10B981",fg="white").pack(fill="x",ipady=8)
        form=tk.Frame(self,bg="#F8FAFC",padx=16,pady=10); form.pack(fill="x")
        for lbl,attr,w in [("Tên hàng:","e_ten",20),("ĐVT:","e_dvt",8),("Tồn đầu năm:","e_ton",12),("Giá vốn:","e_gia",12)]:
            col=len([k for k in form.children])*2
            tk.Label(form,text=lbl,font=("Arial",10),bg="#F8FAFC").grid(row=0,column=col,sticky="e",padx=4)
            e=tk.Entry(form,font=("Arial",10),width=w,relief="solid",bd=1); e.grid(row=0,column=col+1,padx=4,pady=4)
            setattr(self,attr,e)
        bf=tk.Frame(form,bg="#F8FAFC"); bf.grid(row=0,column=8,padx=8)
        tk.Button(bf,text="➕ Thêm/Sửa",font=("Arial",10),bg="#10B981",fg="white",relief="flat",padx=8,pady=4,command=self._save).pack(pady=2)
        tk.Button(bf,text="🗑️ Xóa",font=("Arial",10),bg="#DC2626",fg="white",relief="flat",padx=8,pady=4,command=self._del).pack(pady=2)
        cols=("ten","dvt","ton","gia")
        self.tree=ttk.Treeview(self,columns=cols,show="headings",height=12)
        for c,h,w in zip(cols,("Tên hàng","ĐVT","Tồn đầu năm","Giá vốn"),(220,70,110,110)):
            self.tree.heading(c,text=h); self.tree.column(c,width=w,anchor="w" if c in ("ten","dvt") else "e")
        self.tree.pack(fill="both",expand=True,padx=10,pady=4)
        self.tree.bind("<<TreeviewSelect>>",self._sel)
        tk.Button(self,text="✅ Đóng",font=("Arial",10,"bold"),bg="#10B981",fg="white",relief="flat",padx=16,pady=6,command=self._close).pack(pady=8)
    def _load(self):
        self.tree.delete(*self.tree.get_children())
        for h in self.db.get_hang_hoa():
            self.tree.insert("","end",iid=h["ten_hang"],values=(h["ten_hang"],h["dvt"] or "",
                f"{float(h['ton_dau_nam'] or 0):,.0f}",f"{float(h['gia_von'] or 0):,.0f}"))
    def _sel(self,e):
        sel=self.tree.selection()
        if sel:
            r=self.tree.item(sel[0])["values"]
            for attr,val in [("e_ten",r[0]),("e_dvt",r[1]),("e_ton",str(r[2]).replace(",","")),("e_gia",str(r[3]).replace(",",""))]:
                w=getattr(self,attr); w.delete(0,"end"); w.insert(0,val)
    def _save(self):
        ten=self.e_ten.get().strip()
        if not ten: messagebox.showwarning("","Nhập tên hàng!"); return
        try: ton=float(self.e_ton.get().replace(",","") or 0)
        except: ton=0
        try: gia=float(self.e_gia.get().replace(",","") or 0)
        except: gia=0
        self.db.upsert_hang_hoa(ten,self.e_dvt.get().strip(),ton,gia); self._load()
    def _del(self):
        sel=self.tree.selection()
        if sel and messagebox.askyesno("Xóa",f"Xóa '{sel[0]}'?"): self.db.delete_hang_hoa(sel[0]); self._load()
    def _close(self):
        if self.on_close: self.on_close()
        self.destroy()


class QuyManager(tk.Toplevel):
    def __init__(self,parent,db,on_close=None):
        super().__init__(parent); self.db=db; self.on_close=on_close
        self.title("Quản lý Quỹ / Tài khoản")
        self.geometry("450,380".replace(",","x")); self.geometry("450x380"); self.configure(bg="#F8FAFC"); self.grab_set()
        self._build(); self._load()
        self.protocol("WM_DELETE_WINDOW",self._close)
    def _build(self):
        tk.Label(self,text="⚙️  Quản lý Quỹ & Tài khoản ngân hàng",
                 font=("Arial",12,"bold"),bg="#2563EB",fg="white").pack(fill="x",ipady=8)
        form=tk.Frame(self,bg="#F8FAFC",padx=16,pady=10); form.pack(fill="x")
        tk.Label(form,text="Tài khoản:",font=("Arial",10),bg="#F8FAFC").grid(row=0,column=0,sticky="e",padx=4)
        self.e_tk=tk.Entry(form,font=("Arial",10),width=16,relief="solid",bd=1); self.e_tk.grid(row=0,column=1,padx=4,pady=4)
        tk.Label(form,text="Dư đầu năm:",font=("Arial",10),bg="#F8FAFC").grid(row=1,column=0,sticky="e",padx=4)
        self.e_du=tk.Entry(form,font=("Arial",10),width=16,relief="solid",bd=1); self.e_du.grid(row=1,column=1,padx=4,pady=4)
        bf=tk.Frame(form,bg="#F8FAFC"); bf.grid(row=0,column=2,rowspan=2,padx=8)
        tk.Button(bf,text="➕ Thêm/Sửa",font=("Arial",10),bg="#2563EB",fg="white",relief="flat",padx=8,pady=4,command=self._save).pack(pady=2)
        tk.Button(bf,text="🗑️ Xóa",font=("Arial",10),bg="#DC2626",fg="white",relief="flat",padx=8,pady=4,command=self._del).pack(pady=2)
        cols=("tk","du"); self.tree=ttk.Treeview(self,columns=cols,show="headings",height=10)
        for c,h,w in zip(cols,("Tài khoản","Dư đầu năm"),(200,150)):
            self.tree.heading(c,text=h); self.tree.column(c,width=w,anchor="w" if c=="tk" else "e")
        self.tree.pack(fill="both",expand=True,padx=10,pady=4)
        self.tree.bind("<<TreeviewSelect>>",self._sel)
        tk.Button(self,text="✅ Đóng",font=("Arial",10,"bold"),bg="#2563EB",fg="white",relief="flat",padx=16,pady=6,command=self._close).pack(pady=8)
    def _load(self):
        self.tree.delete(*self.tree.get_children())
        for q in self.db.get_quy_tk(): self.tree.insert("","end",iid=q["tai_khoan"],values=(q["tai_khoan"],f"{float(q['du_dau_nam'] or 0):,.0f}"))
    def _sel(self,e):
        sel=self.tree.selection()
        if sel:
            r=self.tree.item(sel[0])["values"]
            self.e_tk.delete(0,"end"); self.e_tk.insert(0,r[0])
            self.e_du.delete(0,"end"); self.e_du.insert(0,str(r[1]).replace(",",""))
    def _save(self):
        tk=self.e_tk.get().strip()
        if not tk: messagebox.showwarning("","Nhập tên TK!"); return
        try: du=float(self.e_du.get().replace(",","") or 0)
        except: du=0
        self.db.upsert_quy_tk(tk,du); self._load()
    def _del(self):
        sel=self.tree.selection()
        if sel and messagebox.askyesno("Xóa",f"Xóa '{sel[0]}'?"): self.db.delete_quy_tk(sel[0]); self._load()
    def _close(self):
        if self.on_close: self.on_close()
        self.destroy()
