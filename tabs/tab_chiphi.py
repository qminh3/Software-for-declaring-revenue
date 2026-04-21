"""Tab S2c - So chi tiet doanh thu, chi phi (Nhom 3)"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, datetime
from database import LOAI_CHI_PHI

def fmt(v):
    try: f=float(v or 0); return f"{f:,.0f}" if f else ""
    except: return str(v) if v else ""

class ChiPhiTab(tk.Frame):
    def __init__(self,parent,db):
        super().__init__(parent,bg="#F8FAFC")
        self.db=db; self._sel=None; self._build(); self.refresh()

    def _build(self):
        tb=tk.Frame(self,bg="#FEF3C7",highlightbackground="#E2E8F0",highlightthickness=1); tb.pack(fill="x",padx=0,pady=(0,8))
        for txt,clr,cmd in [("➕ Thêm","#D97706",self._new),("✏️ Sửa","#2563EB",self._edit),
                             ("🗑️ Xóa","#DC2626",self._del),("🔄 Mới","#64748B",self.refresh),
                             ("📥 Excel","#0D9488",self._export)]:
            tk.Button(tb,text=txt,font=("Arial",10,"bold"),bg=clr,fg="white",
                      relief="flat",padx=12,pady=6,command=cmd).pack(side="left",padx=4,pady=6)
        rf=tk.Frame(tb,bg="#FEF3C7"); rf.pack(side="right",padx=8)
        tk.Label(rf,text="Năm:",font=("Arial",10),bg="#FEF3C7").pack(side="left")
        self.var_y=tk.StringVar(value=str(date.today().year))
        ttk.Spinbox(rf,textvariable=self.var_y,from_=2020,to=2035,width=6).pack(side="left",padx=4)
        tk.Label(rf,text="Tháng:",font=("Arial",10),bg="#FEF3C7").pack(side="left",padx=(8,0))
        self.var_m=tk.StringVar(value="Tất cả")
        ttk.Combobox(rf,textvariable=self.var_m,values=["Tất cả"]+[str(i) for i in range(1,13)],
                     width=7,state="readonly").pack(side="left",padx=4)
        tk.Button(rf,text="🔍 Lọc",font=("Arial",10),bg="#2563EB",fg="white",
                  relief="flat",padx=10,pady=4,command=self.refresh).pack(side="left",padx=4)
        # Summary
        self.sf=tk.Frame(self,bg="#FDE68A"); self.sf.pack(fill="x",padx=8,pady=2)
        self.lbl=tk.Label(self.sf,text="",font=("Arial",9),bg="#FDE68A",fg="#0F172A")
        self.lbl.pack(side="left",padx=10,pady=4)
        # Tree
        cols=("ngay","so_ct","dien_giai","ten_khach","loai_cp","so_tien","vat_vao","tk_chi")
        hdrs=("Ngày","Số CT","Diễn giải","NCC","Loại chi phí","Số tiền (đ)","VAT đầu vào (đ)","TK Chi")
        ws=(85,85,180,130,160,130,130,70)
        tf=tk.Frame(self,bg="#F8FAFC"); tf.pack(fill="both",expand=True,padx=8,pady=4)
        vsb=ttk.Scrollbar(tf,orient="vertical"); hsb=ttk.Scrollbar(tf,orient="horizontal")
        self.tree=ttk.Treeview(tf,columns=cols,show="headings",yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview); hsb.config(command=self.tree.xview)
        for c,h,w in zip(cols,hdrs,ws):
            self.tree.heading(c,text=h)
            self.tree.column(c,width=w,anchor="e" if c in ("so_tien","vat_vao") else "w")
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        tf.grid_rowconfigure(0,weight=1); tf.grid_columnconfigure(0,weight=1)
        self.tree.tag_configure("nhap",background="#FEF9C3")
        self.tree.tag_configure("luong",background="#DBEAFE")
        self.tree.tag_configure("khac",background="#FFF7ED")
        self.tree.bind("<<TreeviewSelect>>",lambda e: setattr(self,"_sel",int(self.tree.selection()[0])) if self.tree.selection() else None)
        self.tree.bind("<Double-1>",lambda e: self._edit())

    def refresh(self):
        y=int(self.var_y.get()); m=self.var_m.get()
        mf=1 if m=="Tất cả" else int(m); mt=12 if m=="Tất cả" else int(m)
        rows=self.db.get_chi_phi(y,mf,mt); self.tree.delete(*self.tree.get_children())
        tot_cp=tot_vat=0
        for r in rows:
            ngay=r["ngay"]
            try: ngay=datetime.strptime(str(ngay),"%Y-%m-%d").strftime("%d/%m/%Y")
            except: pass
            cp=float(r["so_tien"] or 0); vat=float(r["vat_dau_vao"] or 0)
            tot_cp+=cp; tot_vat+=vat
            loai=r["loai_cp"] or ""
            tag="nhap" if "hàng hóa" in loai.lower() else ("luong" if "lương" in loai.lower() else "khac")
            self.tree.insert("","end",iid=r["id"],tags=(tag,),values=(
                ngay,r["so_ct"] or "",r["dien_giai"] or "",r["ten_khach"] or "",
                loai,fmt(cp),fmt(vat),r["tk_chi"] or ""))
        # S2c summary
        s2c=self.db.get_s2c_summary(y,mf,mt)
        self.lbl.config(text=f"  Chi phí: {fmt(tot_cp)} đ  |  VAT đầu vào: {fmt(tot_vat)} đ  ||  "
                             f"Doanh thu: {fmt(s2c['tong_doanh_thu'])} đ  |  "
                             f"Thu nhập tính thuế: {fmt(s2c['thu_nhap_tinh_thue'])} đ  |  "
                             f"Thuế TNCN 17%: {fmt(s2c['thue_tncn'])} đ")

    def _new(self): ChiPhiForm(self,self.db,None,self.refresh)
    def _edit(self):
        if not self._sel: messagebox.showwarning("","Chọn dòng!"); return
        ChiPhiForm(self,self.db,self._sel,self.refresh)
    def _del(self):
        if not self._sel: messagebox.showwarning("","Chọn dòng!"); return
        if messagebox.askyesno("Xóa","Xóa dòng này?"): self.db.delete_chi_phi(self._sel); self._sel=None; self.refresh()
    def _export(self):
        from excel_export_nhom3 import export_s2c
        y=int(self.var_y.get()); m=self.var_m.get()
        mf=1 if m=="Tất cả" else int(m); mt=12 if m=="Tất cả" else int(m)
        fp=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],initialfile=f"S2c_ChiPhi_{y}.xlsx")
        if not fp: return
        try: export_s2c(self.db,y,mf,mt,fp); messagebox.showinfo("OK",f"Đã xuất:\n{fp}")
        except Exception as e: messagebox.showerror("Lỗi",str(e))


class ChiPhiForm(tk.Toplevel):
    def __init__(self,parent,db,rid=None,on_save=None):
        super().__init__(parent); self.db=db; self.rid=rid; self.on_save=on_save
        self.title("Thêm chi phí" if not rid else "Sửa chi phí")
        sw=self.winfo_screenwidth(); sh=self.winfo_screenheight()
        w,h=min(680,sw-40),min(480,sh-80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.configure(bg="#F8FAFC"); self.grab_set(); self._build()
        if rid: self._load(rid)

    def _build(self):
        tk.Label(self,text="🧾  NHẬP CHI PHÍ (S2c-HKD)",font=("Arial",13,"bold"),bg="#2563EB",fg="white").pack(fill="x",ipady=8)
        body=tk.Frame(self,bg="#F8FAFC",padx=20,pady=12); body.pack(fill="both",expand=True)
        quy=self.db.get_quy_names()

        def r(t,row,col): tk.Label(body,text=t,font=("Arial",9),bg="#F8FAFC",fg="#374151",anchor="e").grid(row=row,column=col,sticky="e",padx=(4,6),pady=6)
        def e(row,col,w=18): ent=tk.Entry(body,font=("Arial",10),width=w,relief="solid",bd=1,bg="white"); ent.grid(row=row,column=col,sticky="ew",padx=(0,12),pady=6); return ent

        r("Ngày:*",0,0); self.e_ngay=e(0,1,w=13); self.e_ngay.insert(0,date.today().strftime("%d/%m/%Y"))
        r("Số CT:",0,2); self.e_soct=e(0,3,w=14)
        r("Diễn giải:",1,0); self.e_dg=tk.Entry(body,font=("Arial",10),relief="solid",bd=1,bg="white"); self.e_dg.grid(row=1,column=1,columnspan=3,sticky="ew",padx=(0,12),pady=6)
        r("NCC/Đối tượng:",2,0); self.e_khach=tk.Entry(body,font=("Arial",10),relief="solid",bd=1,bg="white"); self.e_khach.grid(row=2,column=1,columnspan=3,sticky="ew",padx=(0,12),pady=6)
        r("Loại chi phí:",3,0)
        self.cb_loai=ttk.Combobox(body,values=LOAI_CHI_PHI,font=("Arial",10),state="readonly",width=30); self.cb_loai.grid(row=3,column=1,columnspan=3,sticky="ew",padx=(0,12),pady=6); self.cb_loai.current(0)
        r("Số tiền CP:",4,0)
        self.e_cp=tk.Entry(body,font=("Arial",10),width=16,relief="solid",bd=1,bg="#FEF3C7"); self.e_cp.grid(row=4,column=1,sticky="ew",padx=(0,12),pady=6)
        r("VAT đầu vào:",4,2)
        self.e_vat=tk.Entry(body,font=("Arial",10),width=16,relief="solid",bd=1,bg="#FDE68A"); self.e_vat.grid(row=4,column=3,sticky="ew",padx=(0,12),pady=6)
        r("TK Chi:",5,0)
        self.cb_tk=ttk.Combobox(body,values=quy,font=("Arial",10),width=12); self.cb_tk.grid(row=5,column=1,sticky="w",padx=(0,12),pady=6)
        if quy: self.cb_tk.current(0)
        tk.Label(body,text="💡 VAT đầu vào ghi vào đây để khấu trừ với VAT đầu ra khi tính thuế",
                 font=("Arial",9),bg="#F8FAFC",fg="#0F172A").grid(row=5,column=2,columnspan=2,sticky="w",padx=4)
        body.columnconfigure(1,weight=1); body.columnconfigure(3,weight=1)

        bf=tk.Frame(self,bg="#F8FAFC"); bf.pack(fill="x",padx=20,pady=12)
        tk.Button(bf,text="💾  LƯU",font=("Arial",11,"bold"),bg="#2563EB",fg="white",relief="flat",padx=24,pady=10,command=self._save).pack(side="left",padx=6)
        tk.Button(bf,text="❌  Hủy",font=("Arial",11),bg="#64748B",fg="white",relief="flat",padx=20,pady=10,command=self.destroy).pack(side="left",padx=6)

    def _gf(self,w):
        try: return float(w.get().replace(",","").strip() or 0)
        except: return 0.0

    def _load(self,rid):
        r=self.db.get_chi_phi_by_id(rid)
        if not r: return
        def se(w,v): w.delete(0,"end"); (w.insert(0,str(v)) if v and str(v).strip() not in ("","0","0.0") else None)
        ngay=r["ngay"]
        try: ngay=datetime.strptime(str(ngay),"%Y-%m-%d").strftime("%d/%m/%Y")
        except: pass
        self.e_ngay.delete(0,"end"); self.e_ngay.insert(0,ngay or "")
        se(self.e_soct,r["so_ct"]); se(self.e_dg,r["dien_giai"]); se(self.e_khach,r["ten_khach"])
        if r["loai_cp"] in LOAI_CHI_PHI: self.cb_loai.set(r["loai_cp"])
        se(self.e_cp,fmt(r["so_tien"])); se(self.e_vat,fmt(r["vat_dau_vao"]))
        if r["tk_chi"] in self.db.get_quy_names(): self.cb_tk.set(r["tk_chi"])

    def _save(self):
        ns=self.e_ngay.get().strip()
        try: ngay=datetime.strptime(ns,"%d/%m/%Y").strftime("%Y-%m-%d")
        except: messagebox.showerror("Lỗi","Ngày sai định dạng!"); return
        d={"ngay":ngay,"so_ct":self.e_soct.get().strip(),"dien_giai":self.e_dg.get().strip(),
           "ten_khach":self.e_khach.get().strip(),"loai_cp":self.cb_loai.get(),
           "so_tien":self._gf(self.e_cp),"vat_dau_vao":self._gf(self.e_vat),
           "tk_chi":self.cb_tk.get().strip()}
        if self.rid: self.db.update_chi_phi(self.rid,d)
        else: self.db.insert_chi_phi(d)
        if self.on_save: self.on_save()
        self.destroy()
