"""Tab Bảng kê thu mua hàng hóa, dịch vụ (Mẫu 02/TNDN)"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, datetime

def fmt(v):
    try: f=float(v or 0); return f"{f:,.0f}" if f else ""
    except: return str(v) if v else ""

class BangKeThuMuaTab(tk.Frame):
    def __init__(self, parent, db):
        super().__init__(parent, bg="#F8FAFC")
        self.db=db; self._sel=None; self._build(); self._load()

    def _build(self):
        tb=tk.Frame(self,bg="#FFFFFF",highlightbackground="#E2E8F0",highlightthickness=1); tb.pack(fill="x",padx=0,pady=(0,8))
        for txt,clr,cmd in [("➕ Thêm","#16A34A",self._new),("✏️ Sửa","#2563EB",self._edit),
                             ("🗑️ Xóa","#DC2626",self._del),("🔄 Mới","#64748B",self._load),
                             ("📥 Excel","#0D9488",self._export)]:
            tk.Button(tb,text=txt,font=("Arial",10,"bold"),bg=clr,fg="white",
                      relief="flat",padx=12,pady=6,command=cmd).pack(side="left",padx=4,pady=6)
        rf=tk.Frame(tb,bg="#FFFFFF"); rf.pack(side="right",padx=8)
        tk.Label(rf,text="Năm:",font=("Arial",10),bg="#FFFFFF").pack(side="left")
        self.var_y=tk.StringVar(value=str(date.today().year))
        ttk.Spinbox(rf,textvariable=self.var_y,from_=2020,to=2035,width=6).pack(side="left",padx=4)
        tk.Label(rf,text="Tháng:",font=("Arial",10),bg="#FFFFFF").pack(side="left",padx=(8,0))
        self.var_m=tk.StringVar(value="Tất cả")
        ttk.Combobox(rf,textvariable=self.var_m,values=["Tất cả"]+[str(i) for i in range(1,13)],
                     width=7,state="readonly").pack(side="left",padx=4)
        tk.Button(rf,text="🔍 Lọc",font=("Arial",10),bg="#2563EB",fg="white",
                  relief="flat",padx=10,pady=4,command=self._load).pack(side="left",padx=4)
        # Stats
        self.sf=tk.Frame(self,bg="#EFF6FF"); self.sf.pack(fill="x",padx=8,pady=2)
        self.lbl=tk.Label(self.sf,text="",font=("Arial",9),bg="#EFF6FF",fg="#0F172A")
        self.lbl.pack(side="left",padx=10,pady=4)
        # Tree
        cols=("ngay","ten_nguoi_ban","dia_chi","so_cccd","sdt","ten_hang","sl","don_gia","thanh_tien","ghi_chu")
        hdrs=("Ngày","Người bán","Địa chỉ","Số CCCD","SĐT","Tên hàng/DV","SL","Đơn giá","Tổng thanh toán","Ghi chú")
        ws=(85,140,150,110,90,140,65,100,120,100)
        tf=tk.Frame(self,bg="#F8FAFC"); tf.pack(fill="both",expand=True,padx=8,pady=4)
        vsb=ttk.Scrollbar(tf,orient="vertical"); hsb=ttk.Scrollbar(tf,orient="horizontal")
        self.tree=ttk.Treeview(tf,columns=cols,show="headings",yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview); hsb.config(command=self.tree.xview)
        for c,h,w in zip(cols,hdrs,ws):
            self.tree.heading(c,text=h)
            self.tree.column(c,width=w,anchor="e" if c in ("sl","don_gia","thanh_tien") else "w")
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        tf.grid_rowconfigure(0,weight=1); tf.grid_columnconfigure(0,weight=1)
        self.tree.tag_configure("row",background="white"); self.tree.tag_configure("alt",background="#F1F5F9")
        self.tree.bind("<<TreeviewSelect>>",lambda e: setattr(self,"_sel",int(self.tree.selection()[0])) if self.tree.selection() else None)
        self.tree.bind("<Double-1>",lambda e: self._edit())

    def _load(self):
        y=int(self.var_y.get()); m=self.var_m.get()
        mf=1 if m=="Tất cả" else int(m); mt=12 if m=="Tất cả" else int(m)
        rows=self.db.get_bang_ke(y,mf,mt); self.tree.delete(*self.tree.get_children())
        tot_tt=0
        for i,r in enumerate(rows):
            ngay=r["ngay"]
            try: ngay=datetime.strptime(str(ngay),"%Y-%m-%d").strftime("%d/%m/%Y")
            except: pass
            tt=float(r["thanh_tien"] or 0); tot_tt+=tt
            tag="alt" if i%2 else "row"
            self.tree.insert("","end",iid=r["id"],tags=(tag,),values=(
                ngay,r["ten_nguoi_ban"] or "",r["dia_chi"] or "",r["so_cccd"] or "",r["sdt"] or "",
                r["ten_hang_hoa"] or "",fmt(r["sl"]),fmt(r["don_gia"]),fmt(tt),r["ghi_chu"] or ""))
        self.lbl.config(text=f"  {len(rows)} dòng  |  Tổng thanh toán: {fmt(tot_tt)} đ")

    def _new(self): BangKeForm(self,self.db,None,self._load)
    def _edit(self):
        if not self._sel: messagebox.showwarning("","Chọn dòng!"); return
        BangKeForm(self,self.db,self._sel,self._load)
    def _del(self):
        if not self._sel: messagebox.showwarning("","Chọn dòng!"); return
        if messagebox.askyesno("Xóa","Xóa dòng này?"): self.db.delete_bang_ke(self._sel); self._sel=None; self._load()
    def _export(self):
        from excel_export_nhom3 import export_bang_ke
        y=int(self.var_y.get()); m=self.var_m.get()
        mf=1 if m=="Tất cả" else int(m); mt=12 if m=="Tất cả" else int(m)
        fp=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],initialfile=f"02_TNDN_BangKeThuMua_{y}.xlsx")
        if not fp: return
        try: export_bang_ke(self.db,y,mf,mt,fp); messagebox.showinfo("OK",f"Đã xuất:\n{fp}")
        except Exception as e: messagebox.showerror("Lỗi",str(e))

class BangKeForm(tk.Toplevel):
    def __init__(self,parent,db,rid=None,on_save=None):
        super().__init__(parent); self.db=db; self.rid=rid; self.on_save=on_save
        self.title("Thêm bảng kê thu mua" if not rid else "Sửa bảng kê thu mua")
        sw=self.winfo_screenwidth(); sh=self.winfo_screenheight()
        w,h=min(760,self.winfo_screenheight()-100),min(580,sh-80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.configure(bg="#F8FAFC"); self.grab_set(); self._build()
        if rid: self._load_row(rid)

    def _build(self):
        tk.Label(self,text="📝  BẢNG KÊ THU MUA (Mẫu 02/TNDN)",
                 font=("Arial",13,"bold"),bg="#2563EB",fg="white").pack(fill="x",ipady=8)
        canvas=tk.Canvas(self,bg="#F8FAFC",highlightthickness=0)
        vsb=ttk.Scrollbar(self,orient="vertical",command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set); vsb.pack(side="right",fill="y"); canvas.pack(fill="both",expand=True)
        body=tk.Frame(canvas,bg="#F8FAFC"); wid=canvas.create_window((0,0),window=body,anchor="nw")
        body.bind("<Configure>",lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",lambda e: canvas.itemconfig(wid,width=e.width))
        canvas.bind_all("<MouseWheel>",lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units") if canvas.winfo_exists() else None)

        def r(p,t,row,col,**g): tk.Label(p,text=t,font=("Arial",9),bg=p["bg"],fg="#374151",anchor="e").grid(row=row,column=col,sticky="e",padx=(4,4),pady=4,**g)
        def e(p,row,col,w=14,**g): ent=tk.Entry(p,font=("Arial",10),width=w,relief="solid",bd=1,bg="white"); ent.grid(row=row,column=col,sticky="ew",padx=(0,8),pady=4,**g); return ent

        # Section 1: Người bán
        sec1=tk.LabelFrame(body,text=" 👤  Thông tin người bán ",font=("Arial",10,"bold"),fg="#0F172A",bg="#FFFFFF",padx=12,pady=8)
        sec1.pack(fill="x",padx=12,pady=(10,5))
        r(sec1,"Ngày thu mua:*",0,0); self.e_ngay=e(sec1,0,1,w=13)
        self.e_ngay.insert(0,date.today().strftime("%d/%m/%Y"))
        r(sec1,"Tên người bán:",1,0); self.e_ban=e(sec1,1,1,w=30)
        self.e_ban.grid(columnspan=3)
        r(sec1,"Số CCCD/CMND:",2,0); self.e_cccd=e(sec1,2,1)
        r(sec1,"Điện thoại:",2,2); self.e_sdt=e(sec1,2,3)
        r(sec1,"Địa chỉ:",3,0); self.e_dc=e(sec1,3,1,w=30)
        self.e_dc.grid(columnspan=3)
        sec1.columnconfigure(1,weight=1); sec1.columnconfigure(3,weight=1)

        # Section 2: Hóa đơn
        sec2=tk.LabelFrame(body,text=" 📦  Hàng hóa / Dịch vụ ",font=("Arial",10,"bold"),fg="#0F172A",bg="#FFFFFF",padx=12,pady=8)
        sec2.pack(fill="x",padx=12,pady=5)
        r(sec2,"Tên hàng hóa, DV:",0,0); self.e_hang=e(sec2,0,1,w=30)
        self.e_hang.grid(columnspan=3)
        r(sec2,"Số lượng:",1,0); self.e_sl=e(sec2,1,1,w=12)
        r(sec2,"Đơn giá:",1,2); self.e_dg=e(sec2,1,3,w=14)
        self.e_sl.bind("<FocusOut>",self._calc); self.e_sl.bind("<Return>",self._calc)
        self.e_dg.bind("<FocusOut>",self._calc); self.e_dg.bind("<Return>",self._calc)
        r(sec2,"Tổng thanh toán:",2,0); self.e_tt=tk.Entry(sec2,font=("Arial",10,"bold"),width=14,relief="solid",bd=1,bg="#EFF6FF")
        self.e_tt.grid(row=2,column=1,sticky="ew",padx=(0,8),pady=4)
        tk.Button(sec2,text="⚡ Tính",font=("Arial",10,"bold"),bg="#2563EB",fg="white",relief="flat",padx=12,pady=6,command=self._calc).grid(row=2,column=2,columnspan=2,sticky="w",padx=12)
        r(sec2,"Ghi chú:",3,0); self.e_gc=e(sec2,3,1,w=30)
        self.e_gc.grid(columnspan=3)
        sec2.columnconfigure(1,weight=1); sec2.columnconfigure(3,weight=1)

        # Buttons
        bf=tk.Frame(body,bg="#F8FAFC"); bf.pack(fill="x",padx=12,pady=12)
        tk.Button(bf,text="💾  LƯU",font=("Arial",11,"bold"),bg="#2563EB",fg="white",relief="flat",padx=24,pady=10,command=self._save).pack(side="left",padx=6)
        tk.Button(bf,text="❌  Hủy",font=("Arial",11),bg="#64748B",fg="white",relief="flat",padx=20,pady=10,command=self.destroy).pack(side="left",padx=6)

    def _gf(self,w):
        try: return float(w.get().replace(",","").strip() or 0)
        except: return 0.0

    def _calc(self,e=None):
        sl=self._gf(self.e_sl); dg=self._gf(self.e_dg); tt=sl*dg
        self.e_tt.delete(0,"end")
        if tt: self.e_tt.insert(0,f"{tt:,.0f}")

    def _load_row(self,rid):
        r=self.db.get_bang_ke_by_id(rid)
        if not r: return
        def se(w,v):
            w.delete(0,"end")
            if v is not None and str(v).strip() not in ("","0","0.0"): w.insert(0,str(v))
        ngay=r["ngay"]
        try: ngay=datetime.strptime(str(ngay),"%Y-%m-%d").strftime("%d/%m/%Y")
        except: pass
        self.e_ngay.delete(0,"end"); self.e_ngay.insert(0,ngay or "")
        se(self.e_ban,r["ten_nguoi_ban"]); se(self.e_dc,r["dia_chi"])
        se(self.e_cccd,r["so_cccd"]); se(self.e_sdt,r["sdt"])
        se(self.e_hang,r["ten_hang_hoa"]); se(self.e_gc,r["ghi_chu"])
        se(self.e_sl,fmt(r["sl"])); se(self.e_dg,fmt(r["don_gia"]))
        se(self.e_tt,fmt(r["thanh_tien"]))

    def _save(self):
        ns=self.e_ngay.get().strip()
        try: ngay=datetime.strptime(ns,"%d/%m/%Y").strftime("%Y-%m-%d")
        except: messagebox.showerror("Lỗi","Ngày sai định dạng DD/MM/YYYY!"); return
        d={"ngay":ngay,"ten_nguoi_ban":self.e_ban.get().strip(),"dia_chi":self.e_dc.get().strip(),
           "so_cccd":self.e_cccd.get().strip(),"sdt":self.e_sdt.get().strip(),
           "ten_hang_hoa":self.e_hang.get().strip(),"sl":self._gf(self.e_sl),"don_gia":self._gf(self.e_dg),
           "thanh_tien":self._gf(self.e_tt),"ghi_chu":self.e_gc.get().strip()}
        if self.rid: self.db.update_bang_ke(self.rid,d)
        else: self.db.insert_bang_ke(d)
        if self.on_save: self.on_save()
        self.destroy()
