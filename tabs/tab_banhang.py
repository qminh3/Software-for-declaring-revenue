"""Tab S2b - So doanh thu ban hang (Nhom 3)"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, datetime
from database import NHOM_NGANH, THUE_SUAT_VAT_LIST

def fmt(v):
    try: f=float(v or 0); return f"{f:,.0f}" if f else ""
    except: return str(v) if v else ""

class BanHangTab(tk.Frame):
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
        cols=("ngay","so_hdon","dien_giai","ten_khach","ten_hang","nhom","sl","don_gia","thanh_tien","thue_suat","vat_ra","tong_thu","tk_thu")
        hdrs=("Ngày","Số HĐ","Diễn giải","Khách hàng","Tên hàng","Nhóm","SL","Đơn giá","Thành tiền","% VAT","VAT đầu ra","Tổng thu","TK Thu")
        ws=(85,80,150,120,120,50,65,110,120,60,110,120,60)
        tf=tk.Frame(self,bg="#F8FAFC"); tf.pack(fill="both",expand=True,padx=8,pady=4)
        vsb=ttk.Scrollbar(tf,orient="vertical"); hsb=ttk.Scrollbar(tf,orient="horizontal")
        self.tree=ttk.Treeview(tf,columns=cols,show="headings",yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview); hsb.config(command=self.tree.xview)
        for c,h,w in zip(cols,hdrs,ws):
            self.tree.heading(c,text=h)
            self.tree.column(c,width=w,anchor="e" if c in ("sl","don_gia","thanh_tien","vat_ra","tong_thu") else "w")
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        tf.grid_rowconfigure(0,weight=1); tf.grid_columnconfigure(0,weight=1)
        self.tree.tag_configure("row",background="white"); self.tree.tag_configure("alt",background="#F1F5F9")
        self.tree.bind("<<TreeviewSelect>>",lambda e: setattr(self,"_sel",int(self.tree.selection()[0])) if self.tree.selection() else None)
        self.tree.bind("<Double-1>",lambda e: self._edit())

    def _load(self):
        y=int(self.var_y.get()); m=self.var_m.get()
        mf=1 if m=="Tất cả" else int(m); mt=12 if m=="Tất cả" else int(m)
        rows=self.db.get_ban_hang(y,mf,mt); self.tree.delete(*self.tree.get_children())
        tot_tt=tot_vat=tot_thu=0
        for i,r in enumerate(rows):
            ngay=r["ngay"]
            try: ngay=datetime.strptime(str(ngay),"%Y-%m-%d").strftime("%d/%m/%Y")
            except: pass
            tt=float(r["thanh_tien"] or 0); vat=float(r["vat_dau_ra"] or 0); thu=float(r["tong_thu"] or 0)
            tot_tt+=tt; tot_vat+=vat; tot_thu+=thu
            pct=f"{float(r['thue_suat'] or 0)*100:.0f}%"
            tag="alt" if i%2 else "row"
            nhom_txt=r["nhom_nganh"] or ""
            self.tree.insert("","end",iid=r["id"],tags=(tag,),values=(
                ngay,r["so_hdon"] or "",r["dien_giai"] or "",r["ten_khach"] or "",
                r["ten_hang"] or "",nhom_txt,fmt(r["sl"]),fmt(r["don_gia"]),
                fmt(tt),pct,fmt(vat),fmt(thu),r["tk_thu"] or ""))
        self.lbl.config(text=f"  {len(rows)} dòng  |  Doanh thu (chưa VAT): {fmt(tot_tt)} đ  |  VAT đầu ra: {fmt(tot_vat)} đ  |  Tổng thu (có VAT): {fmt(tot_thu)} đ")

    def _new(self): BanHangForm(self,self.db,None,self._load)
    def _edit(self):
        if not self._sel: messagebox.showwarning("","Chọn dòng!"); return
        BanHangForm(self,self.db,self._sel,self._load)
    def _del(self):
        if not self._sel: messagebox.showwarning("","Chọn dòng!"); return
        if messagebox.askyesno("Xóa","Xóa dòng này?"): self.db.delete_ban_hang(self._sel); self._sel=None; self._load()
    def _export(self):
        from excel_export_nhom3 import export_s2b
        y=int(self.var_y.get()); m=self.var_m.get()
        mf=1 if m=="Tất cả" else int(m); mt=12 if m=="Tất cả" else int(m)
        fp=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],initialfile=f"S2b_HKD_{y}.xlsx")
        if not fp: return
        try: export_s2b(self.db,y,mf,mt,fp); messagebox.showinfo("OK",f"Đã xuất:\n{fp}")
        except Exception as e: messagebox.showerror("Lỗi",str(e))


class BanHangForm(tk.Toplevel):
    def __init__(self,parent,db,rid=None,on_save=None):
        super().__init__(parent); self.db=db; self.rid=rid; self.on_save=on_save
        self.title("Thêm bán hàng" if not rid else "Sửa bán hàng")
        sw=self.winfo_screenwidth(); sh=self.winfo_screenheight()
        w,h=min(760,sw-40),min(560,sh-80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.configure(bg="#F8FAFC"); self.grab_set(); self._build()
        if rid: self._load_row(rid)

    def _build(self):
        tk.Label(self,text="💰  NHẬP DỮ LIỆU BÁN HÀNG (S2b-HKD)",
                 font=("Arial",13,"bold"),bg="#2563EB",fg="white").pack(fill="x",ipady=8)
        canvas=tk.Canvas(self,bg="#F8FAFC",highlightthickness=0)
        vsb=ttk.Scrollbar(self,orient="vertical",command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set); vsb.pack(side="right",fill="y"); canvas.pack(fill="both",expand=True)
        body=tk.Frame(canvas,bg="#F8FAFC"); wid=canvas.create_window((0,0),window=body,anchor="nw")
        body.bind("<Configure>",lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",lambda e: canvas.itemconfig(wid,width=e.width))
        canvas.bind_all("<MouseWheel>",lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

        hang_names=self.db.get_hang_hoa_names(); quy=self.db.get_quy_names()
        BG="#FFFFFF"

        def r(p,t,row,col,**g): tk.Label(p,text=t,font=("Arial",9),bg=p["bg"],fg="#374151",anchor="e").grid(row=row,column=col,sticky="e",padx=(4,4),pady=4,**g)
        def e(p,row,col,w=14,**g): ent=tk.Entry(p,font=("Arial",10),width=w,relief="solid",bd=1,bg="white"); ent.grid(row=row,column=col,sticky="ew",padx=(0,8),pady=4,**g); return ent

        sec1=tk.LabelFrame(body,text=" ℹ️  Thông tin chung ",font=("Arial",10,"bold"),fg="#0F172A",bg="#FFFFFF",padx=12,pady=8)
        sec1.pack(fill="x",padx=12,pady=(10,5))
        r(sec1,"Ngày:*",0,0); self.e_ngay=e(sec1,0,1,w=13); self.e_ngay.insert(0,date.today().strftime("%d/%m/%Y"))
        r(sec1,"Số HĐ:",0,2); self.e_hdon=e(sec1,0,3,w=14)
        r(sec1,"Diễn giải:",1,0); self.e_dg=tk.Entry(sec1,font=("Arial",10),relief="solid",bd=1,bg="white"); self.e_dg.grid(row=1,column=1,columnspan=3,sticky="ew",padx=(0,8),pady=4)
        r(sec1,"Khách hàng:",2,0); self.e_khach=tk.Entry(sec1,font=("Arial",10),relief="solid",bd=1,bg="white"); self.e_khach.grid(row=2,column=1,columnspan=3,sticky="ew",padx=(0,8),pady=4)
        r(sec1,"Tên hàng:",3,0); self.cb_hang=ttk.Combobox(sec1,values=hang_names+["(khác)"],font=("Arial",10),width=22); self.cb_hang.grid(row=3,column=1,sticky="ew",padx=(0,8),pady=4); self.cb_hang.bind("<<ComboboxSelected>>",self._on_hang)
        r(sec1,"ĐVT:",3,2); self.e_dvt=e(sec1,3,3,w=10)
        r(sec1,"Nhóm ngành:",4,0)
        nhom_vals=[f"{n[0]} - {n[1]}" for n in NHOM_NGANH]
        self.cb_nhom=ttk.Combobox(sec1,values=nhom_vals,font=("Arial",10),state="readonly"); self.cb_nhom.grid(row=4,column=1,columnspan=3,sticky="ew",padx=(0,8),pady=4); self.cb_nhom.current(0)
        sec1.columnconfigure(1,weight=2); sec1.columnconfigure(3,weight=1)

        sec2=tk.LabelFrame(body,text=" 🧮  Số lượng, đơn giá & Thuế VAT ",font=("Arial",10,"bold"),fg="#0F172A",bg="#FFFFFF",padx=12,pady=8)
        sec2.pack(fill="x",padx=12,pady=5)
        r(sec2,"SL:",0,0); self.e_sl=e(sec2,0,1,w=12); self.e_sl.bind("<FocusOut>",self._calc); self.e_sl.bind("<Return>",self._calc)
        r(sec2,"Đơn giá (chưa VAT):",0,2); self.e_dg2=e(sec2,0,3,w=14); self.e_dg2.bind("<FocusOut>",self._calc); self.e_dg2.bind("<Return>",self._calc)
        r(sec2,"Thành tiền:",1,0); self.e_tt=tk.Entry(sec2,font=("Arial",10),width=14,relief="solid",bd=1,bg="#EFF6FF"); self.e_tt.grid(row=1,column=1,sticky="ew",padx=(0,8),pady=4); self.e_tt.bind("<FocusOut>",self._calc_from_tt)
        r(sec2,"Thuế suất VAT:",1,2)
        self.cb_vat=ttk.Combobox(sec2,values=[f"{v[0]}" for v in THUE_SUAT_VAT_LIST],font=("Arial",10),width=8,state="readonly"); self.cb_vat.grid(row=1,column=3,sticky="w",padx=(0,8),pady=4); self.cb_vat.current(1)
        r(sec2,"VAT đầu ra:",2,0); self.e_vat=tk.Entry(sec2,font=("Arial",10),width=14,relief="solid",bd=1,bg="#EFF6FF"); self.e_vat.grid(row=2,column=1,sticky="ew",padx=(0,8),pady=4)
        r(sec2,"Tổng thu (có VAT):",2,2); self.e_tong=tk.Entry(sec2,font=("Arial",10),width=14,relief="solid",bd=1,bg="#EFF6FF"); self.e_tong.grid(row=2,column=3,sticky="ew",padx=(0,8),pady=4)
        tk.Button(sec2,text="⚡ Tính",font=("Arial",10,"bold"),bg="#16A34A",fg="white",relief="flat",padx=12,pady=6,command=self._calc).grid(row=0,column=4,rowspan=3,padx=12)
        r(sec2,"TK Thu:",3,0)
        self.cb_tk=ttk.Combobox(sec2,values=quy,font=("Arial",10),width=12); self.cb_tk.grid(row=3,column=1,sticky="w",padx=(0,8),pady=4)
        if quy: self.cb_tk.current(0)
        sec2.columnconfigure(1,weight=1); sec2.columnconfigure(3,weight=1)

        bf=tk.Frame(body,bg="#F8FAFC"); bf.pack(fill="x",padx=12,pady=12)
        tk.Button(bf,text="💾  LƯU",font=("Arial",11,"bold"),bg="#16A34A",fg="white",relief="flat",padx=24,pady=10,command=self._save).pack(side="left",padx=6)
        tk.Button(bf,text="❌  Hủy",font=("Arial",11),bg="#64748B",fg="white",relief="flat",padx=20,pady=10,command=self.destroy).pack(side="left",padx=6)
        tk.Label(bf,text="💡 Nhập SL × Đơn giá → ⚡ Tính để tự tính VAT",font=("Arial",9),bg="#F8FAFC",fg="#6B7280").pack(side="right",padx=12)

    def _on_hang(self,e=None):
        ten=self.cb_hang.get()
        for h in self.db.get_hang_hoa():
            if h["ten_hang"]==ten: self.e_dvt.delete(0,"end"); self.e_dvt.insert(0,h["dvt"] or ""); break

    def _gf(self,w):
        try: return float(w.get().replace(",","").replace("%","").strip() or 0)
        except: return 0.0

    def _calc(self,e=None):
        sl=self._gf(self.e_sl); dg=self._gf(self.e_dg2); tt=sl*dg
        pct_str=self.cb_vat.get().replace("%","")
        try: pct=float(pct_str)/100
        except: pct=0.10
        vat=tt*pct; tong=tt+vat
        self.e_tt.delete(0,"end"); self.e_vat.delete(0,"end"); self.e_tong.delete(0,"end")
        if tt: self.e_tt.insert(0,f"{tt:,.0f}")
        if vat: self.e_vat.insert(0,f"{vat:,.0f}")
        if tong: self.e_tong.insert(0,f"{tong:,.0f}")

    def _calc_from_tt(self,e=None):
        tt=self._gf(self.e_tt)
        pct_str=self.cb_vat.get().replace("%","")
        try: pct=float(pct_str)/100
        except: pct=0.10
        vat=tt*pct; tong=tt+vat
        self.e_vat.delete(0,"end"); self.e_tong.delete(0,"end")
        if vat: self.e_vat.insert(0,f"{vat:,.0f}")
        if tong: self.e_tong.insert(0,f"{tong:,.0f}")

    def _load_row(self,rid):
        r=self.db.get_ban_hang_by_id(rid)
        if not r: return
        def se(w,v):
            w.delete(0,"end")
            if v is not None and str(v).strip() not in ("","0","0.0"): w.insert(0,str(v))
        ngay=r["ngay"]
        try: ngay=datetime.strptime(str(ngay),"%Y-%m-%d").strftime("%d/%m/%Y")
        except: pass
        self.e_ngay.delete(0,"end"); self.e_ngay.insert(0,ngay or "")
        se(self.e_hdon,r["so_hdon"]); se(self.e_dg,r["dien_giai"]); se(self.e_khach,r["ten_khach"])
        self.cb_hang.set(r["ten_hang"] or ""); se(self.e_dvt,r["dvt"])
        nhom=r["nhom_nganh"] or "1"
        try: self.cb_nhom.current(int(nhom)-1)
        except: self.cb_nhom.current(0)
        se(self.e_sl,fmt(r["sl"])); se(self.e_dg2,fmt(r["don_gia"]))
        se(self.e_tt,fmt(r["thanh_tien"])); se(self.e_vat,fmt(r["vat_dau_ra"])); se(self.e_tong,fmt(r["tong_thu"]))
        pct=float(r["thue_suat"] or 0.10)*100
        self.cb_vat.set(f"{pct:.0f}%")
        if r["tk_thu"] in self.db.get_quy_names(): self.cb_tk.set(r["tk_thu"])

    def _save(self):
        ns=self.e_ngay.get().strip()
        try: ngay=datetime.strptime(ns,"%d/%m/%Y").strftime("%Y-%m-%d")
        except: messagebox.showerror("Lỗi","Ngày sai định dạng DD/MM/YYYY!"); return
        pct_str=self.cb_vat.get().replace("%","")
        try: pct=float(pct_str)/100
        except: pct=0.10
        nhom_idx=self.cb_nhom.current(); nhom=str(nhom_idx+1) if nhom_idx>=0 else "1"
        d={"ngay":ngay,"so_hdon":self.e_hdon.get().strip(),"dien_giai":self.e_dg.get().strip(),
           "ten_khach":self.e_khach.get().strip(),"ten_hang":self.cb_hang.get().strip(),
           "dvt":self.e_dvt.get().strip(),"nhom_nganh":nhom,
           "sl":self._gf(self.e_sl),"don_gia":self._gf(self.e_dg2),
           "thanh_tien":self._gf(self.e_tt),"thue_suat":pct,
           "vat_dau_ra":self._gf(self.e_vat),"tong_thu":self._gf(self.e_tong),
           "tk_thu":self.cb_tk.get().strip()}
        if self.rid: self.db.update_ban_hang(self.rid,d)
        else: self.db.insert_ban_hang(d)
        if self.on_save: self.on_save()
        self.destroy()
