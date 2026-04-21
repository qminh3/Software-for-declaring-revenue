"""
SQLite database - Ho Kinh Doanh NHOM 3
Doanh thu >= 3 ty/nam - TT152/2025/TT-BTC
- Thue GTGT: phuong phap KHAU TRU (dau ra - dau vao)
- Thue TNCN: 17% tren thu nhap tinh thue (doanh thu - chi phi hop ly)
- So sach: S2b, S2c, S2d, S2e
"""
import sqlite3
import os
from datetime import datetime, date

import sys as _sys
def _get_db_path():
    if hasattr(_sys, "_MEIPASS"):
        return os.path.join(os.path.dirname(_sys.executable), "hkd_nhom3.db")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "hkd_nhom3.db")
DB_PATH = _get_db_path()

DEFAULT_CONFIG = {
    "ten_hkd": "Công ty TNHH Quang Minh BP",
    "mst":     "3801345604",
    "dia_chi": "Đồng Nai",
    "nhom":    "3",
}

NHOM_NGANH = [
    ("1", "Phân phối, cung cấp hàng hóa (VAT 1%)",        0.01),
    ("2", "Dịch vụ, xây dựng không bao thầu VL (VAT 5%)", 0.05),
    ("3", "Sản xuất, xây dựng có bao thầu VL (VAT 3%)",   0.03),
    ("4", "Kinh doanh khác (VAT 2%)",                      0.02),
]

THUE_SUAT_VAT_LIST = [
    ("5%",  0.05),
    ("10%", 0.10),
    ("0%",  0.00),
]

LOAI_CHI_PHI = [
    "Nhập mua hàng hóa",
    "Chi lương nhân viên",
    "Chi thuê mặt bằng",
    "Chi vận chuyển",
    "Chi điện nước",
    "Chi khấu hao TSCĐ",
    "Chi phí quảng cáo",
    "Chi phí quản lý",
    "Chi lãi vay",
    "Chi phí khác",
]

THUE_TNCN_SUAT = 0.17


class Database:
    def __init__(self):
        self.conn = sqlite3.connect(DB_PATH)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self._init_tables()
        self._seed_demo()

    def _init_tables(self):
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS config(key TEXT PRIMARY KEY, value TEXT);

        CREATE TABLE IF NOT EXISTS ban_hang(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay DATE NOT NULL,
            so_hdon TEXT, dien_giai TEXT, ten_khach TEXT, ten_hang TEXT,
            nhom_nganh TEXT DEFAULT '1',
            sl REAL DEFAULT 0, don_gia REAL DEFAULT 0,
            thanh_tien REAL DEFAULT 0,
            thue_suat REAL DEFAULT 0.10,
            vat_dau_ra REAL DEFAULT 0,
            tong_thu REAL DEFAULT 0,
            tk_thu TEXT DEFAULT 'TM'
        );

        CREATE TABLE IF NOT EXISTS chi_phi(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay DATE NOT NULL,
            so_ct TEXT, dien_giai TEXT, ten_khach TEXT,
            loai_cp TEXT DEFAULT 'Nhập mua hàng hóa',
            so_tien REAL DEFAULT 0,
            vat_dau_vao REAL DEFAULT 0,
            tk_chi TEXT DEFAULT 'TM'
        );

        CREATE TABLE IF NOT EXISTS hang_hoa(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ten_hang TEXT UNIQUE NOT NULL,
            dvt TEXT, ton_dau_nam REAL DEFAULT 0, gia_von REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS quy_tk(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tai_khoan TEXT UNIQUE NOT NULL, du_dau_nam REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS bang_ke_thu_mua(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay DATE NOT NULL,
            ten_nguoi_ban TEXT, dia_chi TEXT, so_cccd TEXT, sdt TEXT,
            ten_hang_hoa TEXT, sl REAL DEFAULT 0, don_gia REAL DEFAULT 0,
            thanh_tien REAL DEFAULT 0, ghi_chu TEXT
        );
        """)
        self.conn.commit()

    def _seed_demo(self):
        cur = self.conn.cursor()
        for k, v in DEFAULT_CONFIG.items():
            cur.execute("INSERT OR IGNORE INTO config(key,value) VALUES(?,?)", (k,v))
        hangs = [("Xi măng","Tấn",200,100000),("Cát","m3",150,500000),
                 ("Đá","m3",300,280000),("Thép","Tấn",80,14000),
                 ("Gạch","Viên",5000,1000),("Thạch cao","Tấn",50,140000),
                 ("Sơn","Thùng",30,330000)]
        for h in hangs:
            cur.execute("INSERT OR IGNORE INTO hang_hoa(ten_hang,dvt,ton_dau_nam,gia_von) VALUES(?,?,?,?)", h)
        quy = [("TM",120000000),("VCB",280000000),("VTB",50000000)]
        for q in quy:
            cur.execute("INSERT OR IGNORE INTO quy_tk(tai_khoan,du_dau_nam) VALUES(?,?)", q)
        cur.execute("SELECT COUNT(*) FROM ban_hang")
        if cur.fetchone()[0] == 0:
            ban = [
                ("2026-01-05","HD001","Xuất bán hàng","Công ty A","Xi măng","1",200,130000,0.10,"VCB"),
                ("2026-01-07","HD002","Xuất bán hàng","Công ty B","Cát","1",80,680000,0.10,"VCB"),
                ("2026-01-10","HD003","Xuất bán hàng","Công ty C","Đá","1",120,350000,0.10,"TM"),
                ("2026-01-12","HD004","Dịch vụ thi công","Khách D","Thi công","2",1,50000000,0.10,"VCB"),
                ("2026-01-15","HD005","Xuất bán hàng","Công ty E","Thép","1",60,18000,0.10,"VCB"),
                ("2026-01-18","HD006","Xuất bán hàng","Khách F","Gạch","1",2000,1600,0.10,"TM"),
                ("2026-01-20","HD007","Xuất bán hàng","Công ty G","Thạch cao","1",40,180000,0.10,"VCB"),
                ("2026-01-22","HD008","Xuất bán hàng","Khách H","Sơn","1",25,420000,0.10,"TM"),
                ("2026-01-25","HD009","Dịch vụ thi công","Khách I","Thi công","2",1,80000000,0.10,"VCB"),
                ("2026-01-28","HD010","Xuất bán hàng","Công ty J","Xi măng","1",150,132000,0.10,"VCB"),
                ("2026-02-03","HD011","Xuất bán hàng","Công ty A","Cát","1",60,685000,0.10,"VCB"),
                ("2026-02-06","HD012","Xuất bán hàng","Công ty K","Đá","1",100,355000,0.10,"TM"),
                ("2026-02-10","HD013","Dịch vụ thi công","Khách L","Thi công","2",1,65000000,0.10,"VCB"),
                ("2026-02-15","HD014","Xuất bán hàng","Công ty M","Thép","1",50,18500,0.10,"VCB"),
                ("2026-02-20","HD015","Xuất bán hàng","Khách N","Xi măng","1",180,135000,0.10,"VTB"),
            ]
            for b in ban:
                tt = b[6]*b[7]; vat = tt*b[8]; tong = tt+vat
                cur.execute("""INSERT INTO ban_hang(ngay,so_hdon,dien_giai,ten_khach,ten_hang,
                    nhom_nganh,sl,don_gia,thanh_tien,thue_suat,vat_dau_ra,tong_thu,tk_thu)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (b[0],b[1],b[2],b[3],b[4],b[5],b[6],b[7],tt,b[8],vat,tong,b[9]))
        cur.execute("SELECT COUNT(*) FROM chi_phi")
        if cur.fetchone()[0] == 0:
            cps = [
                ("2026-01-03","NK001","Nhập mua Xi măng","Công ty S","Nhập mua hàng hóa",20000000,2000000,"VCB"),
                ("2026-01-03","NK001","Nhập mua Cát","Công ty S","Nhập mua hàng hóa",40000000,4000000,"VCB"),
                ("2026-01-03","NK001","Nhập mua Đá","Công ty S","Nhập mua hàng hóa",33600000,3360000,"VCB"),
                ("2026-01-05","PC001","Lương nhân viên T1","","Chi lương nhân viên",15000000,0,"TM"),
                ("2026-01-05","PC002","Thuê mặt bằng T1","","Chi thuê mặt bằng",8000000,800000,"TM"),
                ("2026-01-08","PC003","Chi phí vận chuyển","","Chi vận chuyển",3000000,300000,"TM"),
                ("2026-01-15","NK002","Nhập mua Thép","Công ty T","Nhập mua hàng hóa",840000,84000,"VCB"),
                ("2026-01-20","PC004","Chi điện nước T1","","Chi điện nước",2000000,200000,"TM"),
                ("2026-02-05","PC005","Lương nhân viên T2","","Chi lương nhân viên",15000000,0,"TM"),
                ("2026-02-08","NK003","Nhập mua hàng hóa","Công ty U","Nhập mua hàng hóa",35000000,3500000,"VCB"),
                ("2026-02-10","PC006","Khấu hao TSCĐ T2","","Chi khấu hao TSCĐ",5000000,0,"TM"),
                ("2026-02-20","PC007","Chi điện nước T2","","Chi điện nước",2200000,220000,"TM"),
            ]
            for c in cps:
                cur.execute("""INSERT INTO chi_phi(ngay,so_ct,dien_giai,ten_khach,loai_cp,so_tien,vat_dau_vao,tk_chi)
                    VALUES(?,?,?,?,?,?,?,?)""", c)
        self.conn.commit()

    # Config
    def get_config(self):
        return {r["key"]: r["value"] for r in self.conn.execute("SELECT key,value FROM config").fetchall()}
    def save_config(self, cfg):
        for k,v in cfg.items(): self.conn.execute("INSERT OR REPLACE INTO config(key,value) VALUES(?,?)",(k,v))
        self.conn.commit()

    # Bán hàng
    def get_ban_hang(self, year=None, mf=None, mt=None):
        q="SELECT * FROM ban_hang WHERE 1=1"; p=[]
        if year: q+=" AND strftime('%Y',ngay)=?"; p.append(str(year))
        if mf:   q+=" AND CAST(strftime('%m',ngay) AS INT)>=?"; p.append(int(mf))
        if mt:   q+=" AND CAST(strftime('%m',ngay) AS INT)<=?"; p.append(int(mt))
        return [dict(r) for r in self.conn.execute(q+" ORDER BY ngay,id", p).fetchall()]
    def get_ban_hang_by_id(self, rid):
        r=self.conn.execute("SELECT * FROM ban_hang WHERE id=?",(rid,)).fetchone(); return dict(r) if r else None
    def insert_ban_hang(self, d):
        cols=["ngay","so_hdon","dien_giai","ten_khach","ten_hang","nhom_nganh","sl","don_gia","thanh_tien","thue_suat","vat_dau_ra","tong_thu","tk_thu"]
        str_cols={"ngay","so_hdon","dien_giai","ten_khach","ten_hang","nhom_nganh","tk_thu"}
        vals=[d.get(c, "" if c in str_cols else 0) for c in cols]
        self.conn.execute(f"INSERT INTO ban_hang({','.join(cols)}) VALUES({','.join(['?']*len(cols))})", vals)
        self.conn.commit()
    def update_ban_hang(self, rid, d):
        cols=["ngay","so_hdon","dien_giai","ten_khach","ten_hang","nhom_nganh","sl","don_gia","thanh_tien","thue_suat","vat_dau_ra","tong_thu","tk_thu"]
        str_cols={"ngay","so_hdon","dien_giai","ten_khach","ten_hang","nhom_nganh","tk_thu"}
        vals=[d.get(c,"" if c in str_cols else 0) for c in cols]+[rid]
        self.conn.execute(f"UPDATE ban_hang SET {', '.join([f'{c}=?' for c in cols])} WHERE id=?", vals)
        self.conn.commit()
    def delete_ban_hang(self, rid):
        self.conn.execute("DELETE FROM ban_hang WHERE id=?",(rid,)); self.conn.commit()

    # Chi phí
    def get_chi_phi(self, year=None, mf=None, mt=None):
        q="SELECT * FROM chi_phi WHERE 1=1"; p=[]
        if year: q+=" AND strftime('%Y',ngay)=?"; p.append(str(year))
        if mf:   q+=" AND CAST(strftime('%m',ngay) AS INT)>=?"; p.append(int(mf))
        if mt:   q+=" AND CAST(strftime('%m',ngay) AS INT)<=?"; p.append(int(mt))
        return [dict(r) for r in self.conn.execute(q+" ORDER BY ngay,id", p).fetchall()]
    def get_chi_phi_by_id(self, rid):
        r=self.conn.execute("SELECT * FROM chi_phi WHERE id=?",(rid,)).fetchone(); return dict(r) if r else None
    def insert_chi_phi(self, d):
        cols=["ngay","so_ct","dien_giai","ten_khach","loai_cp","so_tien","vat_dau_vao","tk_chi"]
        str_cols={"ngay","so_ct","dien_giai","ten_khach","loai_cp","tk_chi"}
        vals=[d.get(c,"" if c in str_cols else 0) for c in cols]
        self.conn.execute(f"INSERT INTO chi_phi({','.join(cols)}) VALUES({','.join(['?']*len(cols))})", vals)
        self.conn.commit()
    def update_chi_phi(self, rid, d):
        cols=["ngay","so_ct","dien_giai","ten_khach","loai_cp","so_tien","vat_dau_vao","tk_chi"]
        str_cols={"ngay","so_ct","dien_giai","ten_khach","loai_cp","tk_chi"}
        vals=[d.get(c,"" if c in str_cols else 0) for c in cols]+[rid]
        self.conn.execute(f"UPDATE chi_phi SET {', '.join([f'{c}=?' for c in cols])} WHERE id=?", vals)
        self.conn.commit()
    def delete_chi_phi(self, rid):
        self.conn.execute("DELETE FROM chi_phi WHERE id=?",(rid,)); self.conn.commit()

    # Hàng hóa
    def get_hang_hoa(self): return [dict(r) for r in self.conn.execute("SELECT * FROM hang_hoa ORDER BY ten_hang").fetchall()]
    def get_hang_hoa_names(self): return [r["ten_hang"] for r in self.conn.execute("SELECT ten_hang FROM hang_hoa ORDER BY ten_hang").fetchall()]
    def upsert_hang_hoa(self, ten, dvt, ton, gia_von):
        self.conn.execute("INSERT INTO hang_hoa(ten_hang,dvt,ton_dau_nam,gia_von) VALUES(?,?,?,?) ON CONFLICT(ten_hang) DO UPDATE SET dvt=excluded.dvt,ton_dau_nam=excluded.ton_dau_nam,gia_von=excluded.gia_von",(ten,dvt,ton,gia_von)); self.conn.commit()
    def delete_hang_hoa(self, ten): self.conn.execute("DELETE FROM hang_hoa WHERE ten_hang=?",(ten,)); self.conn.commit()

    # Quỹ
    def get_quy_tk(self): return [dict(r) for r in self.conn.execute("SELECT * FROM quy_tk ORDER BY tai_khoan").fetchall()]
    def get_quy_names(self): return [r["tai_khoan"] for r in self.conn.execute("SELECT tai_khoan FROM quy_tk ORDER BY tai_khoan").fetchall()]
    def upsert_quy_tk(self, tk, du): self.conn.execute("INSERT INTO quy_tk(tai_khoan,du_dau_nam) VALUES(?,?) ON CONFLICT(tai_khoan) DO UPDATE SET du_dau_nam=excluded.du_dau_nam",(tk,du)); self.conn.commit()
    def delete_quy_tk(self, tk): self.conn.execute("DELETE FROM quy_tk WHERE tai_khoan=?",(tk,)); self.conn.commit()

    # Aggregations
    def get_vat_khau_tru(self, year, mf, mt):
        bans=self.get_ban_hang(year,mf,mt); cps=self.get_chi_phi(year,mf,mt)
        vat_ra=sum(float(r["vat_dau_ra"] or 0) for r in bans)
        vat_vao=sum(float(r["vat_dau_vao"] or 0) for r in cps)
        return {"vat_dau_ra":vat_ra,"vat_dau_vao":vat_vao,
                "vat_phai_nop":max(vat_ra-vat_vao,0),"vat_con_khau_tru":max(vat_vao-vat_ra,0)}

    def get_s2c_summary(self, year, mf, mt):
        bans=self.get_ban_hang(year,mf,mt); cps=self.get_chi_phi(year,mf,mt)
        tong_dt=sum(float(r["thanh_tien"] or 0) for r in bans)
        tong_cp=sum(float(r["so_tien"] or 0) for r in cps)
        thu_nhap=tong_dt-tong_cp
        return {"tong_doanh_thu":tong_dt,"tong_chi_phi":tong_cp,
                "thu_nhap_tinh_thue":thu_nhap,
                "thue_tncn":max(thu_nhap,0)*THUE_TNCN_SUAT,
                "thue_suat_tncn":THUE_TNCN_SUAT}

    def get_nxt_data(self, year, mf, mt):
        hangs=self.get_hang_hoa(); result=[]
        for h in hangs:
            tn=h["ten_hang"]
            xuat=self.conn.execute("SELECT COALESCE(SUM(sl),0) as sl,COALESCE(SUM(thanh_tien),0) as tien FROM ban_hang WHERE ten_hang=? AND strftime('%Y',ngay)=? AND CAST(strftime('%m',ngay) AS INT)>=? AND CAST(strftime('%m',ngay) AS INT)<=?",(tn,str(year),int(mf),int(mt))).fetchone()
            xuat_all=self.conn.execute("SELECT COALESCE(SUM(sl),0) as sl FROM ban_hang WHERE ten_hang=? AND strftime('%Y',ngay)=?",(tn,str(year))).fetchone()
            result.append({"ten_hang":tn,"dvt":h["dvt"],"ton_dau_nam":float(h["ton_dau_nam"] or 0),"ton_hien_tai":float(h["ton_dau_nam"] or 0)-float(xuat_all["sl"] or 0),"sl_xuat":float(xuat["sl"] or 0),"doanh_thu":float(xuat["tien"] or 0)})
        return result

    def get_s2e_data(self, year, mf, mt):
        result=[]
        for q in self.get_quy_tk():
            tk=q["tai_khoan"]; ddn=float(q["du_dau_nam"] or 0)
            thu_t=self.conn.execute("SELECT COALESCE(SUM(tong_thu),0) as s FROM ban_hang WHERE tk_thu=? AND strftime('%Y',ngay)=? AND CAST(strftime('%m',ngay) AS INT)>=? AND CAST(strftime('%m',ngay) AS INT)<=?",(tk,str(year),int(mf),int(mt))).fetchone()["s"]
            chi_t=self.conn.execute("SELECT COALESCE(SUM(so_tien),0) as s FROM chi_phi WHERE tk_chi=? AND strftime('%Y',ngay)=? AND CAST(strftime('%m',ngay) AS INT)>=? AND CAST(strftime('%m',ngay) AS INT)<=?",(tk,str(year),int(mf),int(mt))).fetchone()["s"]
            thu_truoc=self.conn.execute("SELECT COALESCE(SUM(tong_thu),0) as s FROM ban_hang WHERE tk_thu=? AND strftime('%Y',ngay)=? AND CAST(strftime('%m',ngay) AS INT)<?",(tk,str(year),int(mf))).fetchone()["s"]
            chi_truoc=self.conn.execute("SELECT COALESCE(SUM(so_tien),0) as s FROM chi_phi WHERE tk_chi=? AND strftime('%Y',ngay)=? AND CAST(strftime('%m',ngay) AS INT)<?",(tk,str(year),int(mf))).fetchone()["s"]
            thu_all=self.conn.execute("SELECT COALESCE(SUM(tong_thu),0) as s FROM ban_hang WHERE tk_thu=? AND strftime('%Y',ngay)=?",(tk,str(year))).fetchone()["s"]
            chi_all=self.conn.execute("SELECT COALESCE(SUM(so_tien),0) as s FROM chi_phi WHERE tk_chi=? AND strftime('%Y',ngay)=?",(tk,str(year))).fetchone()["s"]
            ddk=ddn+thu_truoc-chi_truoc
            result.append({"tai_khoan":tk,"du_dau_nam":ddn,"du_hien_tai":ddn+thu_all-chi_all,"du_dau_ky":ddk,"tong_thu":thu_t,"tong_chi":chi_t,"du_cuoi_ky":ddk+thu_t-chi_t})
        return result

    # Bảng kê thu mua
    def get_bang_ke(self, year=None, mf=None, mt=None):
        q="SELECT * FROM bang_ke_thu_mua WHERE 1=1"; p=[]
        if year: q+=" AND strftime('%Y',ngay)=?"; p.append(str(year))
        if mf:   q+=" AND CAST(strftime('%m',ngay) AS INT)>=?"; p.append(int(mf))
        if mt:   q+=" AND CAST(strftime('%m',ngay) AS INT)<=?"; p.append(int(mt))
        return [dict(r) for r in self.conn.execute(q+" ORDER BY ngay,id", p).fetchall()]
    def get_bang_ke_by_id(self, rid):
        r=self.conn.execute("SELECT * FROM bang_ke_thu_mua WHERE id=?",(rid,)).fetchone(); return dict(r) if r else None
    def insert_bang_ke(self, d):
        cols=["ngay","ten_nguoi_ban","dia_chi","so_cccd","sdt","ten_hang_hoa","sl","don_gia","thanh_tien","ghi_chu"]
        str_cols={"ngay","ten_nguoi_ban","dia_chi","so_cccd","sdt","ten_hang_hoa","ghi_chu"}
        vals=[d.get(c,"" if c in str_cols else 0) for c in cols]
        self.conn.execute(f"INSERT INTO bang_ke_thu_mua({','.join(cols)}) VALUES({','.join(['?']*len(cols))})", vals)
        self.conn.commit()
    def update_bang_ke(self, rid, d):
        cols=["ngay","ten_nguoi_ban","dia_chi","so_cccd","sdt","ten_hang_hoa","sl","don_gia","thanh_tien","ghi_chu"]
        str_cols={"ngay","ten_nguoi_ban","dia_chi","so_cccd","sdt","ten_hang_hoa","ghi_chu"}
        vals=[d.get(c,"" if c in str_cols else 0) for c in cols]+[rid]
        self.conn.execute(f"UPDATE bang_ke_thu_mua SET {', '.join([f'{c}=?' for c in cols])} WHERE id=?", vals)
        self.conn.commit()
    def delete_bang_ke(self, rid):
        self.conn.execute("DELETE FROM bang_ke_thu_mua WHERE id=?",(rid,)); self.conn.commit()

    def close(self): self.conn.close()
