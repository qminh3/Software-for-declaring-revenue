"""Excel export - HKD Nhom 3"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

NUM = '#,##0'
DATE = 'DD/MM/YYYY'

def _b():
    s = Side(style='thin', color='CBD5E1')
    return Border(left=s, right=s, top=s, bottom=s)

def _c(ws, row, col, value, bold=False, bg=None, fg="000000", align="left", num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", bold=bold, color=fg, size=10)
    if bg: c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    c.border = _b()
    if num_fmt: c.number_format = num_fmt
    return c

def _title_block(ws, cfg, title, period, max_col):
    for r in range(1, 6):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max_col)
    _c(ws,1,1, cfg.get("ten_hkd","").upper(), bold=True, bg="1E3A8A", fg="FFFFFF", align="center")
    ws.row_dimensions[1].height = 24
    _c(ws,2,1, f"Địa chỉ: {cfg.get('dia_chi','')}", bg="DBEAFE", align="center")
    _c(ws,3,1, f"MST: {cfg.get('mst','')}  |  Nhóm 3 (Doanh thu ≥ 3 tỷ/năm)", bg="DBEAFE", align="center")
    _c(ws,4,1, title, bold=True, bg="DBEAFE", fg="1E40AF", align="center")
    _c(ws,5,1, f"{period}  |  Xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}", bg="DBEAFE", align="center")
    ws.row_dimensions[6].height = 6


def export_s2b(db, year, mf, mt, filepath):
    rows = db.get_ban_hang(year, mf, mt)
    cfg  = db.get_config()
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S2b-DoanThu"
    ws.freeze_panes = "A9"
    headers = ["Ngày","Số HĐ","Diễn giải","Khách hàng","Tên hàng","Nhóm","SL","Đơn giá","Thành tiền","% VAT","VAT đầu ra","Tổng thu (có VAT)","TK Thu"]
    widths  = [12,10,22,18,18,8,8,14,16,7,16,18,8]
    _title_block(ws, cfg, "SỔ DOANH THU BÁN HÀNG HÓA DỊCH VỤ - MẪU S2b-HKD", f"Năm {year} | Tháng {mf}–{mt}", len(headers))
    for col,(h,w) in enumerate(zip(headers,widths),1):
        _c(ws,7,col,h,bold=True,bg="14532D",fg="FFFFFF",align="center")
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[7].height = 18
    tot_tt=tot_vat=tot_tong=0
    for i,r in enumerate(rows):
        row=8+i
        try: ngay=datetime.strptime(str(r["ngay"]),"%Y-%m-%d")
        except: ngay=r["ngay"]
        bg="DCFCE7" if i%2==0 else "F0FDF4"
        tt=float(r["thanh_tien"] or 0); vat=float(r["vat_dau_ra"] or 0); tong=float(r["tong_thu"] or 0)
        tot_tt+=tt; tot_vat+=vat; tot_tong+=tong
        _c(ws,row,1,ngay,bg=bg,num_fmt=DATE)
        _c(ws,row,2,r["so_hdon"] or "",bg=bg)
        _c(ws,row,3,r["dien_giai"] or "",bg=bg)
        _c(ws,row,4,r["ten_khach"] or "",bg=bg)
        _c(ws,row,5,r["ten_hang"] or "",bg=bg)
        _c(ws,row,6,r["nhom_nganh"] or "",bg=bg,align="center")
        _c(ws,row,7,float(r["sl"] or 0) or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,8,float(r["don_gia"] or 0) or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,9,tt or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,10,f"{float(r['thue_suat'] or 0)*100:.0f}%",bg=bg,align="center")
        _c(ws,row,11,vat or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,12,tong or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,13,r["tk_thu"] or "",bg=bg,align="center")
        ws.row_dimensions[row].height = 15
    tr=8+len(rows)
    ws.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=8)
    _c(ws,tr,1,"TỔNG CỘNG",bold=True,bg="14532D",fg="FFFFFF",align="center")
    for col,val in [(9,tot_tt),(11,tot_vat),(12,tot_tong)]:
        _c(ws,tr,col,val,bold=True,bg="14532D",fg="FFFFFF",num_fmt=NUM,align="right")
    for col in [10,13]: _c(ws,tr,col,None,bg="14532D")
    ws.row_dimensions[tr].height = 20
    wb.save(filepath)


def export_s2c(db, year, mf, mt, filepath):
    cps = db.get_chi_phi(year, mf, mt)
    cfg = db.get_config()
    s2c = db.get_s2c_summary(year, mf, mt)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S2c-ChiPhi"
    headers = ["Ngày","Số CT","Diễn giải","NCC","Loại Chi Phí","Số tiền (đ)","VAT đầu vào (đ)","TK Chi"]
    widths  = [12,10,28,18,22,16,18,8]
    _title_block(ws, cfg, "SỔ CHI TIẾT DOANH THU CHI PHÍ - MẪU S2c-HKD", f"Năm {year} | Tháng {mf}–{mt}", len(headers))
    for col,(h,w) in enumerate(zip(headers,widths),1):
        _c(ws,7,col,h,bold=True,bg="92400E",fg="FFFFFF",align="center")
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.row_dimensions[7].height=18
    tot_cp=tot_vat=0
    for i,r in enumerate(cps):
        row=8+i
        try: ngay=datetime.strptime(str(r["ngay"]),"%Y-%m-%d")
        except: ngay=r["ngay"]
        loai=r["loai_cp"] or ""
        bg="FEF9C3" if "hàng hóa" in loai.lower() else ("DBEAFE" if "lương" in loai.lower() else "FFF7ED")
        cp=float(r["so_tien"] or 0); vat=float(r["vat_dau_vao"] or 0)
        tot_cp+=cp; tot_vat+=vat
        _c(ws,row,1,ngay,bg=bg,num_fmt=DATE)
        _c(ws,row,2,r["so_ct"] or "",bg=bg)
        _c(ws,row,3,r["dien_giai"] or "",bg=bg)
        _c(ws,row,4,r["ten_khach"] or "",bg=bg)
        _c(ws,row,5,loai,bg=bg)
        _c(ws,row,6,cp or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,7,vat or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,8,r["tk_chi"] or "",bg=bg,align="center")
        ws.row_dimensions[row].height=15
    tr=8+len(cps)
    ws.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=5)
    _c(ws,tr,1,"TỔNG CỘNG",bold=True,bg="92400E",fg="FFFFFF",align="center")
    _c(ws,tr,6,tot_cp,bold=True,bg="92400E",fg="FFFFFF",num_fmt=NUM,align="right")
    _c(ws,tr,7,tot_vat,bold=True,bg="92400E",fg="FFFFFF",num_fmt=NUM,align="right")
    _c(ws,tr,8,None,bg="92400E"); ws.row_dimensions[tr].height=20
    # S2c summary rows
    gap=tr+2
    ws.merge_cells(start_row=gap,start_column=1,end_row=gap,end_column=8)
    _c(ws,gap,1,"TỔNG KẾT S2c - XÁC ĐỊNH THU NHẬP TÍNH THUẾ TNCN",bold=True,bg="1E3A8A",fg="FFFFFF",align="center")
    ws.row_dimensions[gap].height=20
    items=[("Tổng Doanh Thu (chưa VAT)",s2c["tong_doanh_thu"],"DBEAFE"),
           ("Tổng Chi Phí Hợp Lý",s2c["tong_chi_phi"],"FEF9C3"),
           ("Thu nhập tính thuế (DT - CP)",s2c["thu_nhap_tinh_thue"],"FDE68A"),
           (f"Thuế TNCN = {s2c['thue_suat_tncn']*100:.0f}% × Thu nhập",s2c["thue_tncn"],"1E3A8A")]
    for i,(lbl,val,bg) in enumerate(items):
        r2=gap+1+i
        ws.merge_cells(start_row=r2,start_column=1,end_row=r2,end_column=5)
        fg="FFFFFF" if bg=="1E3A8A" else "000000"
        _c(ws,r2,1,lbl,bold=bg=="1E3A8A",bg=bg,fg=fg,align="right")
        for col in range(6,9): _c(ws,r2,col,None,bg=bg)
        ws.cell(row=r2,column=6).value=val
        ws.cell(row=r2,column=6).font=Font(name="Arial",bold=True,color=fg,size=11)
        ws.cell(row=r2,column=6).fill=PatternFill("solid",fgColor=bg)
        ws.cell(row=r2,column=6).number_format=NUM
        ws.cell(row=r2,column=6).alignment=Alignment(horizontal="right",vertical="center")
        ws.row_dimensions[r2].height=18
    wb.save(filepath)


def export_baocao(db, year, mf, mt, filepath):
    """Xuất tong hop: VAT, TNCN, NXT, Quy"""
    cfg = db.get_config()
    wb = openpyxl.Workbook()

    # Sheet 1: VAT khau tru
    ws1=wb.active; ws1.title="VAT-KhauTru"
    vat=db.get_vat_khau_tru(year,mf,mt)
    _title_block(ws1,cfg,"THUẾ GTGT - PHƯƠNG PHÁP KHẤU TRỪ",f"Năm {year} | Tháng {mf}–{mt}",4)
    headers=["Khoản mục","Số tiền (đ)","Ghi chú",""]
    for col,(h,w) in enumerate(zip(headers,[36,18,30,10]),1):
        _c(ws1,7,col,h,bold=True,bg="7C3AED",fg="FFFFFF",align="center")
        ws1.column_dimensions[get_column_letter(col)].width=w
    items=[("VAT Đầu Ra (tổng VAT thu từ khách hàng)",vat["vat_dau_ra"],"Từ cột 'VAT đầu ra' tab Bán Hàng","DBEAFE"),
           ("VAT Đầu Vào (tổng VAT đã trả cho NCC)",vat["vat_dau_vao"],"Từ cột 'VAT đầu vào' tab Chi Phí","EDE9FE"),
           ("VAT Phải Nộp = Đầu Ra − Đầu Vào",vat["vat_phai_nop"],"Nộp vào ngân sách nhà nước","7C3AED"),
           ("VAT Còn Được Khấu Trừ Kỳ Sau",vat["vat_con_khau_tru"],"Khi đầu vào > đầu ra","DDD6FE")]
    for i,(lbl,val,note,bg) in enumerate(items):
        row=8+i; fg="FFFFFF" if bg=="7C3AED" else "000000"
        _c(ws1,row,1,lbl,bold=bg=="7C3AED",bg=bg,fg=fg)
        _c(ws1,row,2,val,bold=bg=="7C3AED",bg=bg,fg=fg,num_fmt=NUM,align="right")
        _c(ws1,row,3,note,bg=bg,fg=fg)
        _c(ws1,row,4,None,bg=bg)
        ws1.row_dimensions[row].height=20

    # Sheet 2: TNCN
    ws2=wb.create_sheet("TNCN-LoiNhuan")
    s2c=db.get_s2c_summary(year,mf,mt)
    _title_block(ws2,cfg,"THUẾ TNCN TRÊN THU NHẬP THỰC TẾ (S2c-HKD)",f"Năm {year} | Tháng {mf}–{mt}",4)
    for col,(h,w) in enumerate(zip(["Khoản mục","Số tiền (đ)","",""],[ 36,18,20,10]),1):
        _c(ws2,7,col,h,bold=True,bg="92400E",fg="FFFFFF",align="center")
        ws2.column_dimensions[get_column_letter(col)].width=w
    items2=[("(A) Tổng Doanh Thu (chưa VAT)",s2c["tong_doanh_thu"],"FFFBEB"),
            ("(B) Tổng Chi Phí Hợp Lý",s2c["tong_chi_phi"],"FEF3C7"),
            ("(C) Thu nhập tính thuế = A − B",s2c["thu_nhap_tinh_thue"],"FDE68A"),
            (f"(D) Thuế TNCN = C × {s2c['thue_suat_tncn']*100:.0f}%",s2c["thue_tncn"],"92400E")]
    for i,(lbl,val,bg) in enumerate(items2):
        row=8+i; fg="FFFFFF" if bg=="92400E" else "000000"
        _c(ws2,row,1,lbl,bold=bg=="92400E",bg=bg,fg=fg)
        _c(ws2,row,2,val,bold=bg=="92400E",bg=bg,fg=fg,num_fmt=NUM,align="right")
        for col in [3,4]: _c(ws2,row,col,None,bg=bg)
        ws2.row_dimensions[row].height=20

    # Sheet 3: Quy S2e
    ws3=wb.create_sheet("S2e-Quy")
    quy_data=db.get_s2e_data(year,mf,mt)
    _title_block(ws3,cfg,"SỔ CHI TIẾT TIỀN - MẪU S2e-HKD",f"Năm {year} | Tháng {mf}–{mt}",8)
    hdrs3=["STT","Tài khoản","Dư đầu năm","Dư hiện tại","Dư đầu kỳ","Tổng thu","Tổng chi","Dư cuối kỳ"]
    for col,(h,w) in enumerate(zip(hdrs3,[6,16,16,16,16,16,16,16]),1):
        _c(ws3,7,col,h,bold=True,bg="D97706",fg="FFFFFF",align="center")
        ws3.column_dimensions[get_column_letter(col)].width=w
    ws3.row_dimensions[7].height=18
    tot_thu=tot_chi=tot_dc=0
    for i,r in enumerate(quy_data):
        row=8+i; dc=float(r["du_cuoi_ky"] or 0)
        bg="FEF9C3" if r["tai_khoan"]=="TM" else "DBEAFE"
        if dc<0: bg="FEE2E2"
        _c(ws3,row,1,i+1,bg=bg,align="center")
        _c(ws3,row,2,r["tai_khoan"],bg=bg)
        for col,key in enumerate(["du_dau_nam","du_hien_tai","du_dau_ky","tong_thu","tong_chi","du_cuoi_ky"],3):
            _c(ws3,row,col,float(r[key] or 0),bg=bg,num_fmt=NUM,align="right")
        ws3.row_dimensions[row].height=16
        tot_thu+=float(r["tong_thu"] or 0); tot_chi+=float(r["tong_chi"] or 0); tot_dc+=dc
    tr=8+len(quy_data)
    ws3.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=2)
    _c(ws3,tr,1,"TỔNG CỘNG",bold=True,bg="D97706",fg="FFFFFF",align="center")
    for col in [3,4,5]: _c(ws3,tr,col,None,bg="D97706")
    for col,val in [(6,tot_thu),(7,tot_chi),(8,tot_dc)]:
        _c(ws3,tr,col,val,bold=True,bg="D97706",fg="FFFFFF",num_fmt=NUM,align="right")
    wb.save(filepath)


def export_bang_ke(db, year, mf, mt, filepath):
    rows = db.get_bang_ke(year, mf, mt)
    cfg = db.get_config()
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "02_TNDN-BangKe"
    ws.freeze_panes = "A9"
    headers = ["Ngày","Tên người bán","Địa chỉ","Số căn cước","SĐT (nếu có)","Tên hàng hóa, dịch vụ","SL","Đơn giá","Tổng thanh toán","Ghi chú"]
    widths  = [12, 22, 24, 16, 14, 25, 8, 14, 16, 18]
    _title_block(ws, cfg, "BẢNG KÊ THU MUA HÀNG HÓA, DỊCH VỤ KHÔNG CÓ HÓA ĐƠN", f"Mẫu số: 02/TNDN | Năm {year} | Tháng {mf}–{mt}", len(headers))
    for col,(h,w) in enumerate(zip(headers,widths),1):
        _c(ws,7,col,h,bold=True,bg="D97706",fg="FFFFFF",align="center")
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[7].height = 30
    tot_tt = 0
    for i,r in enumerate(rows):
        row=8+i
        try: ngay=datetime.strptime(str(r["ngay"]),"%Y-%m-%d")
        except: ngay=r["ngay"]
        bg="FEFCE8" if i%2==0 else "FFFFFF"
        tt=float(r["thanh_tien"] or 0)
        tot_tt+=tt
        _c(ws,row,1,ngay,bg=bg,num_fmt=DATE)
        _c(ws,row,2,r["ten_nguoi_ban"] or "",bg=bg)
        _c(ws,row,3,r["dia_chi"] or "",bg=bg)
        _c(ws,row,4,r["so_cccd"] or "",bg=bg,align="center")
        _c(ws,row,5,r["sdt"] or "",bg=bg,align="center")
        _c(ws,row,6,r["ten_hang_hoa"] or "",bg=bg)
        _c(ws,row,7,float(r["sl"] or 0) or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,8,float(r["don_gia"] or 0) or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,9,tt or None,bg=bg,num_fmt=NUM,align="right")
        _c(ws,row,10,r["ghi_chu"] or "",bg=bg)
        ws.row_dimensions[row].height = 16
    tr=8+len(rows)
    ws.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=8)
    _c(ws,tr,1,"TỔNG CỘNG",bold=True,bg="D97706",fg="FFFFFF",align="center")
    for col in range(2,9): _c(ws,tr,col,None,bg="D97706")
    _c(ws,tr,9,tot_tt,bold=True,bg="D97706",fg="FFFFFF",num_fmt=NUM,align="right")
    _c(ws,tr,10,None,bg="D97706")
    ws.row_dimensions[tr].height = 20
    
    # Ký tên
    sr = tr + 2
    ws.merge_cells(start_row=sr,start_column=2,end_row=sr,end_column=4)
    ws.cell(row=sr,column=2,value="Người lập bảng kê").font = Font(name="Arial",bold=True)
    ws.cell(row=sr,column=2).alignment = Alignment(horizontal="center")
    
    ws.merge_cells(start_row=sr,start_column=7,end_row=sr,end_column=9)
    ws.cell(row=sr,column=7,value="Người đại diện hoặc người được ủy quyền").font = Font(name="Arial",bold=True)
    ws.cell(row=sr,column=7).alignment = Alignment(horizontal="center")
    
    ws.merge_cells(start_row=sr+1,start_column=2,end_row=sr+1,end_column=4)
    ws.cell(row=sr+1,column=2,value="(Ký, ghi rõ họ tên)").font = Font(name="Arial",italic=True)
    ws.cell(row=sr+1,column=2).alignment = Alignment(horizontal="center")
    
    ws.merge_cells(start_row=sr+1,start_column=7,end_row=sr+1,end_column=9)
    ws.cell(row=sr+1,column=7,value="(Ký tên, đóng dấu)").font = Font(name="Arial",italic=True)
    ws.cell(row=sr+1,column=7).alignment = Alignment(horizontal="center")

    wb.save(filepath)
