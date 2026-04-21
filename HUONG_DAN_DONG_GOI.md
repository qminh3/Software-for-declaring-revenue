# HƯỚNG DẪN ĐÓNG GÓI - HKD NHÓM 3

## Cách dùng nhanh nhất (1 bước)

1. Giải nén ZIP vào thư mục bất kỳ (ví dụ `C:\HKD_Nhom3\`)
2. **Bấm đúp vào `DongGoiEXE.bat`**
3. Đợi 2–5 phút → File `dist\HKD_Nhom3.exe` được tạo ra
4. Copy file `.exe` sang bất kỳ máy Windows nào → **chạy không cần Python!**

---

## Yêu cầu để đóng gói

- **Python 3.8+** đã cài trên máy build
- **Kết nối Internet** để tải PyInstaller (chỉ lần đầu, ~10MB)
- Dung lượng trống ~500MB (tạm thời khi build)
- File EXE sau khi build: khoảng **30–60MB**

---

## Cấu trúc sau khi build thành công

```
HoKinhDoanh_Nhom3/
├── dist/
│   └── HKD_Nhom3.exe       ← File chạy được, copy đi được!
├── build_temp/              ← Xóa được sau khi build
├── DongGoiEXE.bat           ← Script tự động build
├── HKD_Nhom3.spec           ← Config PyInstaller nâng cao
└── ...
```

---

## Xử lý lỗi thường gặp

| Lỗi | Cách xử lý |
|-----|-----------|
| "No module named..." | `pip install pyinstaller pillow openpyxl --upgrade` |
| "Access denied" | Click phải → **Run as administrator** |
| Antivirus chặn | Tắt tạm Windows Defender khi build |
| Build chậm/treo | Đợi thêm, lần đầu thường mất 5–10 phút |

---

## Build thủ công (nếu BAT không chạy)

```bat
cd C:\đường_dẫn\HoKinhDoanh_Nhom3
python -m PyInstaller HKD_Nhom3.spec
```

---

## Tạo shortcut Desktop cho file EXE

1. Click phải vào `dist\HKD_Nhom3.exe`
2. **Send to → Desktop (create shortcut)**
3. Click phải shortcut → **Properties → Change Icon** → chọn `icon.ico`
