"""
Tao shortcut Desktop cho ung dung Quan Ly Ho Kinh Doanh
Chay file nay 1 lan de cai shortcut, sau do dung shortcut mo app
"""
import os
import sys
import subprocess
import tempfile

APP_DIR  = os.path.dirname(os.path.abspath(__file__))
APP_PY   = os.path.join(APP_DIR, "app.py")
ICON     = os.path.join(APP_DIR, "icon.ico")
NAME     = "HKD Nhom 3"          # ASCII – tranh loi font tren Desktop
NAME_VN  = "Qu\u1ea3n L\xfd H\u1ed9 Kinh Doanh"   # UTF-8 dep (hien trong Properties)


def get_desktop():
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        return desktop
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def create_shortcut(lnk_path):
    """Dung pythonw.exe de chay khong hien console khi click shortcut"""
    # Tim pythonw.exe cung thu muc voi python.exe
    python_exe = sys.executable
    pythonw = python_exe.replace("python.exe", "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = python_exe  # fallback

    # Ghi VBS bang UTF-16 LE (Windows default cho VBScript)
    vbs_lines = [
        'Set oWS = WScript.CreateObject("WScript.Shell")',
        'Set oLink = oWS.CreateShortcut("%s")' % lnk_path.replace("\\", "\\\\"),
        'oLink.TargetPath = "%s"' % pythonw.replace("\\", "\\\\"),
        'oLink.Arguments = """%s"""' % APP_PY.replace("\\", "\\\\"),
        'oLink.WorkingDirectory = "%s"' % APP_DIR.replace("\\", "\\\\"),
        'oLink.IconLocation = "%s"' % ICON.replace("\\", "\\\\"),
        'oLink.Description = "Quan Ly Ho Kinh Doanh"',
        'oLink.WindowStyle = 1',
        'oLink.Save',
    ]
    vbs_content = "\n".join(vbs_lines)

    vbs_path = os.path.join(tempfile.gettempdir(), "make_shortcut_hkd.vbs")
    # Ghi UTF-16 LE voi BOM – Windows VBScript doc tot nhat voi encoding nay
    with open(vbs_path, "w", encoding="utf-8-sig") as f:
        f.write(vbs_content)

    result = subprocess.run(
        ["cscript", "//nologo", vbs_path],
        capture_output=True, text=True
    )
    try:
        os.remove(vbs_path)
    except Exception:
        pass
    return os.path.exists(lnk_path)


def create_shortcut_ps(lnk_path):
    """Fallback: dung PowerShell"""
    python_exe = sys.executable
    pythonw = python_exe.replace("python.exe", "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = python_exe

    # Dung here-string trong PowerShell de tranh van de quote
    ps = (
        "$ws = New-Object -ComObject WScript.Shell\n"
        "$s = $ws.CreateShortcut('" + lnk_path.replace("'", "''") + "')\n"
        "$s.TargetPath = '" + pythonw.replace("'", "''") + "'\n"
        "$s.Arguments = '\"" + APP_PY.replace("'", "''") + "\"'\n"
        "$s.WorkingDirectory = '" + APP_DIR.replace("'", "''") + "'\n"
        "$s.IconLocation = '" + ICON.replace("'", "''") + "'\n"
        "$s.Description = 'Quan Ly Ho Kinh Doanh'\n"
        "$s.WindowStyle = 1\n"
        "$s.Save()\n"
        "Write-Host 'DONE'\n"
    )
    ps_path = os.path.join(tempfile.gettempdir(), "make_shortcut_hkd.ps1")
    with open(ps_path, "w", encoding="utf-8-sig") as f:
        f.write(ps)
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path],
        capture_output=True, text=True
    )
    try:
        os.remove(ps_path)
    except Exception:
        pass
    return os.path.exists(lnk_path)


def main():
    print("=" * 52)
    print("  TAO SHORTCUT DESKTOP")
    print("  Quan Ly Ho Kinh Doanh")
    print("=" * 52)
    print()

    if sys.platform != "win32":
        print("Script nay chi dung tren Windows!")
        input("Nhan Enter de thoat...")
        return

    desktop  = get_desktop()
    lnk_path = os.path.join(desktop, NAME + ".lnk")

    print(f"App folder : {APP_DIR}")
    print(f"Desktop    : {desktop}")
    print(f"Shortcut   : {lnk_path}")
    print()

    if os.path.exists(lnk_path):
        ans = input("Shortcut da ton tai. Ghi de? (Y/N): ").strip().upper()
        if ans != "Y":
            print("Huy bo.")
            input("Nhan Enter de thoat...")
            return
        try:
            os.remove(lnk_path)
        except Exception:
            pass

    print("Dang tao shortcut...")
    ok = False

    # Thu VBScript truoc
    try:
        ok = create_shortcut(lnk_path)
        if ok:
            print("OK (VBScript)")
    except Exception as e:
        print(f"  VBScript: {e}")

    # Fallback PowerShell
    if not ok:
        try:
            ok = create_shortcut_ps(lnk_path)
            if ok:
                print("OK (PowerShell)")
        except Exception as e:
            print(f"  PowerShell: {e}")

    print()
    if ok:
        print("HOAN THANH!")
        print(f"  Shortcut '{NAME}.lnk' da xuat hien tren Desktop.")
        print()
        print("  => Double-click vao icon de mo phan mem!")
    else:
        print("KHONG THE TAO SHORTCUT TU DONG.")
        print()
        print("HUONG DAN THU CONG:")
        print("  1. Click phai Desktop -> New -> Shortcut")
        print(f"  2. Nhap: pythonw \"{APP_PY}\"")
        print(f"  3. Dat ten: {NAME}")
        print("  4. Click phai shortcut -> Properties -> Change Icon")
        print(f"  5. Chon file: {ICON}")

    print()
    input("Nhan Enter de thoat...")


if __name__ == "__main__":
    main()
