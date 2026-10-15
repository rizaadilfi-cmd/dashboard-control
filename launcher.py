"""
PT NIFCO INDONESIA - MILKRUN AUTOMATIC LAUNCHER
Cross-Platform Launcher (Windows & Linux)
- Auto detects Operating System (Windows / Linux / macOS)
- Auto checks & installs dependencies from requirements.txt
- Auto runs backend (uvicorn backend:app --host 127.0.0.1 --port 8000 --reload)
- Auto waits for server ready and opens browser to http://127.0.0.1:8000
"""

import sys
import os
import glob
import platform
import subprocess
import time
import urllib.request
import webbrowser
import signal

# Ensure correct base directory even when running as PyInstaller frozen exe
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)

REQ_FILE = os.path.join(BASE_DIR, "requirements.txt")
BACKEND_FILE = os.path.join(BASE_DIR, "backend.py")
INSTALLER_DIR = os.path.join(BASE_DIR, "Instal Python")
TARGET_URL = "http://127.0.0.1:8000"
HEALTH_URL = f"{TARGET_URL}/api/health"

REQUIRED_MODULES = ["fastapi", "uvicorn", "openpyxl", "pydantic"]

def get_current_os():
    sys_name = platform.system()
    machine = platform.machine()
    if sys_name == "Windows":
        return "Windows", f"{platform.platform()} ({machine})"
    elif sys_name == "Linux":
        return "Linux", f"{platform.platform()} ({machine})"
    elif sys_name == "Darwin":
        return "macOS", f"{platform.platform()} ({machine})"
    return sys_name, f"{platform.platform()} ({machine})"

def is_valid_python(candidate):
    """Test if candidate is a working Python 3 executable."""
    try:
        res = subprocess.run([candidate, "-c", "import sys; sys.exit(0 if sys.version_info[0] >= 3 else 1)"],
                             capture_output=True, text=True, timeout=5)
        return res.returncode == 0
    except Exception:
        return False

def find_python():
    """Detect an active Python 3 interpreter on the machine."""
    # 1. If running as regular script, sys.executable is the active python
    if not getattr(sys, "frozen", False):
        if is_valid_python(sys.executable):
            return sys.executable

    # 2. Test common CLI commands in PATH
    for cmd in ["python", "py", "python3"]:
        if is_valid_python(cmd):
            return cmd

    # 3. On Windows, check standard installation directories
    if platform.system() == "Windows":
        candidates = []
        local_app = os.environ.get("LOCALAPPDATA", "")
        prog_files = os.environ.get("ProgramFiles", "")
        prog_files_x86 = os.environ.get("ProgramFiles(x86)", "")

        if local_app:
            candidates.append(os.path.join(local_app, "Python", "bin", "python.exe"))
            candidates.extend(glob.glob(os.path.join(local_app, "Python", "pythoncore-*", "python.exe")))
            candidates.extend(glob.glob(os.path.join(local_app, "Programs", "Python", "Python*", "python.exe")))
        if prog_files:
            candidates.extend(glob.glob(os.path.join(prog_files, "Python*", "python.exe")))
        if prog_files_x86:
            candidates.extend(glob.glob(os.path.join(prog_files_x86, "Python*", "python.exe")))
        candidates.extend(glob.glob(r"C:\Python*\python.exe"))

        for c in candidates:
            if os.path.isfile(c) and is_valid_python(c):
                return c

    # 4. Check if bundled installer exists in 'Instal Python'
    if os.path.isdir(INSTALLER_DIR):
        installers = glob.glob(os.path.join(INSTALLER_DIR, "*.exe")) + glob.glob(os.path.join(INSTALLER_DIR, "*.msix"))
        if installers:
            installer_path = installers[0]
            print("\n" + "!" * 60)
            print(" [PERHATIAN] Python belum terdeteksi di sistem komputer ini.")
            print(f" [*] Ditemukan file installer di folder 'Instal Python':")
            print(f"     {os.path.basename(installer_path)}")
            print("=" * 60)
            choice = input(" Apakah Anda ingin menjalankan installer sekarang? (Y/n): ").strip().lower()
            if choice in ("", "y", "ya", "yes"):
                print(f" [*] Menjalankan installer: {installer_path}...")
                try:
                    if installer_path.endswith(".exe"):
                        subprocess.run([installer_path], check=True)
                    else:
                        os.system(f'start "" "{installer_path}"')
                    print(" [*] Selesai instalasi. Memeriksa kembali interpreter Python...")
                    time.sleep(2)
                    for cmd in ["python", "py", "python3"]:
                        if is_valid_python(cmd):
                            return cmd
                except Exception as e:
                    print(f" [!] Gagal menjalankan installer: {e}")

    return None

def check_and_install_dependencies(py_cmd):
    print("\n" + "=" * 60)
    print(" [1/3] MEMERIKSA PAKET & DEPENDENCIES...")
    print("=" * 60)

    missing = []
    for mod in REQUIRED_MODULES:
        try:
            res = subprocess.run([py_cmd, "-c", f"import {mod}"], capture_output=True, timeout=5)
            if res.returncode != 0:
                missing.append(mod)
        except Exception:
            missing.append(mod)

    if missing:
        print(f" [-] Modul berikut butuh diinstal: {', '.join(missing)}")
        if os.path.exists(REQ_FILE):
            print(f" [*] Menginstall paket dari: {REQ_FILE}...")
            cmd = [py_cmd, "-m", "pip", "install", "-r", REQ_FILE]
            print(f" [RUN] {' '.join(cmd)}")
            res = subprocess.run(cmd)
            if res.returncode != 0:
                print(" [!] Gagal menginstall dependensi via pip.")
                print("     Pastikan koneksi internet aktif untuk download paket pertama kali.")
            else:
                print(" [OK] Seluruh paket berhasil diinstall!")
        else:
            print(f" [*] Menginstall modul langsung via pip: {' '.join(missing)}...")
            cmd = [py_cmd, "-m", "pip", "install"] + missing
            subprocess.run(cmd)
    else:
        print(" [OK] Semua paket dependensi sudah lengkap dan siap digunakan.")

def is_server_running(timeout=1.5):
    """Check if Milkrun server is responding on target port."""
    try:
        req = urllib.request.Request(HEALTH_URL, headers={"User-Agent": "MilkrunLauncher/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status == 200:
                return True
    except Exception:
        pass
    return False

def wait_for_server(timeout=35):
    start = time.time()
    while time.time() - start < timeout:
        if is_server_running(timeout=1.2):
            return True
        time.sleep(0.8)
    return False

def open_in_browser(url):
    """Safely open dashboard URL in user's default browser."""
    opened = False
    try:
        opened = webbrowser.open(url)
    except Exception:
        pass

    if not opened:
        if platform.system() == "Windows":
            os.system(f'start "" "{url}"')
        elif platform.system() == "Linux":
            os.system(f'xdg-open "{url}" >/dev/null 2>&1 &')
        elif platform.system() == "Darwin":
            os.system(f'open "{url}"')

def main():
    os_name, os_details = get_current_os()
    print("=" * 65)
    print("         PT NIFCO INDONESIA - MILKRUN DASHBOARD LAUNCHER")
    print("=" * 65)
    print(f" [*] Sistem Operasi Terdeteksi : {os_name} ({os_details})")
    print(f" [*] Lokasi Folder Kerja       : {BASE_DIR}")

    # Check if server is already running on port 8000
    if is_server_running():
        print("\n [OK] Server Milkrun SUDAH AKTIF di http://127.0.0.1:8000!")
        print(" [*] Langsung membuka dashboard di browser...")
        open_in_browser(TARGET_URL)
        print("\n Dashboard berhasil dibuka. Launcher selesai.")
        time.sleep(2)
        sys.exit(0)

    # Detect Python
    py_cmd = find_python()
    if not py_cmd:
        print("\n" + "!" * 65)
        print(" [ERROR] Python 3 tidak ditemukan di komputer ini!")
        print(" Silakan install Python 3 melalui file di folder 'Instal Python'")
        print(" atau download di: https://www.python.org/downloads/")
        print(" (Penting: centang 'Add Python to PATH' saat instalasi)")
        print("!" * 65)
        input("\n Tekan ENTER untuk keluar...")
        sys.exit(1)

    print(f" [*] Interpreter Python        : {py_cmd}")

    if not os.path.exists(BACKEND_FILE):
        print(f"\n [ERROR] File backend.py tidak ditemukan di folder:\n {BASE_DIR}")
        input("\n Tekan ENTER untuk keluar...")
        sys.exit(1)

    # 1. Check and install dependencies
    check_and_install_dependencies(py_cmd)

    # 2. Start backend server
    print("\n" + "=" * 65)
    print(" [2/3] MENJALANKAN BACKEND MILKRUN...")
    print("=" * 65)
    print(f" [*] Perintah: uvicorn backend:app --host 127.0.0.1 --port 8000 --reload")
    
    cmd = [py_cmd, "-m", "uvicorn", "backend:app", "--host", "127.0.0.1", "--port", "8000", "--reload"]

    proc = subprocess.Popen(
        cmd,
        cwd=BASE_DIR,
        stdout=None,
        stderr=None
    )

    # 3. Wait and auto open browser
    print("\n" + "=" * 65)
    print(" [3/3] MENUNGGU SERVER SIAP & MEMBUKA BROWSER...")
    print("=" * 65)
    print(f" [*] Memeriksa status server di {TARGET_URL}...")

    ready = wait_for_server()
    if ready:
        print(f"\n INFO: Uvicorn running on http://127.0.0.1:8000")
        print(f" [OK] Server BERHASIL AKTIF!")
        print(f" [*] Membuka browser ke alamat: {TARGET_URL}")
        open_in_browser(TARGET_URL)
    else:
        print(f"\n [!] Server memerlukan waktu lebih lama untuk start.")
        print(f" [*] Mencoba tetap membuka browser ke {TARGET_URL}...")
        open_in_browser(TARGET_URL)

    print("\n" + "=" * 65)
    print(" MILKRUN SERVER SEDANG BERJALAN DENGAN SUKSES!")
    print(f" Alamat Dashboard: {TARGET_URL}")
    print(" (Jangan tutup jendela terminal ini agar server tetap aktif)")
    print(" Tekan CTRL + C untuk menghentikan server.")
    print("=" * 65 + "\n")

    def handle_shutdown(signum=None, frame=None):
        print("\n [*] Menghentikan server Milkrun...")
        try:
            if platform.system() == "Windows":
                subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
            else:
                proc.terminate()
                try:
                    proc.wait(timeout=3)
                except Exception:
                    proc.kill()
        except Exception:
            pass
        print(" [OK] Server berhasil dimatikan dengan aman. Sampai jumpa!")
        sys.exit(0)

    signal.signal(signal.SIGINT, handle_shutdown)
    if hasattr(signal, "SIGBREAK"):
        signal.signal(signal.SIGBREAK, handle_shutdown)

    try:
        proc.wait()
    except KeyboardInterrupt:
        handle_shutdown()

if __name__ == "__main__":
    main()
