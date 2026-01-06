"""
Build script để tạo file exe từ excel_tool.py
Chạy: python build_exe.py
"""
import subprocess
import sys
import shutil
import io
from pathlib import Path

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def build_exe():
    """Build executable using PyInstaller"""
    
    print("=" * 60)
    print("BUILD TOOL CHAM CONG CONG TAC")
    print("=" * 60)
    
    # Check PyInstaller
    try:
        import PyInstaller
        print(f"✓ PyInstaller installed (version {PyInstaller.__version__})")
    except ImportError:
        print("⚠ PyInstaller not installed.")
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed")
    
    # Check dependencies
    print("\nChecking dependencies...")
    try:
        import openpyxl
        print(f"✓ openpyxl version {openpyxl.__version__}")
    except ImportError:
        print("⚠ openpyxl not installed. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    
    try:
        import PIL
        print(f"✓ Pillow version {PIL.__version__}")
    except ImportError:
        print("⚠ Pillow not installed. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    
    # Clean old build/dist
    script_dir = Path(__file__).parent
    build_dir = script_dir / "build"
    dist_dir = script_dir / "dist"
    
    if build_dir.exists():
        print(f"\nCleaning old build directory...")
        shutil.rmtree(build_dir, ignore_errors=True)
    
    if dist_dir.exists():
        print(f"Cleaning old dist directory...")
        shutil.rmtree(dist_dir, ignore_errors=True)
    
    # Build with spec file
    spec_file = script_dir / "ChamCongCongTac.spec"
    
    # Use python -m PyInstaller instead of pyinstaller command
    cmd = [sys.executable, "-m", "PyInstaller", "--clean", str(spec_file)]
    
    print(f"\nBuilding executable...")
    print(f"Spec file: {spec_file}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='utf-8', errors='replace')
        print("\n" + "=" * 60)
        print("✓ BUILD SUCCESS!")
        print("=" * 60)
        
        exe_path = dist_dir / "ChamCongCongTac.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\nExe file: {exe_path}")
            print(f"Size: {size_mb:.2f} MB")
            print(f"\nYou can copy this exe to run on other machines.")
            print(f"Compatible: Windows 7, 8, 10, 11")
        else:
            print("⚠ Exe file not found!")
            
    except subprocess.CalledProcessError as e:
        print("\n" + "=" * 60)
        print("✗ BUILD ERROR")
        print("=" * 60)
        print(e.stderr)
        sys.exit(1)

if __name__ == "__main__":
    build_exe()
