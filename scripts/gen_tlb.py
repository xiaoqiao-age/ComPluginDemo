"""
Generate CSharpComPlugin.tlb for VB6 early binding reference.
Approach: MIDL via Visual Studio Developer Command Prompt environment.
"""
import os
import sys
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SLN_DIR = os.path.dirname(SCRIPT_DIR)
TLB_PATH = os.path.join(SLN_DIR, "src", "VB6ComPlugin", "CSharpComPlugin.tlb")


def find_vs_path():
    vswhere = r"C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
    if not os.path.exists(vswhere):
        return None
    result = subprocess.run(
        [vswhere, "-latest", "-property", "installationPath"],
        capture_output=True, text=True
    )
    path = result.stdout.strip()
    return path if path else None


def find_sdk_include():
    base = r"C:\Program Files (x86)\Windows Kits\10\Include"
    if not os.path.isdir(base):
        return None
    versions = sorted(os.listdir(base))
    return os.path.join(base, versions[-1]) if versions else None


def main():
    print("=" * 50)
    print("  CSharpComPlugin TLB Generator")
    print("=" * 50)
    print()

    vs_path = find_vs_path()
    if not vs_path:
        print("[ERROR] Visual Studio not found")
        return 1
    print(f"[OK] VS: {vs_path}")

    sdk_inc = find_sdk_include()
    if not sdk_inc:
        print("[ERROR] Windows SDK includes not found")
        return 1
    print(f"[OK] SDK: {sdk_inc}")

    idl_path = os.path.join(SLN_DIR, "src", "CSharpComPlugin", "CSharpComPlugin.idl")
    if not os.path.exists(idl_path):
        print(f"[ERROR] IDL file not found: {idl_path}")
        return 1

    # Write a temp batch script that sets up VS environment then runs MIDL
    bat_path = os.path.join(SCRIPT_DIR, "_midl_tmp.bat")
    with open(bat_path, "w", encoding="gbk") as f:
        f.write(f'@echo off\n')
        f.write(f'call "{vs_path}\\Common7\\Tools\\VsDevCmd.bat" -arch=x86 >nul 2>&1\n')
        f.write(f'midl /tlb "{TLB_PATH}" /I "{sdk_inc}\\um" /I "{sdk_inc}\\shared" "{idl_path}"\n')
        f.write(f'exit /b %ERRORLEVEL%\n')

    print(f"\n[RUN] midl /tlb CSharpComPlugin.tlb CSharpComPlugin.idl")
    result = subprocess.run(
        ["cmd.exe", "/c", bat_path],
        capture_output=True, text=True, encoding="gbk", errors="replace"
    )

    # Clean up temp bat
    try:
        os.remove(bat_path)
    except:
        pass

    print(result.stdout)
    if result.stderr:
        print(result.stderr)

    if os.path.exists(TLB_PATH):
        size = os.path.getsize(TLB_PATH)
        print(f"\n[OK] TLB generated: {TLB_PATH}")
        print(f"     Size: {size:,} bytes")
        return 0
    else:
        print(f"\n[ERROR] TLB was not generated")
        return 1


if __name__ == "__main__":
    sys.exit(main())
