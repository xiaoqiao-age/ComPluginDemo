"""
VB6 COM Plugin Deploy Tool
从注册表读取 CLSID, 更新 manifest (含 CSharpComPlugin 依赖), 复制 DLL 到宿主输出目录
"""
import winreg, shutil, os, sys, glob

ACCESS_32 = winreg.KEY_READ | winreg.KEY_WOW64_32KEY
ACCESS_64 = winreg.KEY_READ

PROGID = "VB6ComPlugin.StringProcessor"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SLN_DIR = os.path.dirname(SCRIPT_DIR)
VB6_SRC = os.path.join(SLN_DIR, "src", "VB6ComPlugin")
HOST_OUT = os.path.join(SLN_DIR, "src", "ComPluginHost", "bin", "Debug", "net8.0-windows", "VB6ComPlugin")
CSHARP_OUT = os.path.join(SLN_DIR, "src", "ComPluginHost", "bin", "Debug", "net8.0-windows", "CSharpComPlugin")

# Manifest with CSharpComPlugin dependency for Reg-Free COM
MANIFEST_TEMPLATE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity type="win32" name="VB6ComPlugin" version="1.0.0.0" />
  <file name="VB6ComPlugin.dll">
    <comClass
      clsid="{clsid}"
      threadingModel="Apartment"
      progid="VB6ComPlugin.StringProcessor"
      description="StringProcessor Plugin - VB6" />
  </file>
  <!-- VB6 can CreateObject("ComPluginDemo.Calculator") via this dependency -->
  <dependency>
    <dependentAssembly>
      <assemblyIdentity type="win32" name="CSharpComPlugin" version="1.0.0.0" />
    </dependentAssembly>
  </dependency>
</assembly>
'''

def read_clsid():
    """Read CLSID from registry (try both 32-bit and 64-bit views)"""
    for access, label in [(ACCESS_32, "WOW64-32"), (ACCESS_64, "64-bit")]:
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{PROGID}\\CLSID", access=access)
            clsid = winreg.QueryValueEx(key, "")[0]
            winreg.CloseKey(key)
            print(f"[OK] CLSID ({label}): {clsid}")
            return clsid, access
        except FileNotFoundError:
            continue
    return None, None

def read_dll_path(clsid, access):
    """Read InprocServer32 DLL path from registry"""
    for subkey in [f"CLSID\\{clsid}\\InprocServer32", f"Wow6432Node\\CLSID\\{clsid}\\InprocServer32"]:
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey, access=access)
            path = winreg.QueryValueEx(key, "")[0]
            winreg.CloseKey(key)
            if os.path.exists(path):
                return path
        except FileNotFoundError:
            continue
    # Fallback: search in VB6 source directory
    fallback = os.path.join(VB6_SRC, "VB6ComPlugin.dll")
    if os.path.exists(fallback):
        return fallback
    return None

def copy_csharp_plugin():
    """Copy CSharpComPlugin files into VB6ComPlugin/CSharpComPlugin/ subdirectory
    This enables SxS dependency resolution: VB6ComPlugin.manifest → CSharpComPlugin"""
    dest_dir = os.path.join(HOST_OUT, "CSharpComPlugin")
    os.makedirs(dest_dir, exist_ok=True)

    if not os.path.isdir(CSHARP_OUT):
        print(f"[WARN] CSharpComPlugin output not found: {CSHARP_OUT}")
        print("       Run 'dotnet build' first")
        return False

    count = 0
    for f in os.listdir(CSHARP_OUT):
        src = os.path.join(CSHARP_OUT, f)
        dst = os.path.join(dest_dir, f)
        if os.path.isfile(src):
            shutil.copy2(src, dst)
            count += 1
    print(f"[OK] Copied {count} CSharpComPlugin files to: {dest_dir}")
    return True

def main():
    print("=" * 50)
    print("  VB6 COM Plugin Deploy Tool")
    print("  (with CSharpComPlugin Reg-Free dependency)")
    print("=" * 50)
    print()

    # 1. Read CLSID
    clsid, access = read_clsid()
    if not clsid:
        print("[ERROR] Registry has no entry for: " + PROGID)
        print()
        print("Please compile in VB6 IDE first:")
        print("  1. Open VB6ComPlugin.vbp")
        print("  2. File -> Make VB6ComPlugin.dll")
        print("  3. Run this script again")
        return 1

    # 2. Find DLL
    dll_path = read_dll_path(clsid, access)
    if not dll_path:
        print("[ERROR] Cannot find VB6ComPlugin.dll")
        return 1
    print(f"[OK] DLL: {dll_path}")

    # 3. Write manifest (with CSharpComPlugin dependency) to both locations
    manifest = MANIFEST_TEMPLATE.format(clsid=clsid)
    os.makedirs(HOST_OUT, exist_ok=True)

    for dest_dir in [HOST_OUT, VB6_SRC]:
        manifest_path = os.path.join(dest_dir, "VB6ComPlugin.manifest")
        with open(manifest_path, "w", encoding="utf-8") as f:
            f.write(manifest)
        print(f"[OK] Manifest: {manifest_path}")

    # 4. Copy VB6 DLL to output
    dest_dll = os.path.join(HOST_OUT, "VB6ComPlugin.dll")
    shutil.copy2(dll_path, dest_dll)
    print(f"[OK] DLL copied to: {dest_dll}")

    # 5. Copy CSharpComPlugin into VB6ComPlugin subdirectory (for SxS dependency)
    print()
    copy_csharp_plugin()

    # 6. Verify
    print()
    print("Output directory structure:")
    for root, dirs, files in os.walk(HOST_OUT):
        level = root.replace(HOST_OUT, "").count(os.sep)
        indent = "  " * level
        subdir = os.path.basename(root)
        if level > 0:
            print(f"  {indent}{subdir}/")
        for f in sorted(files):
            size = os.path.getsize(os.path.join(root, f))
            print(f"  {indent}  {f}  ({size:,} bytes)")

    print()
    print("Done! Reg-Free COM structure:")
    print("  VB6ComPlugin.manifest  → declares VB6 COM class")
    print("                         → depends on CSharpComPlugin")
    print("  CSharpComPlugin/       → C# COM files (resolved by SxS)")
    print()
    print("Now run ComPluginHost.exe")
    return 0

if __name__ == "__main__":
    sys.exit(main())
