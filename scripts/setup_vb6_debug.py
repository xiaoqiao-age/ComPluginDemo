"""
Setup VB6 debug environment.

Copies CSharpComPlugin build output to src/VB6ComPlugin/CSharpComPlugin/
so that VB6 IDE debug mode can find the manifest for Reg-Free COM.

Usage:
    python scripts/setup_vb6_debug.py

After running:
    1. Open VB6ComPlugin.vbp in VB6 IDE
    2. Set breakpoints in StringProcessor.cls / frmComExplorer.frm
    3. Press F5 (Run -> Start) to start debugging
    4. Start ComPluginHost.exe
    5. Select StringProcessor, type "explore", click Execute
    6. VB6 breakpoints will be hit
    7. In frmComExplorer, click "Create via Manifest" to test Reg-Free COM
"""
import os
import sys
import shutil
import glob

# Project root
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Source: CSharpComPlugin build output
CSHARP_BUILD_DIR = os.path.join(PROJECT_ROOT, "src", "CSharpComPlugin", "bin", "Debug", "net8.0")

# Destination: VB6 project subdirectory
VB6_CSHARP_DIR = os.path.join(PROJECT_ROOT, "src", "VB6ComPlugin", "CSharpComPlugin")

# Source manifest
CSHARP_MANIFEST = os.path.join(PROJECT_ROOT, "src", "CSharpComPlugin", "CSharpComPlugin.manifest")

# Files to copy from build output
REQUIRED_FILES = [
    "CSharpComPlugin.dll",
    "CSharpComPlugin.comhost.dll",
    "CSharpComPlugin.deps.json",
    "CSharpComPlugin.runtimeconfig.json",
]


def main():
    print("=" * 60)
    print("VB6 Debug Environment Setup")
    print("=" * 60)

    # Check CSharpComPlugin build
    if not os.path.isdir(CSHARP_BUILD_DIR):
        print(f"\n[ERROR] CSharpComPlugin not built yet.")
        print(f"  Expected: {CSHARP_BUILD_DIR}")
        print(f"\n  Run: dotnet build src/CSharpComPlugin/CSharpComPlugin.csproj")
        sys.exit(1)

    # Create destination
    os.makedirs(VB6_CSHARP_DIR, exist_ok=True)
    print(f"\nTarget: {VB6_CSHARP_DIR}")

    # Copy manifest
    if os.path.exists(CSHARP_MANIFEST):
        dest = os.path.join(VB6_CSHARP_DIR, "CSharpComPlugin.manifest")
        shutil.copy2(CSHARP_MANIFEST, dest)
        print(f"  [OK] CSharpComPlugin.manifest")
    else:
        print(f"  [WARN] Manifest not found: {CSHARP_MANIFEST}")

    # Copy build output files
    copied = 0
    for filename in REQUIRED_FILES:
        src = os.path.join(CSHARP_BUILD_DIR, filename)
        if os.path.exists(src):
            dest = os.path.join(VB6_CSHARP_DIR, filename)
            shutil.copy2(src, dest)
            size = os.path.getsize(dest)
            print(f"  [OK] {filename} ({size:,} bytes)")
            copied += 1
        else:
            print(f"  [SKIP] {filename} (not found)")

    # Also copy .NET runtime files that comhost needs
    # These are in the same directory as the build output
    runtime_files = glob.glob(os.path.join(CSHARP_BUILD_DIR, "*.dll"))
    extra = 0
    for src in runtime_files:
        filename = os.path.basename(src)
        if filename.startswith("CSharpComPlugin"):
            continue  # Already handled
        dest = os.path.join(VB6_CSHARP_DIR, filename)
        if not os.path.exists(dest) or os.path.getmtime(src) > os.path.getmtime(dest):
            shutil.copy2(src, dest)
            extra += 1

    print(f"\n  Copied {copied} core files + {extra} runtime files")
    print(f"\n{'=' * 60}")
    print("VB6 Hot Debugging Steps:")
    print("=" * 60)
    print("""
  1. Open VB6 IDE -> File -> Open Project -> VB6ComPlugin.vbp
  2. Set breakpoints in:
     - StringProcessor.cls (e.g., Execute function)
     - frmComExplorer.frm (e.g., btnCreateViaManifest_Click)
  3. Press F5 (Run -> Start) in VB6 IDE
     -> VB6 IDE temporarily registers the COM class in registry
  4. Start ComPluginHost.exe (C# host)
     -> Host loads StringProcessor via registry fallback
     -> Host passes Calculator COM object to VB6
  5. In C# host:
     - Select "StringProcessor" plugin
     - Type "explore" in input box
     - Click "Execute"
  6. VB6 breakpoints will be hit!
     - frmComExplorer opens showing Calculator info
  7. In frmComExplorer:
     - Click "Create via Manifest" to test Reg-Free COM
       (uses modActCtx Win32 Activation Context API)
     - Type math expression, click Execute to calculate
""")


if __name__ == "__main__":
    main()
