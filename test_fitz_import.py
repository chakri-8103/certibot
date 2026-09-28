import os
import sys

site_packages = [p for p in sys.path if "site-packages" in p][0]
pymupdf_dir = os.path.join(site_packages, "pymupdf")
if os.path.exists(pymupdf_dir):
    os.environ["PATH"] = pymupdf_dir + os.pathsep + os.environ.get("PATH", "")
    if hasattr(os, "add_dll_directory"):
        try:
            os.add_dll_directory(pymupdf_dir)
        except Exception:
            pass

try:
    import fitz
    print("SUCCESS: fitz imported! Version:", fitz.__version__)
except Exception as e:
    import ctypes
    print("fitz error:", e)
    pyd_path = os.path.join(pymupdf_dir, "_extra.pyd")
    try:
        ctypes.WinDLL(pyd_path)
    except Exception as e2:
        print("ctypes WinDLL error:", e2)
