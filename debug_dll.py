import os
import sys

site_packages = [p for p in sys.path if "site-packages" in p][0]
for name in os.listdir(site_packages):
    if "fitz" in name.lower() or "pymupdf" in name.lower():
        full_path = os.path.join(site_packages, name)
        print(name, "-> isdir:", os.path.isdir(full_path))
        if os.path.isdir(full_path):
            print("  Files:", os.listdir(full_path))
