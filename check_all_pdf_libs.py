import sys

libs = ["fitz", "pymupdf", "pypdf", "pdfplumber", "reportlab", "pikepdf", "pdf2image"]

for lib in libs:
    try:
        mod = __import__(lib)
        print(f"[{lib}] SUCCESS! Version: {getattr(mod, '__version__', 'unknown')}")
    except Exception as e:
        print(f"[{lib}] FAILED: {type(e).__name__}: {e}")
