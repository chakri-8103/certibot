import sys
try:
    import pymupdfb
    print("pymupdfb is installed!")
except Exception as e:
    print("pymupdfb error:", e)

try:
    import fitz
    print("fitz version:", fitz.__version__)
except Exception as e:
    print("fitz error:", e)
