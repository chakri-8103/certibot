import os
import io
import re
import zipfile
import random
import shutil
import datetime
import urllib.request
import numpy as np
import openpyxl
import pandas as pd
import pypdf
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register Book Antiqua font with multiple search paths
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

def init_fonts():
    global FONT_REGULAR, FONT_BOLD
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        (os.path.join(base_dir, "fonts", "BKANT.TTF"), os.path.join(base_dir, "fonts", "ANTQUAB.TTF")),
        ("C:/Windows/Fonts/BKANT.TTF", "C:/Windows/Fonts/ANTQUAB.TTF"),
        (os.path.join(base_dir, "BKANT.TTF"), os.path.join(base_dir, "ANTQUAB.TTF"))
    ]
    for reg, bld in candidates:
        if os.path.exists(reg) and os.path.exists(bld):
            try:
                pdfmetrics.registerFont(TTFont("BookAntiqua", reg))
                pdfmetrics.registerFont(TTFont("BookAntiqua-Bold", bld))
                FONT_REGULAR = "BookAntiqua"
                FONT_BOLD = "BookAntiqua-Bold"
                break
            except Exception as e:
                print("Font registration error:", e)

init_fonts()

# Grade Points Mapping
GRADE_POINTS = {
    "O": 10, "10": 10,
    "A+": 9, "9": 9,
    "A": 8, "8": 8,
    "B+": 7, "7": 7,
    "B": 6, "6": 6,
    "C": 5, "5": 5,
    "D": 4, "P": 4, "4": 4,
    "F": 0, "FAIL": 0, "AB": 0, "ABSENT": 0, "ABS": 0, "0": 0,
    "NR": 0, "SMP": 0
}

# 14 Official Programme Header Titles
ALL_HEADER_TITLES = [
    "STATEMENT OF GRADES FOR B.COM (HONOURS) - COMPUTER APPLICATIONS",
    "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE",
    "STATEMENT OF GRADES FOR B.Sc (HONOURS) - DATA SCIENCE",
    "STATEMENT OF GRADES FOR B.Sc (HONOURS) - MICROBIOLOGY",
    "STATEMENT OF GRADES FOR B.Sc (HONOURS) - FISHERIES",
    "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ANIMATION",
    "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ARTIFICIAL INTELLIGENCE",
    "STATEMENT OF GRADES FOR B.Sc (HONOURS) - CHEMISTRY",
    "STATEMENT OF GRADES FOR BBA (HONOURS)",
    "STATEMENT OF GRADES FOR BBA (HONOURS) - DIGITAL MARKETING",
    "STATEMENT OF GRADES FOR BBA (HONOURS) - BUSINESS ANALYTICS",
    "STATEMENT OF GRADES FOR BCA (HONOURS) - COMPUTER APPLICATIONS",
    "STATEMENT OF GRADES FOR BCA (HONOURS) - DATA SCIENCE",
    "STATEMENT OF GRADES FOR B.COM (HONOURS) - BFSI"
]

# Sheet to Header Title Mapping Dictionary
SHEET_TO_HEADER_TITLE = {
    "BCOM-TEL": "STATEMENT OF GRADES FOR B.COM (HONOURS) - COMPUTER APPLICATIONS",
    "BCOM-HIN": "STATEMENT OF GRADES FOR B.COM (HONOURS) - COMPUTER APPLICATIONS",
    "BCOM-TEL-POINTS": "STATEMENT OF GRADES FOR B.COM (HONOURS) - COMPUTER APPLICATIONS",
    "BCOM": "STATEMENT OF GRADES FOR B.COM (HONOURS) - COMPUTER APPLICATIONS",
    "CS TEL": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE",
    "CS HIN": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE",
    "CS DIS": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE",
    "CS-TEL-II-SEM": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE",
    "CS": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE",
    "DS TEL": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - DATA SCIENCE",
    "DS HIN": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - DATA SCIENCE",
    "DS": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - DATA SCIENCE",
    "MB TEL": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - MICROBIOLOGY",
    "MB HIN": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - MICROBIOLOGY",
    "MB DIS": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - MICROBIOLOGY",
    "MB": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - MICROBIOLOGY",
    "FISH TEL": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - FISHERIES",
    "FISH HIN": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - FISHERIES",
    "FISH": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - FISHERIES",
    "ANIM TEL": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ANIMATION",
    "ANIM TEL ": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ANIMATION",
    "ANIM HIN": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ANIMATION",
    "ANIM": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ANIMATION",
    "AI-TEL": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ARTIFICIAL INTELLIGENCE",
    "AI-HIN": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ARTIFICIAL INTELLIGENCE",
    "AI": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ARTIFICIAL INTELLIGENCE",
    "CHEM TEL": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - CHEMISTRY",
    "CHEM HIN": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - CHEMISTRY",
    "CHEM": "STATEMENT OF GRADES FOR B.Sc (HONOURS) - CHEMISTRY",
    "BBA TEL": "STATEMENT OF GRADES FOR BBA (HONOURS)",
    "BBA HIN": "STATEMENT OF GRADES FOR BBA (HONOURS)",
    "BBA": "STATEMENT OF GRADES FOR BBA (HONOURS)",
    "BBA DM-TEL": "STATEMENT OF GRADES FOR BBA (HONOURS) - DIGITAL MARKETING",
    "BBA DM-HIN": "STATEMENT OF GRADES FOR BBA (HONOURS) - DIGITAL MARKETING",
    "BBA DM": "STATEMENT OF GRADES FOR BBA (HONOURS) - DIGITAL MARKETING",
    "BBA BA-TEL": "STATEMENT OF GRADES FOR BBA (HONOURS) - BUSINESS ANALYTICS",
    "BBA BA-HIN": "STATEMENT OF GRADES FOR BBA (HONOURS) - BUSINESS ANALYTICS",
    "BBA BA": "STATEMENT OF GRADES FOR BBA (HONOURS) - BUSINESS ANALYTICS",
    "BCA TEL": "STATEMENT OF GRADES FOR BCA (HONOURS) - COMPUTER APPLICATIONS",
    "BCA HIN": "STATEMENT OF GRADES FOR BCA (HONOURS) - COMPUTER APPLICATIONS",
    "BCA DIS": "STATEMENT OF GRADES FOR BCA (HONOURS) - COMPUTER APPLICATIONS",
    "BCA": "STATEMENT OF GRADES FOR BCA (HONOURS) - COMPUTER APPLICATIONS",
    "BCA DS TEL": "STATEMENT OF GRADES FOR BCA (HONOURS) - DATA SCIENCE",
    "BCA DS HIN": "STATEMENT OF GRADES FOR BCA (HONOURS) - DATA SCIENCE",
    "BCA DS DIS": "STATEMENT OF GRADES FOR BCA (HONOURS) - DATA SCIENCE",
    "BCA DS": "STATEMENT OF GRADES FOR BCA (HONOURS) - DATA SCIENCE",
    "BFSI TEL": "STATEMENT OF GRADES FOR B.COM (HONOURS) - BFSI",
    "BFSI HIN": "STATEMENT OF GRADES FOR B.COM (HONOURS) - BFSI",
    "BFSI": "STATEMENT OF GRADES FOR B.COM (HONOURS) - BFSI",
}

def resolve_header_title(sheet_name, custom_title=None, group_name=None):
    if custom_title and custom_title.strip() and custom_title.strip().upper() != "AUTO":
        return custom_title.strip()
    
    s = (sheet_name or "").strip().upper()
    g = (str(group_name) if group_name is not None else "").strip().upper()

    for k, v in SHEET_TO_HEADER_TITLE.items():
        if k.upper() == s:
            return v
    
    if "BBA DM" in s: return "STATEMENT OF GRADES FOR BBA (HONOURS) - DIGITAL MARKETING"
    elif "BBA BA" in s: return "STATEMENT OF GRADES FOR BBA (HONOURS) - BUSINESS ANALYTICS"
    elif s.startswith("BBA"): return "STATEMENT OF GRADES FOR BBA (HONOURS)"
    elif "BCA DS" in s: return "STATEMENT OF GRADES FOR BCA (HONOURS) - DATA SCIENCE"
    elif s.startswith("BCA"): return "STATEMENT OF GRADES FOR BCA (HONOURS) - COMPUTER APPLICATIONS"
    elif "BFSI" in s: return "STATEMENT OF GRADES FOR B.COM (HONOURS) - BFSI"
    elif s.startswith("BCOM"): return "STATEMENT OF GRADES FOR B.COM (HONOURS) - COMPUTER APPLICATIONS"
    elif s.startswith("CS") or "COMPUTER" in s: return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE"
    elif s.startswith("DS") or "DATA" in s: return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - DATA SCIENCE"
    elif s.startswith("MB") or "MICRO" in s: return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - MICROBIOLOGY"
    elif s.startswith("FISH"): return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - FISHERIES"
    elif s.startswith("ANIM"): return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ANIMATION"
    elif s.startswith("AI") or "ARTIFICIAL" in s: return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - ARTIFICIAL INTELLIGENCE"
    elif s.startswith("CHEM"): return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - CHEMISTRY"

    if "BBA DM" in g: return "STATEMENT OF GRADES FOR BBA (HONOURS) - DIGITAL MARKETING"
    elif "BBA BA" in g: return "STATEMENT OF GRADES FOR BBA (HONOURS) - BUSINESS ANALYTICS"
    elif g.startswith("BBA"): return "STATEMENT OF GRADES FOR BBA (HONOURS)"
    elif "BCA DS" in g: return "STATEMENT OF GRADES FOR BCA (HONOURS) - DATA SCIENCE"
    elif g.startswith("BCA"): return "STATEMENT OF GRADES FOR BCA (HONOURS) - COMPUTER APPLICATIONS"
    elif "BFSI" in g: return "STATEMENT OF GRADES FOR B.COM (HONOURS) - BFSI"
    elif "BCOM" in g: return "STATEMENT OF GRADES FOR B.COM (HONOURS) - COMPUTER APPLICATIONS"
    elif "CS" in g: return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE"
    elif "DS" in g: return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - DATA SCIENCE"
    elif "MB" in g: return "STATEMENT OF GRADES FOR B.Sc (HONOURS) - MICROBIOLOGY"

    return f"STATEMENT OF GRADES FOR {sheet_name.strip()}"


def create_pristine_background_template(original_pdf_path, output_clean_path="pristine_background.pdf"):
    """
    Strips only the dynamic text rendering BT...ET blocks and the old signature
    from the master PDF content stream.
    Preserves 100% of background watermarks, logos, borders, fonts, and table lines.
    Zero white rectangles drawn!
    """
    reader = pypdf.PdfReader(original_pdf_path)
    page = reader.pages[0]
    
    contents_obj = page["/Contents"].get_object()

    def clean_stream_bytes(data_bytes):
        # 1. Dynamic text removal
        pattern = re.compile(rb'BT[\s\S]*?ET')
        bt_matches = list(pattern.finditer(data_bytes))
        if len(bt_matches) == 39:
            # The exact dynamic sample data text block indices to remove (including 9 and 10 for header titles)
            dynamic_indices = {1, 5, 6, 8, 9, 10, 11, 16, 18, 21, 22, 23, 24, 25, 27, 29, 32, 33, 34, 35, 36, 37, 38}
            new_b = bytearray()
            last_end = 0
            for i, m in enumerate(bt_matches):
                start, end = m.span()
                new_b.extend(data_bytes[last_end:start])
                if i not in dynamic_indices:
                    new_b.extend(data_bytes[start:end])
                last_end = end
            new_b.extend(data_bytes[last_end:])
            data_bytes = bytes(new_b)

        # 2. Old signature removal (/Image11 Do and its clipping block)
        sig_pattern = re.compile(rb'q\r?\n550\.400024[\s\S]*?/Image11 Do\r?\nQ\r?\nQ')
        if sig_pattern.search(data_bytes):
            data_bytes = sig_pattern.sub(b'', data_bytes)
        elif b'/Image11 Do' in data_bytes:
            data_bytes = data_bytes.replace(b'/Image11 Do', b'')

        return data_bytes

    if isinstance(contents_obj, pypdf.generic.ArrayObject):
        for stream_ref in contents_obj:
            obj = stream_ref.get_object()
            cleaned = clean_stream_bytes(obj.get_data())
            obj.set_data(cleaned)
    else:
        cleaned = clean_stream_bytes(contents_obj.get_data())
        contents_obj.set_data(cleaned)

    writer = pypdf.PdfWriter()
    writer.add_page(page)

    with open(output_clean_path, "wb") as f:
        writer.write(f)

    return output_clean_path

def fetch_campus_signature(campus_id, cache_dir=None):
    """
    Fetches principal signature for campus_id from:
    https://analysis.aditya.ac.in/uploads/principal_signature/{campus_id}_1.jpg
    Converts white background to transparent and caches as PNG.
    """
    if not cache_dir:
        cache_dir = os.path.join(os.getcwd(), "static", "uploads", "signatures")
    os.makedirs(cache_dir, exist_ok=True)
    
    cid_clean = str(campus_id).strip()
    if cid_clean.endswith(".0"):
        cid_clean = cid_clean[:-2]
    if not cid_clean:
        cid_clean = "44"  # Default campus ID if empty
        
    target_png = os.path.join(cache_dir, f"campus_{cid_clean}_sig.png")
    if os.path.exists(target_png):
        return target_png

    url = f"https://analysis.aditya.ac.in/uploads/principal_signature/{cid_clean}_1.jpg"
    temp_jpg = os.path.join(cache_dir, f"temp_{cid_clean}.jpg")

    try:
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            jpg_bytes = resp.read()
            with open(temp_jpg, "wb") as f:
                f.write(jpg_bytes)

        im = Image.open(temp_jpg).convert("RGBA")
        arr = np.array(im)
        # Mask near-white pixels (paper background)
        white_mask = (arr[:, :, 0] > 195) & (arr[:, :, 1] > 195) & (arr[:, :, 2] > 195)
        arr[white_mask, 3] = 0
        clean_im = Image.fromarray(arr)
        clean_im.save(target_png, "PNG")

        try:
            if os.path.exists(temp_jpg):
                os.remove(temp_jpg)
        except Exception:
            pass

        return target_png
    except Exception as e:
        print(f"Warning: Could not fetch signature from {url}: {e}")
        return None


def ensure_default_signatures(signatures_dir="signatures"):
    os.makedirs(signatures_dir, exist_ok=True)
    existing = [f for f in os.listdir(signatures_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if not existing:
        root_default = os.path.join(os.getcwd(), "default_signature.png")
        if os.path.exists(root_default):
            dest = os.path.join(signatures_dir, "default_signature.png")
            shutil.copy(root_default, dest)
            existing.append("default_signature.png")
        else:
            colors_list = [(0, 0, 128, 255), (0, 51, 102, 255), (20, 20, 80, 255)]
            for i, col in enumerate(colors_list, 1):
                img = Image.new("RGBA", (300, 100), (255, 255, 255, 0))
                draw = ImageDraw.Draw(img)
                draw.line([(20, 60), (60, 20), (100, 70), (140, 30), (180, 80), (250, 40)], fill=col, width=4)
                draw.line([(10, 75 + i*2), (280, 70)], fill=col, width=3)
                sig_file = os.path.join(signatures_dir, f"controller_sig_{i}.png")
                img.save(sig_file)
                existing.append(f"controller_sig_{i}.png")
    return [os.path.join(signatures_dir, f) for f in existing]

def format_display_date(date_val):
    if not date_val:
        return "06.04.2026"
    date_str = str(date_val).strip()
    try:
        if "-" in date_str:
            dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
            return dt.strftime("%d.%m.%Y")
        elif "/" in date_str:
            parts = date_str.split("/")
            if len(parts[0]) == 4:
                dt = datetime.datetime.strptime(date_str, "%Y/%m/%d")
            else:
                dt = datetime.datetime.strptime(date_str, "%d/%m/%Y")
            return dt.strftime("%d.%m.%Y")
    except Exception:
        pass
    return date_str

def get_random_date_between(start_date_str, end_date_str):
    try:
        d1 = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")
        d2 = datetime.datetime.strptime(end_date_str, "%Y-%m-%d")
        if d1 > d2:
            d1, d2 = d2, d1
        delta_days = (d2 - d1).days
        random_day = d1 + datetime.timedelta(days=random.randint(0, delta_days))
        return random_day.strftime("%d.%m.%Y")
    except Exception:
        return datetime.datetime.now().strftime("%d.%m.%Y")

def generate_single_pdf(
    pristine_template_path,
    student_record,
    subjects_info,
    output_path,
    signature_path,
    issue_date,
    header_line_1="STATEMENT OF GRADES FOR B.Sc (HONOURS) - COMPUTER SCIENCE",
    header_line_2="DEGREE EXAMINATIONS AT THE END OF FIRST SEMESTER - JANUARY - 2026"
):
    reader = pypdf.PdfReader(pristine_template_path)
    template_page = reader.pages[0]
    page_width = float(template_page.mediabox.width)
    page_height = float(template_page.mediabox.height)

    def ry(top_y):
        return page_height - top_y

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    # Zero white boxes drawn! Text sits directly over background watermarks & table grids.
    can.setFillColor(colors.black)

    # 0. Degree Programme & Examination Header Titles (Centered)
    if header_line_1:
        h1_str = str(header_line_1).strip()
        can.setFont(FONT_BOLD, 10.2 if len(h1_str) > 65 else 11.5)
        can.drawCentredString(page_width / 2.0, ry(176.88), h1_str)

    if header_line_2:
        h2_str = str(header_line_2).strip()
        can.setFont(FONT_BOLD, 10.2 if len(h2_str) > 68 else 11.5)
        can.drawCentredString(page_width / 2.0, ry(195.24), h2_str)

    # 1. Serial Number
    serial_no = student_record.get("Serial No:") or student_record.get("SERIAL NO") or student_record.get("SUC NUMBER") or "250070"
    serial_no_str = str(serial_no).strip()
    if serial_no_str.endswith(".0"):
        serial_no_str = serial_no_str[:-2]
    can.setFont(FONT_BOLD, 14.5)
    can.drawString(510.84, ry(69.48), serial_no_str)

    # 2. Register Number (Hall Ticket)
    hall_ticket = student_record.get("HALL TICKET") or student_record.get("SUC NUMBER") or ""
    hall_ticket_str = str(hall_ticket).strip()
    if hall_ticket_str.endswith(".0"):
        hall_ticket_str = hall_ticket_str[:-2]
    can.setFont(FONT_BOLD, 15.0)
    can.drawString(455.28, ry(218.64), hall_ticket_str)

    # 3. Student Name (All in one line via auto-scaling; greedy wrap if extraordinarily long)
    student_name = str(student_record.get("STUDENT NAME", "")).strip().upper()
    max_name_w = 320.0  # Right margin boundary before table edge at 553.5 pt

    # Try single line first by scaling down font size
    sz = 11.5
    while can.stringWidth(student_name, FONT_BOLD, sz) > max_name_w and sz > 7.6:
        sz -= 0.1
    sz = round(sz, 1)

    if can.stringWidth(student_name, FONT_BOLD, sz) <= max_name_w:
        # All name in ONE line cleanly
        can.setFont(FONT_BOLD, sz)
        can.drawString(230.04, ry(240.96), student_name)
    else:
        # Extraordinarily long name: greedy wrap so Line 1 fills up completely and remaining words go to Line 2
        words = student_name.split()
        wrap_sz = 9.0
        l1_words, l2_words = [], []
        for w in words:
            test_l1 = " ".join(l1_words + [w])
            if can.stringWidth(test_l1, FONT_BOLD, wrap_sz) <= max_name_w:
                l1_words.append(w)
            else:
                l2_words.append(w)

        l1 = " ".join(l1_words)
        l2 = " ".join(l2_words)

        while (can.stringWidth(l1, FONT_BOLD, wrap_sz) > max_name_w or (l2 and can.stringWidth(l2, FONT_BOLD, wrap_sz) > max_name_w)) and wrap_sz > 7.0:
            wrap_sz -= 0.2
        wrap_sz = round(wrap_sz, 1)

        can.setFont(FONT_BOLD, wrap_sz)
        if l2:
            can.drawString(230.04, ry(235.50), l1)
            can.drawString(230.04, ry(248.50), l2)
        else:
            can.drawString(230.04, ry(240.96), l1)

    # 4. Date
    can.setFont(FONT_REGULAR, 10.7)
    can.drawString(33.72, ry(771.24), f"DATE : {issue_date}")

    # 5. Table Rows
    total_credits = 0
    total_points = 0
    has_fail = False

    num_subs = len(subjects_info)

    # Uniform font size for ALL subjects in this marksheet
    # Base size according to subject count
    if num_subs <= 7:
        base_sub_size = 9.5
        table_num_size = 10.5
    elif num_subs == 8:
        base_sub_size = 8.8
        table_num_size = 9.8
    elif num_subs == 9:
        base_sub_size = 8.2
        table_num_size = 9.2
    else:  # 10 or more subjects
        base_sub_size = 7.6
        table_num_size = 8.5

    # Check if ANY subject name is too long to fit in the column (width ~ 320 pt)
    # If any subject overflows, scale down the uniform font size for ALL subjects
    max_col_w = 320.0
    for sub in subjects_info:
        s_title = str(sub.get("display_name", "")).strip()
        while can.stringWidth(s_title, FONT_REGULAR, base_sub_size) > max_col_w and base_sub_size > 6.0:
            base_sub_size -= 0.2

    uniform_sub_size = round(base_sub_size, 1)

    # Row vertical positioning
    # Standard 7 rows: [336.60, 369.12, 401.76, 434.28, 466.80, 499.32, 531.84]
    standard_y = [336.60, 369.12, 401.76, 434.28, 466.80, 499.32, 531.84]
    if num_subs <= 7:
        row_y_starts = standard_y[:num_subs]
    else:
        # Distribute rows evenly between 336.60 and 538.00 (just above TOTAL line at 553.5)
        step = (538.00 - 336.60) / max(num_subs - 1, 1)
        row_y_starts = [336.60 + (i * step) for i in range(num_subs)]

    for i, sub in enumerate(subjects_info):
        sname = str(sub["display_name"]).strip()
        cred = sub["credits"]
        gr = str(sub["grade"]).strip().upper()

        gp = GRADE_POINTS.get(gr, 0)
        if gr in ["FAIL", "F", "AB", "ABSENT", "ABS", "NR", "SMP"]:
            has_fail = True
            gp = 0

        cp = cred * gp
        total_credits += cred
        total_points += cp

        y_top = row_y_starts[i] if i < len(row_y_starts) else 336.60 + (i * 32.5)

        # All subjects use the EXACT SAME uniform font size
        can.setFont(FONT_REGULAR, uniform_sub_size)
        can.drawString(33.60, ry(y_top), sname)

        # Table numbers & Grade in Book Antiqua
        can.setFont(FONT_REGULAR, table_num_size)
        can.drawString(371.52, ry(y_top), str(cred))
        can.drawString(420.00, ry(y_top), gr)
        can.drawString(475.00, ry(y_top), str(gp))
        can.drawString(515.00, ry(y_top), str(cp))

    # 6. Totals Row
    excel_sgpa = student_record.get("SGPA")
    if pd.notna(excel_sgpa) and str(excel_sgpa).strip():
        try:
            sgpa_val = float(excel_sgpa)
            sgpa_str = f"{sgpa_val:.2f}"
        except ValueError:
            sgpa_str = str(excel_sgpa)
    else:
        if has_fail or total_credits == 0:
            sgpa_str = "0.00"
        else:
            sgpa_str = f"{(total_points / total_credits):.2f}"

    can.setFont(FONT_BOLD, 14.5)
    can.drawString(366.72, ry(559.44), str(total_credits))
    can.drawString(510.00, ry(559.44), str(total_points))

    # 7. SGPA
    can.setFont(FONT_BOLD, 14.5)
    can.drawString(111.12, ry(578.76), sgpa_str)

    # 8. Controller / Principal Signature PNG
    if signature_path and os.path.exists(signature_path):
        can.drawImage(
            signature_path,
            410, ry(718),
            width=95, height=42,
            mask='auto',
            preserveAspectRatio=True
        )

    can.save()
    packet.seek(0)

    overlay_pdf = pypdf.PdfReader(packet)
    writer = pypdf.PdfWriter()
    merged_page = reader.pages[0]
    if output_path is None:
        out_buf = io.BytesIO()
        writer.write(out_buf)
        out_buf.seek(0)
        return out_buf.getvalue()
    elif isinstance(output_path, (io.BytesIO, io.RawIOBase, io.BufferedIOBase)):
        writer.write(output_path)
        return output_path
    else:
        clean_out_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(clean_out_path), exist_ok=True)
        with open(clean_out_path, "wb") as f:
            writer.write(f)
        return clean_out_path

def process_excel_and_generate_all(
    template_path,
    excel_path,
    signature_folder,
    output_base_dir="generated_pdfs",
    selected_sheets=None,
    date_mode="fixed",
    fixed_date="06.04.2026",
    start_date="2026-04-01",
    end_date="2026-04-15",
    default_campus_id="31",
    header_title="AUTO",
    semester="FIRST SEMESTER",
    exam_month="JANUARY",
    exam_year="2026",
    progress_callback=None
):
    sig_cache_dir = signature_folder if signature_folder else os.path.join(os.getcwd(), "static", "uploads", "signatures")
    default_sig_path = fetch_campus_signature(default_campus_id, cache_dir=sig_cache_dir)

    signatures = [os.path.join(signature_folder, f) for f in os.listdir(signature_folder)
                  if f.lower().endswith(('.png', '.jpg', '.jpeg')) and not f.startswith("temp_")] if (signature_folder and os.path.exists(signature_folder)) else []
    if not signatures and not default_sig_path:
        signatures = ensure_default_signatures("signatures")

    # Clean previous output directory so only output.zip will exist
    if os.path.exists(output_base_dir):
        for item in os.listdir(output_base_dir):
            item_path = os.path.join(output_base_dir, item)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                elif os.path.isfile(item_path):
                    os.remove(item_path)
            except Exception:
                pass

    os.makedirs(output_base_dir, exist_ok=True)

    # Use master pristine_template.pdf directly without creating extra background files
    clean_template_path = os.path.join(os.getcwd(), "pristine_template.pdf")
    if not os.path.exists(clean_template_path):
        clean_template_path = os.path.join(output_base_dir, "pristine_background.pdf")
        create_pristine_background_template(template_path, clean_template_path)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    all_sheets = wb.sheetnames

    if selected_sheets:
        if isinstance(selected_sheets, str):
            selected_sheets = [selected_sheets]
        sheets_to_process = [s for s in all_sheets if s in selected_sheets or s.strip() in [sel.strip() for sel in selected_sheets]]
    else:
        # Default to first sheet (Single sheet generation only)
        sheets_to_process = [all_sheets[0]] if all_sheets else []

    if not sheets_to_process:
        raise ValueError("Selected sheet was not found in the uploaded Excel file.")

    total_students = 0
    generated_count = 0
    failed_count = 0

    student_jobs = []
    for sheet_name in sheets_to_process:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        if df.empty or "HALL TICKET" not in [str(c).upper() for c in df.columns]:
            continue
        
        ht_col = [c for c in df.columns if "HALL TICKET" in str(c).upper()][0]
        
        cols = list(df.columns)
        lang_idx = None
        res_idx = None
        for i, c in enumerate(cols):
            if "LANGUAGE" in str(c).upper():
                lang_idx = i
            elif "RESULT" in str(c).upper():
                res_idx = i
                
        if lang_idx is not None and res_idx is not None and res_idx > lang_idx + 1:
            subject_cols = cols[lang_idx+1:res_idx]
        else:
            subject_cols = [c for c in cols if c not in [ht_col, "S.NO", "SUC NUMBER", "STUDENT NAME", "GROUP", "BRANCH", "LANGUAGE", "RESULT", "SGPA"]]

        row0 = df.iloc[0]
        subjects_info = []
        for col in subject_cols:
            cred_val = row0[col]
            try:
                cred = int(cred_val)
            except Exception:
                cred = 3
            subjects_info.append({
                "col_key": col,
                "display_name": str(col).strip(),
                "credits": cred
            })

        valid_rows = df[df[ht_col].notna()]
        if not valid_rows.empty and valid_rows.index[0] == 0:
            valid_rows = valid_rows.iloc[1:]

        for _, srow in valid_rows.iterrows():
            ht_val = str(srow[ht_col]).strip()
            if ht_val.endswith(".0"):
                ht_val = ht_val[:-2]
            if ht_val and ht_val.lower() != "nan":
                student_jobs.append({
                    "sheet_name": sheet_name.strip(),
                    "hall_ticket": ht_val,
                    "record": srow.to_dict(),
                    "subjects_info": subjects_info
                })

    total_students = len(student_jobs)
    if progress_callback:
        progress_callback(0, total_students, 0, 0, "Starting generation...")

    zip_path = os.path.join(output_base_dir, "output.zip")

    # Generate directly into output.zip in memory — NO individual PDF files stored on disk!
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for i, job in enumerate(student_jobs, 1):
            sheet_name = job["sheet_name"]
            ht_val = job["hall_ticket"]
            srow = job["record"]
            subjects_info = job["subjects_info"]

            if date_mode == "random":
                cur_date = get_random_date_between(start_date, end_date)
            else:
                cur_date = format_display_date(fixed_date or "06.04.2026")

            # Determine student campus ID
            student_cid = str(default_campus_id).strip()
            for col_name in srow.keys():
                col_upper = str(col_name).upper().strip()
                if col_upper in ["CAMPUS ID", "CAMPUS_ID", "CAMPUS", "CAMPUS CODE", "COLLEGE CODE"]:
                    if pd.notna(srow[col_name]):
                        cval = str(srow[col_name]).strip()
                        if cval.endswith(".0"):
                            cval = cval[:-2]
                        if cval:
                            student_cid = cval
                            break
            
            if student_cid == str(default_campus_id).strip() and default_sig_path:
                sig_path = default_sig_path
            else:
                sig_path = fetch_campus_signature(student_cid, cache_dir=sig_cache_dir)

            st_subs = []
            for sub in subjects_info:
                gr = srow.get(sub["col_key"], "FAIL")
                st_subs.append({
                    "display_name": sub["display_name"],
                    "credits": sub["credits"],
                    "grade": gr if pd.notna(gr) else "FAIL"
                })

            # Determine dynamic Header Lines
            h1 = resolve_header_title(sheet_name, custom_title=header_title, group_name=srow.get("GROUP"))
            sem_clean = semester.strip() if semester else "FIRST SEMESTER"
            month_clean = exam_month.strip() if exam_month else "JANUARY"
            year_clean = str(exam_year).strip() if exam_year else "2026"
            h2 = f"DEGREE EXAMINATIONS AT THE END OF {sem_clean} - {month_clean} - {year_clean}"

            try:
                pdf_bytes = generate_single_pdf(
                    pristine_template_path=clean_template_path,
                    student_record=srow,
                    subjects_info=st_subs,
                    output_path=None,  # Generated in memory, zero files on disk!
                    signature_path=sig_path,
                    issue_date=cur_date,
                    header_line_1=h1,
                    header_line_2=h2
                )
                zipf.writestr(f"{sheet_name}/{ht_val}.pdf", pdf_bytes)
                generated_count += 1
            except Exception as e:
                failed_count += 1
                print(f"Error generating PDF for {ht_val} in '{sheet_name}': {e}")

            if progress_callback:
                progress_callback(i, total_students, generated_count, failed_count, f"Processing {sheet_name}: {ht_val}.pdf")

    # Clean up pristine_background.pdf if it was temporarily generated
    temp_bg = os.path.join(output_base_dir, "pristine_background.pdf")
    if os.path.exists(temp_bg):
        try:
            os.remove(temp_bg)
        except Exception:
            pass

    return {
        "total": total_students,
        "generated": generated_count,
        "failed": failed_count,
        "zip_path": zip_path,
        "sample_pdf": None
    }
