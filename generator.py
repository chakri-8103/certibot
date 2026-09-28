import os
import io
import re
import zipfile
import random
import datetime
import openpyxl
import pandas as pd
import pypdf
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib import colors

# Grade Points Mapping
GRADE_POINTS = {
    "O": 10, "10": 10,
    "A+": 9, "9": 9,
    "A": 8, "8": 8,
    "B+": 7, "7": 7,
    "B": 6, "6": 6,
    "C": 5, "5": 5,
    "D": 4, "P": 4, "4": 4,
    "F": 0, "FAIL": 0, "AB": 0, "ABSENT": 0, "0": 0
}

def create_pristine_background_template(original_pdf_path, output_clean_path="pristine_background.pdf"):
    """
    Strips only the dynamic text rendering BT...ET blocks from the master PDF content stream.
    Preserves 100% of background watermarks, logos, borders, fonts, and table lines.
    Zero white rectangles drawn!
    """
    reader = pypdf.PdfReader(original_pdf_path)
    page = reader.pages[0]
    
    contents_obj = page["/Contents"].get_object()

    def clean_stream_bytes(data_bytes):
        pattern = re.compile(b'BT[\\s\\S]*?ET')
        bt_matches = list(pattern.finditer(data_bytes))
        if len(bt_matches) == 39:
            # The exact 21 dynamic sample data text block indices to remove
            dynamic_indices = {1, 5, 6, 8, 11, 16, 18, 21, 22, 23, 24, 25, 27, 29, 32, 33, 34, 35, 36, 37, 38}
            new_b = bytearray()
            last_end = 0
            for i, m in enumerate(bt_matches):
                start, end = m.span()
                new_b.extend(data_bytes[last_end:start])
                if i not in dynamic_indices:
                    new_b.extend(data_bytes[start:end])
                last_end = end
            new_b.extend(data_bytes[last_end:])
            return bytes(new_b)
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

def ensure_default_signatures(signatures_dir="signatures"):
    os.makedirs(signatures_dir, exist_ok=True)
    existing = [f for f in os.listdir(signatures_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if not existing:
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
    issue_date
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

    # 1. Serial Number
    serial_no = student_record.get("Serial No:") or student_record.get("SERIAL NO") or student_record.get("SUC NUMBER") or "250070"
    serial_no_str = str(serial_no).strip()
    if serial_no_str.endswith(".0"):
        serial_no_str = serial_no_str[:-2]
    can.setFont("Helvetica-Bold", 14.8)
    can.drawString(510.84, ry(69.48), serial_no_str)

    # 2. Register Number (Hall Ticket)
    hall_ticket = student_record.get("HALL TICKET") or student_record.get("SUC NUMBER") or ""
    hall_ticket_str = str(hall_ticket).strip()
    if hall_ticket_str.endswith(".0"):
        hall_ticket_str = hall_ticket_str[:-2]
    can.setFont("Helvetica-Bold", 15.6)
    can.drawString(455.28, ry(218.64), hall_ticket_str)

    # 3. Student Name
    student_name = str(student_record.get("STUDENT NAME", "")).strip().upper()
    can.setFont("Helvetica-Bold", 11.5)
    can.drawString(230.04, ry(240.96), student_name)

    # 4. Date
    can.setFont("Helvetica", 10.7)
    can.drawString(33.72, ry(771.24), f"DATE : {issue_date}")

    # 5. Table Rows
    total_credits = 0
    total_points = 0
    has_fail = False

    row_y_starts = [336.60, 369.12, 401.76, 434.28, 466.80, 499.32, 531.84]

    for i, sub in enumerate(subjects_info):
        sname = sub["display_name"]
        cred = sub["credits"]
        gr = str(sub["grade"]).strip().upper()

        gp = GRADE_POINTS.get(gr, 0)
        if gr in ["FAIL", "F", "AB", "ABSENT"]:
            has_fail = True
            gp = 0

        cp = cred * gp
        total_credits += cred
        total_points += cp

        y_top = row_y_starts[i] if i < len(row_y_starts) else 336.60 + (i * 32.5)

        can.setFont("Helvetica", 9.8)
        if len(sname) > 46:
            can.setFont("Helvetica", 8.0)
        can.drawString(33.60, ry(y_top), sname)

        can.setFont("Helvetica", 10.7)
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

    can.setFont("Helvetica-Bold", 14.8)
    can.drawString(366.72, ry(559.44), str(total_credits))
    can.drawString(510.00, ry(559.44), str(total_points))

    # 7. SGPA
    can.setFont("Helvetica-Bold", 14.8)
    can.drawString(111.12, ry(578.76), sgpa_str)

    # 8. Controller Signature PNG
    if signature_path and os.path.exists(signature_path):
        can.drawImage(
            signature_path,
            380, ry(720),
            width=130, height=45,
            mask='auto',
            preserveAspectRatio=True
        )

    can.save()
    packet.seek(0)

    overlay_pdf = pypdf.PdfReader(packet)
    writer = pypdf.PdfWriter()
    merged_page = reader.pages[0]
    merged_page.merge_page(overlay_pdf.pages[0])
    writer.add_page(merged_page)

    clean_out_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(clean_out_path), exist_ok=True)
    with open(clean_out_path, "wb") as f:
        writer.write(f)

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
    progress_callback=None
):
    signatures = [os.path.join(signature_folder, f) for f in os.listdir(signature_folder)
                  if f.lower().endswith(('.png', '.jpg', '.jpeg'))] if os.path.exists(signature_folder) else []
    if not signatures:
        signatures = ensure_default_signatures("signatures")

    # Generate pristine clean background template without dynamic text
    clean_template_path = os.path.join(output_base_dir, "pristine_background.pdf")
    os.makedirs(output_base_dir, exist_ok=True)
    create_pristine_background_template(template_path, clean_template_path)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    all_sheets = wb.sheetnames

    if selected_sheets:
        if isinstance(selected_sheets, str):
            selected_sheets = [selected_sheets]
        sheets_to_process = [s for s in all_sheets if s in selected_sheets or s.strip() in [sel.strip() for sel in selected_sheets]]
    else:
        sheets_to_process = all_sheets

    total_students = 0
    generated_count = 0
    failed_count = 0
    generated_files = []

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

    for i, job in enumerate(student_jobs, 1):
        sheet_name = job["sheet_name"]
        ht_val = job["hall_ticket"]
        srow = job["record"]
        subjects_info = job["subjects_info"]

        if date_mode == "random":
            cur_date = get_random_date_between(start_date, end_date)
        else:
            cur_date = fixed_date or "06.04.2026"

        sig_path = random.choice(signatures)

        st_subs = []
        for sub in subjects_info:
            gr = srow.get(sub["col_key"], "FAIL")
            st_subs.append({
                "display_name": sub["display_name"],
                "credits": sub["credits"],
                "grade": gr if pd.notna(gr) else "FAIL"
            })

        sheet_dir = os.path.join(output_base_dir, sheet_name)
        out_pdf_path = os.path.join(sheet_dir, f"{ht_val}.pdf")

        try:
            generate_single_pdf(
                pristine_template_path=clean_template_path,
                student_record=srow,
                subjects_info=st_subs,
                output_path=out_pdf_path,
                signature_path=sig_path,
                issue_date=cur_date
            )
            generated_count += 1
            generated_files.append(out_pdf_path)
        except Exception as e:
            failed_count += 1
            print(f"Error generating PDF for {ht_val} in '{sheet_name}': {e}")

        if progress_callback:
            progress_callback(i, total_students, generated_count, failed_count, f"Processing {sheet_name}: {ht_val}.pdf")

    # Automatically create output.zip
    zip_path = os.path.join(output_base_dir, "output.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_base_dir):
            for file in files:
                if file in ["output.zip", "pristine_background.pdf"]:
                    continue
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, output_base_dir)
                zipf.write(file_path, arcname)

    return {
        "total": total_students,
        "generated": generated_count,
        "failed": failed_count,
        "zip_path": zip_path,
        "sample_pdf": generated_files[0] if generated_files else None
    }
