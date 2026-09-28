import os
import io
import pypdf
import pandas as pd
import pypdfium2 as pdfium
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib import colors

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

def generate_grade_sheet(
    template_path,
    student_record,
    subjects_info,
    output_path,
    signature_path="default_signature.png",
    issue_date="06.04.2026"
):
    reader = pypdf.PdfReader(template_path)
    template_page = reader.pages[0]
    page_width = float(template_page.mediabox.width)
    page_height = float(template_page.mediabox.height)
    
    def ry(top_y):
        return page_height - top_y

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    def whiteout(x, top_y, w, h):
        can.setFillColor(colors.white)
        can.rect(x, page_height - top_y - h, w, h, fill=1, stroke=0)
    
    # 1. Whiteout precise target areas
    whiteout(500, 56, 90, 20)       # Serial No
    whiteout(445, 203, 140, 22)     # Register No
    whiteout(225, 230, 340, 16)     # Student Name
    whiteout(33, 325, 530, 222)     # Marks Table Rows
    whiteout(360, 545, 190, 20)     # Totals Row
    whiteout(105, 564, 120, 20)     # SGPA Row
    whiteout(33, 760, 180, 16)      # Date
    whiteout(360, 665, 190, 60)     # Controller Signature Box
    
    
    # 2. Draw Dynamic Content
    can.setFillColor(colors.black)
    
    # Serial Number
    serial_no = student_record.get("Serial No:") or student_record.get("SERIAL NO") or student_record.get("SUC NUMBER") or "250070"
    serial_no_str = str(serial_no).strip()
    if serial_no_str.endswith(".0"):
        serial_no_str = serial_no_str[:-2]
    can.setFont("Helvetica-Bold", 14.8)
    can.drawString(510.84, ry(69.48), serial_no_str)
    
    # Register Number (Hall Ticket)
    hall_ticket = student_record.get("HALL TICKET") or student_record.get("SUC NUMBER") or ""
    hall_ticket_str = str(hall_ticket).strip()
    if hall_ticket_str.endswith(".0"):
        hall_ticket_str = hall_ticket_str[:-2]
    can.setFont("Helvetica-Bold", 15.6)
    can.drawString(455.28, ry(218.64), hall_ticket_str)
    
    # Student Name
    student_name = str(student_record.get("STUDENT NAME", "")).strip().upper()
    can.setFont("Helvetica-Bold", 11.5)
    can.drawString(230.04, ry(240.96), student_name)
    
    # Date
    can.setFont("Helvetica", 10.7)
    can.drawString(33.72, ry(771.24), f"DATE : {issue_date}")
    
    # Table Rows
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
        
        # Subject Name
        can.setFont("Helvetica", 9.8)
        if len(sname) > 46:
            can.setFont("Helvetica", 8.0)
        can.drawString(33.60, ry(y_top), sname)
        
        # Credits, Grade, Grade Points, Credit Points
        can.setFont("Helvetica", 10.7)
        can.drawString(371.52, ry(y_top), str(cred))
        can.drawString(420.00, ry(y_top), gr)
        can.drawString(475.00, ry(y_top), str(gp))
        can.drawString(515.00, ry(y_top), str(cp))
        
    # Totals Row
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
    
    # SGPA
    can.setFont("Helvetica-Bold", 14.8)
    can.drawString(111.12, ry(578.76), sgpa_str)
    
    # Controller Signature
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
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

# Re-run for BCA TEL student 250371611003
excel_file = "I SEM DIGIVAL PURPOSE-25 AB F+RV.xlsx"
df = pd.read_excel(excel_file, sheet_name="BCA TEL")
lang_idx = df.columns.get_loc("LANGUAGE")
res_idx = df.columns.get_loc("RESULT")
subject_cols = list(df.columns[lang_idx+1:res_idx])
row0 = df.iloc[0]
subjects_info = [{"col_key": col, "display_name": str(col).strip(), "credits": int(row0[col]) if str(row0[col]).isdigit() else 3} for col in subject_cols]

srow = df[df["HALL TICKET"] == 250371611003.0].iloc[0]
st_subs = [{"display_name": sub["display_name"], "credits": sub["credits"], "grade": srow[sub["col_key"]]} for sub in subjects_info]

out_pdf = "generated_pdfs/BCA TEL/250371611003.pdf"
generate_grade_sheet(
    template_path="250371509001.pdf",
    student_record=srow.to_dict(),
    subjects_info=st_subs,
    output_path=out_pdf,
    signature_path="default_signature.png",
    issue_date="06.04.2026"
)

# Render updated preview image
pdf = pdfium.PdfDocument(out_pdf)
image = pdf[0].render(scale=2.0).to_pil()
image.save("preview_250371611003.png")
print("Rendered perfected preview_250371611003.png!")
