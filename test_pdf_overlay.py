import os
import io
import pypdf
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors

template_pdf_path = "250371509001.pdf"
output_pdf_path = "sample_generated.pdf"

# Read original template
reader = pypdf.PdfReader(template_pdf_path)
template_page = reader.pages[0]
page_width = float(template_page.mediabox.width)
page_height = float(template_page.mediabox.height)

packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=(page_width, page_height))

# In ReportLab, Y=0 is at bottom of page!
# Helper function to convert Y from top (0 at top) to ReportLab Y (0 at bottom)
def ry(top_y):
    return page_height - top_y

# 1. Whiteout / Redact dynamic regions on template
can.setFillColor(colors.white)

# Serial No region
can.rect(500, ry(82), 80, 18, fill=1, stroke=0)

# Register No region
can.rect(445, ry(230), 135, 18, fill=1, stroke=0)

# Student Name region
can.rect(225, ry(255), 340, 18, fill=1, stroke=0)

# Table Marks Rows region (from top_y=330 to 552)
can.rect(33, ry(552), 530, 222, fill=1, stroke=0)

# Total summary row (from top_y=552 to 570)
can.rect(360, ry(572), 170, 18, fill=1, stroke=0)

# SGPA region (from top_y=572 to 590)
can.rect(105, ry(590), 120, 18, fill=1, stroke=0)

# Date region (from top_y=765 to 785)
can.rect(33, ry(785), 200, 20, fill=1, stroke=0)

# Controller Signature area (from top_y=660 to 725)
can.rect(370, ry(725), 180, 65, fill=1, stroke=0)


# 2. Draw Dynamic Values
can.setFillColor(colors.black)

# Serial Number
can.setFont("Helvetica-Bold", 14.8)
can.drawString(510.84, ry(81.5), "252207")

# Register Number
can.setFont("Helvetica-Bold", 15.6)
can.drawString(455.28, ry(227.5), "250371505002")

# Student Name
can.setFont("Helvetica-Bold", 11.5)
can.drawString(230.04, ry(251.5), "AMBUJALAPU NAGENDHRA")

# Date
can.setFont("Helvetica", 10.7)
can.drawString(33.72, ry(781), "DATE : 27.09.2026")

# Sample Table Subjects
subjects = [
    ("ENGLISH", 3, "A", 8, 24),
    ("TELUGU", 3, "A+", 9, 27),
    ("INTRODUCTION TO ARTIFICIAL INTELLIGENCE", 4, "B+", 7, 28),
    ("COMPUTER FUNDAMENTALS AND OFFICE AUTOMATION", 3, "O", 10, 30),
    ("PROBLEM SOLVING USING C", 3, "A+", 9, 27),
    ("COMPUTER FUNDAMENTALS AND OFFICE AUTOMATION - LAB", 1, "A+", 9, 9),
    ("PROBLEM SOLVING USING C LAB", 1, "A+", 9, 9),
]

row_y_starts = [336.60, 369.12, 401.76, 434.28, 466.80, 499.32, 531.84]

for i, (sname, cred, gr, gp, cp) in enumerate(subjects):
    y_top = row_y_starts[i] if i < len(row_y_starts) else 336.60 + (i * 32.5)
    
    # Subject Name
    can.setFont("Helvetica", 9.8)
    can.drawString(33.60, ry(y_top + 9), sname)
    
    # Credits, Grade, Grade Points, Credit Points
    can.setFont("Helvetica", 10.7)
    can.drawString(371.52, ry(y_top + 9), str(cred))
    can.drawString(420.00, ry(y_top + 9), str(gr))
    can.drawString(475.00, ry(y_top + 9), str(gp))
    can.drawString(515.00, ry(y_top + 9), str(cp))

# Totals
can.setFont("Helvetica-Bold", 14.8)
can.drawString(366.72, ry(568.5), "18")
can.drawString(510.00, ry(568.5), "154")

# SGPA
can.setFont("Helvetica-Bold", 14.8)
can.drawString(111.12, ry(587.5), "8.56")

can.save()

packet.seek(0)
overlay_pdf = pypdf.PdfReader(packet)

# Merge overlay page onto template page
writer = pypdf.PdfWriter()
merged_page = reader.pages[0]
merged_page.merge_page(overlay_pdf.pages[0])
writer.add_page(merged_page)

with open(output_pdf_path, "wb") as f:
    writer.write(f)

print(f"SUCCESS: Generated test overlay PDF at '{output_pdf_path}'")
