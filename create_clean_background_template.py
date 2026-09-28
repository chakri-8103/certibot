import pypdf
import re
import io

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]
content_bytes = page.get_contents().get_data()
content_str = content_bytes.decode('latin1', errors='ignore')

# Identify blocks to REMOVE (the sample dynamic data blocks)
# 1. Block 1: Serial number '250070' (y=92.639999, size 19.680799)
# 8. Block 8: Name 'AAKULA SRI VARDHAN' (x=306.720001, y=321.279999)
# 18. Block 18: Register No '250371509001' (x=607.039978, y=291.519989)
# 27. Block 27: Totals '18 79' (y=745.919983)
# 29. Block 29: SGPA '0.00' (y=771.679993)
# 38. Block 38: Date 'DATE : 06.04.2026' (y=1028.319946)
# Marks Table Subject Names & Grades:
# Blocks 5, 6, 11, 16, 21, 22, 23, 24, 25, 32, 33, 34, 35, 36, 37

dynamic_y_targets = [
    92.639999,    # Serial No value
    291.519989,   # Register No value
    321.279999,   # Name value (x=306)
    448.000000,   # ENGLISH
    448.799988,   # ENGLISH grades
    491.359985,   # TELUGU
    492.160004,   # TELUGU grades
    534.880005,   # AI
    535.679993,   # AI grades
    578.239990,   # FIT
    579.039978,   # FIT grades
    613.119995,   # FIT LAB 1
    630.080017,   # FIT LAB 2
    622.400024,   # FIT LAB grades
    664.960022,   # C
    665.760010,   # C grades
    708.320007,   # C LAB
    709.119995,   # C LAB grades
    745.919983,   # Total numbers
    771.679993,   # SGPA number
    1028.319946,  # Date
]

bt_blocks = re.findall(r'BT[\s\S]*?ET', content_str)
print(f"Total BT blocks: {len(bt_blocks)}")

removed_count = 0
new_content_str = content_str

for block in bt_blocks:
    # Check if block matches any dynamic Y coordinate or target
    is_dynamic = False
    
    # Check if block is Name value at x=306.72
    if "306.720001" in block:
        is_dynamic = True
    elif "681.119995 92.639999" in block:
        is_dynamic = True
    elif "607.039978 291.519989" in block:
        is_dynamic = True
    elif "488.959991 745.919983" in block:
        is_dynamic = True
    elif "148.160004 771.679993" in block:
        is_dynamic = True
    elif "44.959999 1028.319946" in block:
        is_dynamic = True
    else:
        # Check table marks rows (y between 440 and 720)
        tm_match = re.search(r'1 0 0\.000000 -1 ([^\s]+) ([^\s]+) Tm', block)
        if tm_match:
            x_val = float(tm_match.group(1))
            y_val = float(tm_match.group(2))
            if 435.0 <= y_val <= 715.0:
                is_dynamic = True

    if is_dynamic:
        new_content_str = new_content_str.replace(block, "")
        removed_count += 1

print(f"Removed {removed_count} dynamic text blocks from content stream.")

# Replace page content stream
new_content_bytes = new_content_str.encode('latin1')
page.get_contents().set_data(new_content_bytes)

writer = pypdf.PdfWriter()
writer.add_page(page)

with open("clean_template.pdf", "wb") as f:
    writer.write(f)

print("Saved clean_template.pdf with zero dynamic text!")
