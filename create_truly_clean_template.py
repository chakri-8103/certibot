import pypdf
import re

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]
content = page.get_contents().get_data().decode('latin1', errors='ignore')

# Extract all BT...ET blocks
bt_blocks = list(re.finditer(r'BT[\s\S]*?ET', content))
print(f"Total BT blocks found: {len(bt_blocks)}")

dynamic_indices = {1, 5, 6, 8, 11, 16, 18, 21, 22, 23, 24, 25, 27, 29, 32, 33, 34, 35, 36, 37, 38}

# Rebuild content string by replacing dynamic blocks with empty string
new_content = ""
last_idx = 0

for i, match in enumerate(bt_blocks):
    start, end = match.span()
    new_content += content[last_idx:start]
    if i in dynamic_indices:
        # Replace block with empty space (remove text rendering)
        new_content += ""
    else:
        new_content += match.group(0)
    last_idx = end

new_content += content[last_idx:]

# Update content stream
page.get_contents().set_data(new_content.encode('latin1'))

writer = pypdf.PdfWriter()
writer.add_page(page)

with open("pristine_template.pdf", "wb") as f:
    writer.write(f)

print("SUCCESS: Saved pristine_template.pdf with zero dynamic text!")
