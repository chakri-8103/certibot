import pypdf
import re

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]
content = page.get_contents().get_data().decode('latin1', errors='ignore')

bt_blocks = list(re.finditer(r'BT[\s\S]*?ET', content))

print(f"Total BT blocks: {len(bt_blocks)}")

# The exact indices of dynamic data text blocks to remove:
dynamic_indices = {1, 5, 6, 8, 11, 16, 18, 21, 22, 23, 24, 25, 27, 29, 32, 33, 34, 35, 36, 37, 38}

new_content = content
for idx in sorted(dynamic_indices, reverse=True):
    block_text = bt_blocks[idx].group(0)
    new_content = new_content.replace(block_text, "")

page.get_contents().set_data(new_content.encode('latin1'))

writer = pypdf.PdfWriter()
writer.add_page(page)

with open("perfect_clean_template.pdf", "wb") as f:
    writer.write(f)

print(f"Removed exactly {len(dynamic_indices)} dynamic data blocks. Saved perfect_clean_template.pdf!")
