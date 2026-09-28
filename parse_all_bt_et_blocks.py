import pypdf
import re

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]
content_bytes = page.get_contents().get_data()
content_str = content_bytes.decode('latin1', errors='ignore')

# Find all BT ... ET blocks
bt_blocks = re.findall(r'BT[\s\S]*?ET', content_str)
print(f"Total BT...ET blocks found: {len(bt_blocks)}")

for i, block in enumerate(bt_blocks):
    # Find Tm matrix in block
    tm_match = re.search(r'1 0 [^\s]+ -1 ([^\s]+) ([^\s]+) Tm', block)
    pos_str = f"x={tm_match.group(1)}, y={tm_match.group(2)}" if tm_match else "x=?, y=?"
    
    # Clean string representation
    block_preview = block.replace('\n', ' ')
    print(f"Block {i:2d} | {pos_str:25s} | {block_preview[:100]}")
