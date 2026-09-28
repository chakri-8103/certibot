import pypdf
import re

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]
content = page.get_contents().get_data().decode('latin1', errors='ignore')

# Find all TJ or Tj blocks
matches = re.findall(r'(\[.*?\]\s*TJ|\(.*?\)\s*Tj)', content)
print(f"Total text rendering commands found: {len(matches)}")
for i, m in enumerate(matches):
    print(f"{i:2d}: {m[:100]}")
