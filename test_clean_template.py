import pypdf
import io

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]

print("Page contents stream length:", len(page.get_contents().get_data()))

# Inspect instructions in content stream
content_data = page.get_contents().get_data().decode('latin1', errors='ignore')
print("Content sample (first 500 chars):")
print(content_data[:500])
