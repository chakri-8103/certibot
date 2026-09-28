import pypdf

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]
mb = page.mediabox
print(f"MediaBox: width={mb.width}, height={mb.height}, lower_left={mb.lower_left}, upper_right={mb.upper_right}")
