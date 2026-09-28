import pypdfium2 as pdfium

pdf = pdfium.PdfDocument("generated_pdfs/BCA TEL/250371611003.pdf")
page = pdf[0]
image = page.render(scale=2.0).to_pil()
image.save("preview_250371611003.png")
print("SUCCESS: Rendered preview_250371611003.png using pypdfium2!")

# Also render original template for comparison
pdf_orig = pdfium.PdfDocument("250371509001.pdf")
page_orig = pdf_orig[0]
image_orig = page_orig.render(scale=2.0).to_pil()
image_orig.save("preview_original.png")
print("SUCCESS: Rendered preview_original.png!")
