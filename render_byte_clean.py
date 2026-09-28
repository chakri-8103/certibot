import pypdfium2 as pdfium

pdf = pdfium.PdfDocument("byte_clean_template.pdf")
image = pdf[0].render(scale=2.0).to_pil()
image.save("byte_clean_template.png")
print("Saved byte_clean_template.png!")
