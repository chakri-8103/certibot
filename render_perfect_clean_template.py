import pypdfium2 as pdfium

pdf = pdfium.PdfDocument("perfect_clean_template.pdf")
image = pdf[0].render(scale=2.0).to_pil()
image.save("perfect_clean_template.png")
print("Saved perfect_clean_template.png!")
