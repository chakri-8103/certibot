import pypdfium2 as pdfium

pdf = pdfium.PdfDocument("clean_template.pdf")
image = pdf[0].render(scale=2.0).to_pil()
image.save("clean_template.png")
print("Saved clean_template.png!")
