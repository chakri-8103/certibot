import pypdfium2 as pdfium

pdf = pdfium.PdfDocument("pristine_template.pdf")
image = pdf[0].render(scale=2.0).to_pil()
image.save("pristine_template.png")
print("Saved pristine_template.png!")
