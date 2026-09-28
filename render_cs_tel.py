import pypdfium2 as pdfium

pdf = pdfium.PdfDocument("generated_pdfs/CS TEL/250371509001.pdf")
image = pdf[0].render(scale=2.0).to_pil()
image.save("cs_tel_250371509001.png")
print("Saved cs_tel_250371509001.png!")
