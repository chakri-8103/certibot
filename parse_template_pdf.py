import pypdf

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]

elements = []
def visitor(text, cm, tm, font_dict, font_size):
    t = text.strip()
    if t:
        elements.append({
            "x": round(tm[4], 2),
            "y": round(tm[5], 2),
            "size": round(font_size, 2),
            "text": t
        })

page.extract_text(visitor_text=visitor)

# Sort elements by Y descending (top to bottom in PDF coords) then X ascending
elements.sort(key=lambda item: (-item["y"], item["x"]))

for el in elements:
    print(f"y={el['y']:6.2f}, x={el['x']:6.2f}, size={el['size']:4.1f} | {el['text']!r}")
