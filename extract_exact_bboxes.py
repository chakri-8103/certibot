import pypdf

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]
page_height = float(page.mediabox.height)
page_width = float(page.mediabox.width)

text_elements = []

def visitor(text, cm, tm, font_dict, font_size):
    t = text.strip()
    if not t:
        return
    # Combined transformation matrix: CTM * TM
    # cm: [a, b, c, d, e, f]
    # tm: [a, b, c, d, e, f]
    # Position in user space:
    x_user = cm[0] * tm[4] + cm[2] * tm[5] + cm[4]
    y_user = cm[1] * tm[4] + cm[3] * tm[5] + cm[5]
    
    # Convert Y user (from bottom) to PDF Y (0 at bottom) or top-down Y (0 at top)
    y_top = page_height - y_user
    
    # Calculate effective font size
    scale_y = (cm[1]**2 + cm[3]**2)**0.5 if cm else 1.0
    effective_size = font_size * scale_y
    
    text_elements.append({
        "text": t,
        "x": round(x_user, 2),
        "y_bottom": round(y_user, 2),
        "y_top": round(y_top, 2),
        "size": round(effective_size, 2)
    })

page.extract_text(visitor_text=visitor)

# Sort from top of page to bottom
text_elements.sort(key=lambda item: item["y_top"])

print(f"--- RENDERED TEXT LOCATIONS (Page Height={page_height}) ---")
for el in text_elements:
    print(f"y_top={el['y_top']:6.2f} (y_bottom={el['y_bottom']:6.2f}), x={el['x']:6.2f}, size={el['size']:4.1f} | {el['text']!r}")
