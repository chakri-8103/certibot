import pypdf
import re

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]

contents_obj = page["/Contents"].get_object()

if isinstance(contents_obj, pypdf.generic.ArrayObject):
    print(f"Page /Contents has {len(contents_obj)} stream objects.")
    for idx, stream_ref in enumerate(contents_obj):
        stream_data = stream_ref.get_object().get_data()
        pattern = re.compile(b'BT[\\s\\S]*?ET')
        bt_matches = list(pattern.finditer(stream_data))
        print(f"Stream {idx}: {len(bt_matches)} BT...ET matches")
        
        if len(bt_matches) > 0:
            dynamic_indices = {1, 5, 6, 8, 11, 16, 18, 21, 22, 23, 24, 25, 27, 29, 32, 33, 34, 35, 36, 37, 38}
            new_bytes = bytearray()
            last_end = 0
            for i, m in enumerate(bt_matches):
                start, end = m.span()
                new_bytes.extend(stream_data[last_end:start])
                if i in dynamic_indices:
                    pass
                else:
                    new_bytes.extend(stream_data[start:end])
                last_end = end
            new_bytes.extend(stream_data[last_end:])
            stream_ref.get_object().set_data(bytes(new_bytes))

writer = pypdf.PdfWriter()
writer.add_page(page)

with open("byte_clean_template.pdf", "wb") as f:
    writer.write(f)

print("Saved byte_clean_template.pdf!")
