import pypdf

reader = pypdf.PdfReader("250371509001.pdf")
page = reader.pages[0]

contents = page.get_contents()
if isinstance(contents, pypdf.generic.ArrayObject):
    print(f"page.get_contents() is an Array of {len(contents)} streams!")
    for i, stream_obj in enumerate(contents):
        data = stream_obj.get_data().decode('latin1', errors='ignore')
        print(f"\n--- Stream {i} (length {len(data)}) ---")
        print(data[:300])
else:
    data = contents.get_data().decode('latin1', errors='ignore')
    print(f"page.get_contents() is a Single stream of length {len(data)}")
    print(data[:300])
