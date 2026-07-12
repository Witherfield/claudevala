import os
from pathlib import Path
from pypdf import PdfReader, PdfWriter

CURRENT_FOLDER = Path(os.path.dirname(os.path.realpath(__file__)))
INPUT_FILE = os.path.join(CURRENT_FOLDER, "1974.pdf")
OUTPUT_DIR = os.path.join(CURRENT_FOLDER, "output")

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

reader = PdfReader(INPUT_FILE)

for page_num, page in enumerate(reader.pages, start=1):
    writer = PdfWriter()
    writer.add_page(page)

    output_file = os.path.join(OUTPUT_DIR, f"{page_num:03d}.pdf")

    with open(output_file, "wb") as f:
        writer.write(f)

print(f"Created {len(reader.pages)} PDF files.")
