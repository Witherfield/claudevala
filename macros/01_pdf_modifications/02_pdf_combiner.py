import os
from pypdf import PdfReader, PdfWriter

CURRENT_FOLDER = os.path.dirname(os.path.realpath(__file__))
INPUT_FOLDER = os.path.join(CURRENT_FOLDER, "output")
OUTPUT_PDF = os.path.join(CURRENT_FOLDER, "COMBINED.pdf")

pdf_files = sorted(
    f for f in os.listdir(INPUT_FOLDER)
    if f.lower().endswith(".pdf")
)

writer = PdfWriter()

for pdf_file in pdf_files:
    pdf_path = os.path.join(INPUT_FOLDER, pdf_file)
    reader = PdfReader(pdf_path)

    for page in reader.pages:
        writer.add_page(page)

with open(OUTPUT_PDF, "wb") as f:
    writer.write(f)

print(f"Combined {len(pdf_files)} PDFs into {OUTPUT_PDF}.")