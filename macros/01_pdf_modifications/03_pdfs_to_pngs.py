import os
from pdf2image import convert_from_path

CURRENT_FOLDER = os.path.dirname(os.path.realpath(__file__))
INPUT_FOLDER = os.path.join(CURRENT_FOLDER, "output")
OUTPUT_FOLDER = os.path.join(CURRENT_FOLDER, "png")

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

for filename in sorted(os.listdir(INPUT_FOLDER)):
    if not filename.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(INPUT_FOLDER, filename)

    # One-page PDF
    image = convert_from_path(pdf_path, dpi=300)[0]

    png_name = os.path.splitext(filename)[0] + ".png"
    png_path = os.path.join(OUTPUT_FOLDER, png_name)

    image.save(png_path, "PNG")

    print(f"Saved {png_name}")