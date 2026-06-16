import os
from PIL import Image
import pytesseract

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

print(pytesseract.get_languages(config=''))

CURRENT_FOLDER = os.path.dirname(os.path.realpath(__file__))
INPUT_FOLDER = os.path.join(CURRENT_FOLDER, "png")
OUTPUT_FOLDER = os.path.join(CURRENT_FOLDER, "txt")

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

for filename in sorted(os.listdir(INPUT_FOLDER)):
    if not filename.lower().endswith(".png"):
        continue

    image_path = os.path.join(INPUT_FOLDER, filename)

    text = pytesseract.image_to_string(
        Image.open(image_path),
        lang="pol+fin"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        os.path.splitext(filename)[0] + ".txt"
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"OCR: {filename}")