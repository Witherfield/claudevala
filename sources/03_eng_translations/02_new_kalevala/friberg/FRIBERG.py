import re

with open("FRIBERG.txt", "r", encoding="utf-8") as fin, \
     open("output.txt", "w", encoding="utf-8") as fout:

    for line in fin:
        text = re.sub(r"\s*\d+$", "", line.rstrip("\n"))

        if text.strip():      # Skip empty lines
            fout.write(text + "\n")
