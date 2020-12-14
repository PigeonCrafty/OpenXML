import zipfile
import re
from pathlib import Path
import os
from bs4 import BeautifulSoup
import lxml

pptx = Path("c:/Users/pigeonz.CZ/OneDrive - RWS/AI/openXML/XML_compare/welcome/Core Why Indeed Pitch Deck.pptx")
with zipfile.ZipFile(str(pptx), "r") as z:
    for filename in z.namelist():
        if re.match("ppt/slides/slide[0-9]*.xml", filename):
            # slide_xmls.append(filename)
            xml = z.read(filename)
            output_file = str(os.path.join(pptx.parent, "slides_data.xml"))
            output = open(output_file, "w", encoding="utf-8")
            output.write(str(xml, encoding="utf-8"))
            output.close()

with open(output_file, "r", encoding="utf-8") as file:
    soup = BeautifulSoup(file.read(), "lxml")
    print(soup.getText())