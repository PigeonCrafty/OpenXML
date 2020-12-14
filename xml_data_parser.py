import os
import re
import requests
import zipfile
from pathlib import Path
from bs4 import BeautifulSoup
import lxml

def list_files(dir_input):
    global pptx_ls
    global name_ls
    for item in os.scandir(dir_input):
        filePath = Path(item)
        fileName = str(filePath.stem)
        fileType = str(filePath.suffix.lower())
        # 添加更多条件以筛选XML(如正则表达式)，当前为路径下所有XML
        if fileType == ".xml":
            name_ls.append(filePath)
        elif fileType == ".pptx":
            pptx_ls.append(filePath)

def list_pptx(dir_input):
    pptx_ls = []
    for item in os.scandir(dir_input):
        filePath = Path(item)
        fileName = str(filePath.stem)
        fileType = str(filePath.suffix.lower())

def writeIn(text):
    global output_txt
    output_txt.write(str(text) + "\t")

def writeIfAttr(tag, tagPr):
    if tagPr in tag.attrs:
        output = tag.attrs[tagPr]
        writeIn(output)
    else:
        writeIn("")
        output = None
    return output

def get_slides(file_path):
    ls_slides = []
    if re.match("slide\d.xml", str(file_path)):
        ls_slides.append(file_path)

def soup_xml(filePath):
    with open(str(filePath), "r", encoding="utf-8") as file:
        xml = file.read()
    soup = BeautifulSoup(xml, "lxml")
    return soup

def if_write_print(name):
    if str(type(name)) == "<class 'pathlib.WindowsPath'>":
        writeIn(name.stem)                   # Write in XML name
        print("[Parsing]: " + name_ls[i].stem)
    elif str(type(name)) == "<class 'str'>":
        writeIn(name)
        print("[Parsing]: " + name_ls[i])

def get_textShapes(soup):
    shapes = soup.find_all("p:sp")
    shapes_txt = []
    for shape in shapes:
        if len(shape.find_all("a:t")) != 0 and shape.get_text() != None:
            shapes_txt.append(shape)
    return shapes_txt

def get_shapePr(shape):
    xfrm = shape.find("a:xfrm")
    if xfrm != None:
        off = xfrm.find("a:off")
        ext = xfrm.find("a:ext")
        x = int(writeIfAttr(off, "x"))
        y = int(writeIfAttr(off, "y"))
        cx = int(writeIfAttr(ext, "cx"))
        cy = int(writeIfAttr(ext, "cy"))
    else:
        for i in range(4):
            writeIn("")
    if xfrm != None and x != None:
        shape_width = cx-x
        shape_height = cy-y
        writeIn(abs(shape_width))
        writeIn(abs(shape_height))
    else:
        writeIn("")
        writeIn("")

def get_ranges(shape):
    ranges = shape.find_all("a:r")
    for r in ranges:
        if r.find("a:t") == None:
            ranges.remove(r)
    return ranges


# Get an input
user_input = Path(input(">>> Enter a dirctory or single pptx path: "))
print(">>> In progress...")
# Check if path exist
if Path.exists(user_input) != True:
    print("The path doesn't exist!")
    os._exit(0)

pptx_ls = []
name_ls = []
soup_ls = []     # xml object in soup


# Check the path is file to read, or a direcotry to traverse
if Path.is_file(user_input):
    if user_input.suffix == ".xml":
        name_ls.append(user_input)
    elif user_input.suffix == ".pptx":
        pptx_ls.append(user_input)
    else: 
        print("Error: Not an .xml or .pptx!")
        os._exit(0)
    output_txt = open(str(os.path.join(user_input.parent, user_input.name.rstrip(".xml") + "_data.txt")), "a+", encoding="utf-8")
elif Path.is_dir(user_input):
    list_files(user_input)
    output_txt = open(str(os.path.join(user_input.parent, user_input.name + "_data.txt")), "a+", encoding="utf-8")
else:
    print("Cannot process your input!")
    os._exit(0)

if len(pptx_ls) == 0:   # In XML mode
    for xml_path in name_ls:
        soup_ls.append(soup_xml(xml_path))
else:                   # In PPTX mode
    for pptx in pptx_ls:
        with zipfile.ZipFile(str(pptx), "r") as pz:
            for filename in pz.namelist():
                if re.match("ppt/slides/slide[0-9]*.xml", filename):        # Read slidesX.xml only in pptx
                    name_ls.append(str(pptx.name) + "\\" + filename)
                    content = pz.read(filename)
                    soup_ls.append(BeautifulSoup(content, "lxml"))
                    
# Write titles for txt
output_txt.write("XmlName" + "\t" + 
                 "ShapeID" + "\t" + # "Autofit" + "\t" + 
                 "X" + "\t" + "Y" + "\t" +  "cX" + "\t" + "cY" + "\t" + "ShapeWidth" + "\t" + "ShapeHeight" + "\t" + 
                 "TextSize" + "\t" + "TextLength" + "\t" +
                 "LatinTypeface" + "\t" + "LatinPanose" + "\t" + "LatinPitchFamily" + "\t" + "LatinCharset" + "\t" + 
                 "EaTypeface" + "\t" + "EaPanose" + "\t" + "EaPitchFamiliy" + "\t" + "EaCharset" + "\t" + 
                 "TextStrings" + "\n")

if len(name_ls) != len(soup_ls):
    print(">>> Warning: The numbers don't match!")

# Read every PPT -> xml -> shape ...
for i in range(len(soup_ls)):
    if_write_print(name_ls[i])
    shapes = get_textShapes(soup_ls[i])   # Get a list of all shapes with texts
    shape_num = 0
    for shape in shapes:
        if shape_num > 0: 
            if_write_print(name_ls[i])
        spId = shape.find("p:cnvpr")
        writeIn(spId["id"])             # Write in ShapeID
        # if shape.find("a:normAutofit") != None or shape.find("a:spAutoFit") != None:
        #     writeIn("Yes")              # Write in Autofit
        # else:   
        #     writeIn("No")
        get_shapePr(shape)         # Write in ShapePr(x, y, cx, cy)
        ranges = get_ranges(shape)      # Get a list of all ranges of texts
        if len(ranges) == 0:
            for l in range(16):
                writeIn("")
            continue
        rng = ranges[0]
        rngPr = rng.find("a:rpr")
        text_size = writeIfAttr(rngPr, "sz")        # Write in Size
        output_txt.write(str(len(shape.get_text())) + "\t")

        latin = rngPr.find("a:latin")
        if latin != None:
            latin_typeface = writeIfAttr(latin, "typeface")
            latin_panose = writeIfAttr(latin, "panose")
            latin_pf = writeIfAttr(latin, "pitchFamily")
            latin_charset = writeIfAttr(latin, "charset")
        else:
            for j in range(4):
                writeIn("")
        
        ea = rngPr.find("a:ea")
        if ea != None:
            ea_typeface = writeIfAttr(ea, "typeface")
            ea_panose = writeIfAttr(ea, "panose")
            ea_pf = writeIfAttr(ea, "pitchFamily")
            ea_charset = writeIfAttr(ea, "charset")
        else:
            for k in range(4):
                writeIn("")
        output_txt.write(shape.get_text() + "\n")
        shape_num += 1

output_txt.close()
print(">>> All done :)")