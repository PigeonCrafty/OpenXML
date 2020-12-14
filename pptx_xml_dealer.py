import os
import re
import requests
import zipfile
from pathlib import Path
# from bs4 import BeautifulSoup
# import lxml
# import xml.dom.minidom
# from xml.etree.ElementTree import parse, Element
import xml.etree.ElementTree as ET

# 使用ET解析XML必须先注册命名空间！
ET.register_namespace('', "http://www.w3.org/2001")
ET.register_namespace('a', "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace('p', "http://schemas.openxmlformats.org/presentationml/2006/main")
ET.register_namespace('a16', "http://schemas.microsoft.com/office/drawing/2014/main")
ET.register_namespace('p14', "http://schemas.microsoft.com/office/powerpoint/2010/main")

# 临时文件用于测试功能
pptx = "c:\\Users\\pigeonz.CZ\\OneDrive - RWS\\AI\\openXML\\XML\\simple\\2.pptx"
# xml_name = "c:\\Users\\pigeonz.CZ\\OneDrive - RWS\\AI\\openXML\\XML\\simple\\2.xml"
xml_temp = "c:/Users/pigeonz.CZ/OneDrive - RWS/AI/openXML/XML/simple/2"
xml_file = "c:\\Users\\pigeonz.CZ\\OneDrive - RWS\\AI\\openXML\\XML\\simple\\2\\ppt\\slides\\slide2.xml"
xml_file2 = "c:\\Users\\pigeonz.CZ\\OneDrive - RWS\\AI\\openXML\\XML\\simple\\2.2.xml"

# # 必须先解压出XML再进行解析处理，目前尚未找到writestr()替换压缩包内文件的方法！
# with zipfile.ZipFile(str(pptx), "r") as z:
#     for fileName in z.namelist():
#         # print(fileName)
#         if re.match("ppt/slides/slide[0-9]*.xml", fileName):
#             # xml_zfile = z.read(fileName)
#             z.extract(fileName, xml_temp)


"""
    tag = element.text                  访问Element标签
    attrib = element.attrib             访问Element属性
    text = element.text                 访问Element文本

    Element.text = ''                   直接改变字段内容
    Element.append(Element)             为当前的Elment对象添加子对象
    Element.remove(Element)             删除Element节点
    Element.set(key, value)             添加和修改属性
    ElementTree.write('filename.xml')   写出（更新）XMl文件
"""
# 使用ET无法正确解析OpenXML？
file = open(xml_file2, encoding="utf-8", mode="r")
xml = file.read()
tree = ET.fromstring(xml)
root = tree.getroot()
print(root.tag)
print(root.text)
print(root.attrib)
file.close()

rpr = root.find("a:rpr")
text = root.find("a:t")
print("!")

# tree.write(xml, encoding="utf-8", xml_declaration=True)





            # # NOT THE BEST WAY!!
            # dom = xml.dom.minidom.parse(xml_zfile)
            # collection = dom.documentElement
            # ranges = collection.getElementsByTagName("a:r")
            # for r in ranges:
            #     rpr = r.firstChild
            #     t = r.lastChild
            #     if rpr.hasAttribute("b"):
            #         rpr.setAttribute("b", "0")
            # print(str(dom.text))
            # print("Done")

            # SOUP DOESN'T WORK!!!
            # soup = BeautifulSoup(xml_texts, "lxml")
            # texts = soup.find_all("a:t")
            # for t in texts:
            #     rpr = t.find_previous_sibling("a:rpr")
            #     if "b" in rpr.attrs and rpr["b"] == '1':
            #         del rpr["b"]
            # print(soup.prettify())
            
            # z.writestr(fileName, str(soup))