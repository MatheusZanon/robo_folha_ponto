from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar, LAParams
import xml.etree.ElementTree as ET

def pdf_to_xml(pdf_file):
    root = ET.Element("Document")
    for page_layout in extract_pages(pdf_file):
        page_elem = ET.SubElement(root, "Page")
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                para_elem = ET.SubElement(page_elem, "Paragraph")
                for text_line in element:
                    line_elem = ET.SubElement(para_elem, "Line")
                    line_elem.text = text_line.get_text()
    return ET.ElementTree(root)