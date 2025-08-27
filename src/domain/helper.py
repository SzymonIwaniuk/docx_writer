import os
import sys

from docx2pdf import convert
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# Convert vile from docx to pdf via Aspose Words
def save_as_pdf(docx_path: str) -> str:
    """Convert DOCX to PDF using Word (Windows only)."""
    pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, pdf_path)
    return pdf_path


# Helper function for set borders of table
def set_cell_border(cell, size="4", color="000000"):
    """
    Set cell borders to 1 pt.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for border_name in ["top", "left", "bottom", "right"]:
        element = OxmlElement(f"w:{border_name}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        tcBorders = tcPr.find(qn("w:tcBorders"))

        if tcBorders is None:
            tcBorders = OxmlElement("w:tcBorders")
            tcPr.append(tcBorders)

        tcBorders.append(element)


# Due to development process and boundle everything into exe
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller exe."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
