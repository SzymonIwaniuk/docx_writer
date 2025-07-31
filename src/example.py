import os

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

from datetime import datetime
from typing import Optional


def create_laptop_borrow_contract(
        name: str, 
        date: str,
        laptop_model: str,
        return_date: str,
        template_path: Optional[str] = None,
        save_path: Optional[str] = None,
) -> str:

    # Load template or create new doc
    if template_path and os.path.exists(template_path):
        doc = Document(template_path)
    else:
        doc = Document()
    
    # Add logo on the top and center it
    logo = doc.add_heading("", 0)
    logo_color = logo.add_run("SaMASZ Sp z.o.o")
    logo_color.font.color.rgb = RGBColor(0, 128, 0)
    logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Laptop borrowing contract title
    contract = doc.add_heading("", 2)
    text_contract = contract.add_run("Laptop Borrowing Contract")
    text_contract.font.color.rgb = RGBColor(0, 0, 0)
    text_contract.bold = True
    contract.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Document context
    doc.add_paragraph(f"Date: {date}")
    doc.add_paragraph(f"Borrower's Name: {name}")
    doc.add_paragraph(f"Laptop Model: {laptop_model}")
    doc.add_paragraph(f"Expected Return Date: {return_date}")
    
    # Contract formula
    doc.add_paragraph(
        "I acknowledge receipt of the above-listed laptop and agree to return it in good condition by the return date. "
        "I accept responsibility for loss or damage to the equipment."
    )
    
    # Signature and date
    doc.add_paragraph("\n\nSignature: _________________________")
    doc.add_paragraph("Date Signed: _________________________")

    # Save path
    if not save_path:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        filename = f"Laptop_Contract_{name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        save_path = os.path.join(desktop, filename)
    
    # Save document
    doc.save(save_path)
    print(f"Document saved to: {save_path}")
    return save_path


if __name__ == "__main__":
    create_laptop_borrow_contract(
        name="Szymon Iwaniuk",
        date="2025-07-31",
        laptop_model="Dell Latitude 7390 2-in-1",
        return_date="2025-08-31",
        template_path=None,
        save_path=None       
    )

