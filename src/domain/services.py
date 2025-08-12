import datetime
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx import Document
from typing import List


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


def pass_item_contract(it_worker: str, borrower: str, id: str, item: str, quantity: str, date = None) -> str:

    """
    Generates a Word document contract for passing IT equipment to a borrower.

    This function loads a predefined Word template, replaces placeholders with provided values,
    and saves the filled document to the user's desktop with a filename based on the borrower's name and date.

    Args:
        it_worker (str): Name of the IT worker handing over the item.
        borrower (str): Name of the person receiving the item.
        date (Optional[str]): Date of the handover. If None, defaults to today's date in YYYY-MM-DD format.
        id (str): Identifier for the transaction or item.
        item (str): Description of the item being handed over.
        quantity (str): Quantity of the item being handed over.

    Returns:
    str: Path of saved docx file.
    """

    # Temporary hardcoded TODO
    doc = Document(r"src\templates\pass_item_template.docx")

    if date is None:
        date = datetime.datetime.today().strftime("%Y-%m-%d")

    replacements = {
        "{{it_worker}}": it_worker,
        "{{borrower}}": borrower,
        "{{date}}": date,
        "{{id}}": id,
        "{{item}}": item,
        "{{quantity}}": quantity,
    }

    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for placeholder, value in replacements.items():
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, value)

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = f"Przekazanie_sprzetu_{borrower.replace(' ', '_')}_{date}.docx"
    save_path = os.path.join(desktop, filename)

    # Save document
    doc.save(save_path)
    return save_path


def change_item_contract(
    it_worker: str,
    borrower: str,
    take_id: str,
    take_item: str,
    take_qty: str,
    give_id: str,
    give_item: str,
    give_qty: str,
    date = None
) -> str:
    doc = Document(r"src\templates\change_item_template.docx")

    if date is None:
        date = datetime.datetime.today().strftime("%Y-%m-%d")

    replacements = {
        "{{it_worker}}": it_worker,
        "{{borrower}}": borrower,
        "{{take_id}}": take_id,
        "{{take_item}}": take_item,
        "{{take_qty}}": take_qty,
        "{{give_id}}": give_id,
        "{{give_item}}": give_item,
        "{{give_qty}}": give_qty,
        "{{date}}": date,
    }

    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for placeholder, value in replacements.items():
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, value)

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = f"Wymiana_sprzetu_{borrower.replace(' ', '_')}_{date}.docx"
    save_path = os.path.join(desktop, filename)

    # Save document
    doc.save(save_path)
    return save_path

    
def utilization_items_contract(
    items: list,
    participants: List[str],
    date=None,
) -> str:
    doc = Document(r"src\templates\utilization_items_template.docx")

    if date is None:
        date = datetime.datetime.today().strftime("%Y-%m-%d")
    
    participants_section = ""
    for name in participants:
        participants_section += f"{name} " + "\n" + "." * 40 + "\n"

    replacements = {
        "{{participants_section}}": participants_section,
        "{{date}}": date,
    }

    # Replace with provided data
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for placeholder, value in replacements.items():
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, value)

    if doc.tables:
        table = doc.tables[0]
        for item in items:
            row = table.add_row()
            row.cells[0].text = item.get("id", "")
            row.cells[1].text = item.get("name", "")
            row.cells[2].text = item.get("inventarization_num", "")
            row.cells[3].text = item.get("date", "")
            for cell in row.cells:
                set_cell_border(cell, size="4", color="000000")

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = f"Utylizacja_sprzetu_{date}.docx"
    save_path = os.path.join(desktop, filename)

    # Save document
    doc.save(save_path)
    return save_path
  