import datetime
import os
from docx import Document
from typing import List
from src.domain.helper import save_as_pdf, set_cell_border, resource_path


def pass_item_contract(it_worker: str, borrower: str, id: str, item: str, quantity: str, date=None) -> str:
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

    template_path = resource_path(os.path.join("templates", "pass_item_template.docx"))
    doc = Document(template_path)

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

    filename = f"Przekazanie_sprzetu_{borrower.replace(' ', '_')}_{date}.docx"
    save_path = os.path.join(r"C:\docx_writer\attachments", filename)

    # Save document
    doc.save(save_path)

    # Conversion to pdf
    pdf_path = save_as_pdf(save_path)
    return pdf_path


def change_item_contract(
    it_worker: str,
    borrower: str,
    take_id: str,
    take_item: str,
    take_qty: str,
    give_id: str,
    give_item: str,
    give_qty: str,
    date=None,
) -> str:

    template_path = resource_path(os.path.join("templates", "change_item_template.docx"))
    doc = Document(template_path)

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

    filename = f"Wymiana_sprzetu_{borrower.replace(' ', '_')}_{date}.docx"
    save_path = os.path.join(r"C:\docx_writer\attachments", filename)

    # Save document
    doc.save(save_path)

    # Conversion to pdf
    pdf_path = save_as_pdf(save_path)
    return pdf_path


def utilization_items_contract(
    items: list,
    participants: List[str],
    date=None,
) -> str:
    template_path = resource_path(os.path.join("templates", "utilization_items_template.docx"))
    doc = Document(template_path)

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

    filename = f"Utylizacja_sprzetu_{date}.docx"
    save_path = os.path.join(r"C:\docx_writer\attachments", filename)

    # Save document
    doc.save(save_path)

    # Conversion to pdf
    pdf_path = save_as_pdf(save_path)
    return pdf_path