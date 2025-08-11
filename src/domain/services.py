import datetime
import os

from docx import Document


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

