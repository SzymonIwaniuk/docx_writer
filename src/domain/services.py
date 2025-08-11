import datetime
import os
from typing import Optional

from docx import Document


def fill_docx_template(it_worker: str, borrower: str, date: Optional[str], id: str, item: str, quantity: str) -> str:
    # Temporary hardcoded
    doc = Document(r"C:\Users\admin\Desktop\pliki\docx_writer\src\templates\pass_item_template.docx")

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
    print(f"Document saved to: {save_path}")
    return save_path


if __name__ == "__main__":

    fill_docx_template(
        it_worker="Szymon Iwaniuk",
        borrower="Mike Wazowski",
        date="2025-08-11",
        id="K123",
        item="Dell Laptop",
        quantity="1",
    )
