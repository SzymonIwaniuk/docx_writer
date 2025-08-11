from docx import Document
import os
import datetime


def fill_docx_template(it_worker, borrower, date, id, item, quantity):
    # Temporary hardcoded
    doc = Document(r"C:\Users\admin\Desktop\pliki\docx_writer\src\templates\pass_item_template.docx")


    replacements = {
        "{{it_worker}}": it_worker,
        "{{borrower}}": borrower,
        "{{date}}": date,
        "{{id}}": id,
        "{{item}}": item,
        "{{quantity}}": quantity
    }

    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for placeholder, value in replacements.items():
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, value)

     # Save path
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = f"Laptop_Contract_{borrower.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    save_path = os.path.join(desktop, filename)

    # Save document
    doc.save(save_path)
    print(f"Document saved to: {save_path}")
    return save_path


if __name__ == '__main__':
        
    fill_docx_template(
        it_worker="Szymon Iwaniuk",
        borrower="Jan Kowalski",
        date="2025-08-11",
        id="REQ-2025-001",
        item="Monitor",
        quantity="2",
    )
