import os

import pytest
from docx import Document

from src.domain.services import pass_item_contract


def test_pass_item_contract_content() -> None:
    data = {
        "it_worker": "Szymon Iwaniuk",
        "borrower": "Mike Wazowski",
        "id": "K123",
        "item": "Laptop Dell 12345AB",
        "quantity": "1",
        "date": "2025-08-11",
    }
    doc = Document(pass_item_contract(**data))

    content = "\n".join([p.text for p in doc.paragraphs])

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                content += "\n" + cell.text

    for value in data.values():
        assert value in content


def test_pass_item_contract_save_path() -> None:
    data = {
        "it_worker": "Szymon Iwaniuk",
        "borrower": "Mike Wazowski",
        "id": "K123",
        "item": "Laptop Dell 12345AB",
        "quantity": "1",
        "date": "2025-08-11",
    }

    output_path = pass_item_contract(**data)

    assert os.path.exists(output_path)
    assert output_path.endswith(".docx")
    assert data["borrower"].replace(" ", "_") in output_path
