import pytest
import os
from src.domain.services import pass_item_contract


data = {
    "it_worker": "Szymon Iwaniuk",
    "borrower": "Mike Wazowski",
    "id": "K123",
    "item": "Laptop Dell 12345AB",
    "quantity": "1",
    "date": "2025-08-11"
}


def test_pass_item_contract_content() -> None:
    pass

def test_pass_item_contract_save_path() -> None:
    output_path = pass_item_contract(data)
    
    assert os.path.exists(output_path)
    assert output_path.endswith(".docx")
    assert data["borrower"].replace(" ", "_") in output_path

    
print(pass_item_contract)