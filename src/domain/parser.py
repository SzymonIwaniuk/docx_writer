import argparse
from src.domain.services import pass_item_contract, change_item_contract, utilization_items_contract    


## Helper functions
# Items parser to utilization due to erors with handling json via powershell
def parse_items(arg: str):
    items = []
    for part in arg.split(";"):
        part = part.strip()
        if not part:
            continue
        # fields: id,name,inventarization_num,date
        id_, name, inv, date = part.split(",")
        items.append({
            "id": id_,
            "name": name,
            "inventarization_num": inv,
            "date": date
        })
    return items

def parse_participants(arg: str):
    return [p.strip() for p in arg.split(";") if p.strip()]


def parser():
    parser = argparse.ArgumentParser(description="Generate IT equipment contracts.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # pass_item_contract subcommand
    pass_parser = subparsers.add_parser("pass_item")
    pass_parser.add_argument("--it_worker", required=True)
    pass_parser.add_argument("--borrower", required=True)
    pass_parser.add_argument("--id", required=True)
    pass_parser.add_argument("--item", required=True)
    pass_parser.add_argument("--quantity", required=True)
    pass_parser.add_argument("--date", required=False)

    # change_item_contract subcommand
    change_parser = subparsers.add_parser("change_item")
    change_parser.add_argument("--it_worker", required=True)
    change_parser.add_argument("--borrower", required=True)
    change_parser.add_argument("--take_id", required=True)
    change_parser.add_argument("--take_name", required=True)
    change_parser.add_argument("--take_qty", required=True)
    change_parser.add_argument("--give_id", required=True)
    change_parser.add_argument("--give_name", required=True)
    change_parser.add_argument("--give_qty", required=True)
    change_parser.add_argument("--date", required=False)

    # utilization_items_contract subcommand
    utilization_parser = subparsers.add_parser("utilization")

    utilization_parser.add_argument(
        "--items",
        type=parse_items,
        required=True
    )

    utilization_parser.add_argument(
        "--participants",
        type=parse_participants,
        required=True
    )

    utilization_parser.add_argument("--date", required=False)

    args = parser.parse_args()

    if args.command == "pass_item":
        pass_item_contract(
            args.it_worker,
            args.borrower,
            args.id,
            args.item,
            args.quantity,
            args.date,
        )

    elif args.command == "change_item":
        change_item_contract(
            args.it_worker,
            args.borrower,
            args.take_id,
            args.take_name,
            args.take_qty,
            args.give_id,
            args.give_name,
            args.give_qty,
            args.date,
        )
    
    elif args.command == "utilization":
        utilization_items_contract(
            args.items,
            args.participants,
            args.date,
        )

if __name__ == '__main__':
    parser()