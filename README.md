# AutoMachine Recipe Manager V.2

A desktop GUI application for managing tea drink recipes and exporting them to Excel and QR code formats for automated machines.

Built with Python + Tkinter.

---

## Features

- **Recipe editor** — create and edit recipes with size (16oz / 22oz), ice level, sugar level, and per-ingredient usage amounts
- **Multi-market support** — organise recipes by country/market tag (Thai, Singapore, Australia, Spain, USA) and recipe group
- **Excel import / export** — load and save recipe workbooks (`.xlsx`) with auto-formatted tables
- **QR code export** — generate a finalized QR string sheet (`Finalized QR String`) mapping each recipe row to SI codes for the machine
- **Ingredient master list** — manage a grouped ingredient list (J\_ / H\_ prefixes) shared across all recipes
- **Code mapping manager** — configure Sugar, Ice, and Size → SI code mappings used in QR export
- **Groups & Tags manager** — add or remove country tags and the recipe groups that belong to them
- **Validation panel** — warns about unknown ingredients, missing usage values, or non-numeric entries
- **Drag-to-reorder** — reorder recipes in the sidebar by dragging
- **Undo / Redo** — up to 30 levels of undo

---

## Requirements

- Python 3.8+
- [openpyxl](https://openpyxl.readthedocs.io/)

```
pip install openpyxl
```

---

## Usage

```
python AutoMachine_Recipe_Manager.py
```

---

## Project Files

| File | Description |
|---|---|
| `AutoMachine_Recipe_Manager.py` | Main application |
| `ingredients.json` | Ingredient master list (auto-saved) |
| `groups_tags.json` | Country tags and recipe groups (auto-saved) |
| `code_mappings.json` | SI code mappings for Sugar / Ice / Size (auto-saved on first edit) |

---

## QR String Format

Each exported row produces a string in this format:

```
ORDERXXXX|<RecipeCode>|<SizeCode>,<IceCode>,<SugarCode>|
```

Rows are sorted by SI code number ascending (SI0001 → SI0006…).
