# Part Connector

## Setup Project
Once you forked and cloned the repo, run:
```bash
poetry install
```
to install dependencies.
Then write code in the src/ folder.

## Quality Check
To setup pre-commit hook (you only need to do this once):
```bash
poetry run pre-commit install
```
To manually run pre-commit checks:
```bash
poetry run pre-commit run --all-file
```
To manually run ruff check and auto fix:
```bash
poetry run ruff check --fix
```


## Building as Executable
To build the project as a standalone .exe:

Install dependencies (including PyInstaller):

poetry install
Build the executable:
```bash
poetry run pyinstaller --onefile --name itemList_part --hidden-import win32timezone --hidden-import win32com.client build_exe.py
```
The executable will be created in the dist folder.

The --hidden-import flags ensure PyInstaller includes the Windows COM dependencies needed for QuickBooks integration.

Running the Executable
After building, launch the CLI directly from Command Prompt:

Change into the dist directory (or reference the full path):
```bash
cd dist
```
Run the executable with the same arguments the Python entry point expects:
```bash
./itemList_part.exe --workbook C:\path\to\company_data.xlsx --output C:\path\to\report.json
```
--output is optional. If omitted it wil generate the json file in the current folder
If you omit --output, the report defaults to conflicts_output.json in the current directory.

## Example Output
```bash
{
    "status": "success",
    "generated_at": "2025-12-03T20:15:23.932381+00:00",
    "same_items": 5,
    "conflicts": [
        {
            "record_id": "5",
            "qb_name": "Nail",
            "excel_name": "S7214-09A-SQRT",
            "qb_price": 14.69,
            "excel_price": 0.42,
            "reason": "data_mismatch"
        },
        {
            "record_id": "S1a2i",
            "qb_name": "9001",
            "excel_name": null,
            "qb_price": 1234.56,
            "excel_price": null,
            "reason": "missing_in_excel"
        },
        {
            "record_id": "TEST04",
            "qb_name": "Bolt",
            "excel_name": null,
            "qb_price": 13.2,
            "excel_price": null,
            "reason": "missing_in_excel"
        },
        {
            "record_id": "TEST03",
            "qb_name": "Gadget",
            "excel_name": null,
            "qb_price": 10.99,
            "excel_price": null,
            "reason": "missing_in_excel"
        },
        {
            "record_id": "9",
            "qb_name": "hh",
            "excel_name": null,
            "qb_price": 99.99,
            "excel_price": null,
            "reason": "missing_in_excel"
        }
    ],
    "add_new_items": [
        {
            "id": "8",
            "name": "Surya",
            "price": 16.14
        }
    ],
    "error": null
}
```
