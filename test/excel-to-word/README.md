# Excel to Word Document Generator

This tool reads rows from an Excel file (`.xlsx`) and populates a Word template (`.docx`) for each row using `docxtpl` (Jinja2 syntax).

---

## 1. Setup in Visual Studio Code (VS Code)

1. Open VS Code.
2. Open this folder in VS Code:
   - Go to **File -> Open Folder...** and select `/Users/perestal/.gemini/antigravity/scratch/excel-to-word`.
3. Open the integrated terminal in VS Code:
   - Shortcut: `Ctrl + ~` (or `Cmd + ~` on Mac) or menu **Terminal -> New Terminal**.
4. Create and activate a virtual environment (recommended):
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```
5. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   *(Or run `pip install docxtpl openpyxl` directly)*

6. Select the Python Interpreter in VS Code:
   - Press `Cmd + Shift + P` (Mac) or `Ctrl + Shift + P` (Windows/Linux).
   - Type and select: **Python: Select Interpreter**.
   - Choose the interpreter labeled `('.venv': venv)` or `./.venv/bin/python`.

---

## 2. Setting Up Your Word Template

In your `.docx` template, place variables inside double curly braces `{{ ... }}` matching your Excel column names:

- If your Excel column is **Full Name**, use: `{{ Full_Name }}`
- If your Excel column is **Address**, use: `{{ Address }}`
- If your Excel column is **Date**, use: `{{ Date }}`
- If your Excel column is **Amount**, use: `{{ Amount }}`

> **Tip:** You can keep all formatting in Word (fonts, colors, bold, tables, headers, footers). The placeholders will take on whatever style you apply to the `{{ ... }}` text in Word!

---

## 3. Running the Generator

Put your files in the folder (or provide their paths) and run:

```bash
python generate_docs.py --excel "your_data.xlsx" --template "your_template.docx"
```

### Optional Arguments:
- `--output` / `-o`: Specify output folder (default: `output/`).
- `--name-col` / `-n`: Specify an Excel column to use for generated document file names (e.g. `--name-col "Full_Name"`).

Example:
```bash
python generate_docs.py --excel clients.xlsx --template contract_template.docx --name-col Full_Name
```
