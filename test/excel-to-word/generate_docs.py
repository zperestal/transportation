import os
import re
import sys
import argparse
from datetime import datetime, date
import openpyxl
from docxtpl import DocxTemplate

def sanitize_variable_name(header_name: str) -> str:
    """
    Sanitize header name into a valid python/jinja variable identifier.
    e.g. 'First Name' -> 'First_Name', 'Total ($)' -> 'Total'
    """
    if header_name is None:
        return ""
    # Strip whitespace
    name = str(header_name).strip()
    # Replace whitespace and hyphens with underscores
    name = re.sub(r'[\s\-]+', '_', name)
    # Remove characters that aren't alphanumeric or underscore
    name = re.sub(r'[^\w]', '', name)
    # If starting with a digit, prepend an underscore
    if name and name[0].isdigit():
        name = f"_{name}"
    return name

def format_cell_value(val):
    """
    Formats dates and None values cleanly for Word documents.
    """
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    return val

def generate_documents(
    excel_path: str,
    template_path: str,
    output_dir: str = "output",
    filename_column: str = None
):
    """
    Reads rows from the Excel file and generates a Word doc for each row using docxtpl.
    """
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found: {excel_path}")
        sys.exit(1)

    if not os.path.exists(template_path):
        print(f"Error: Template file not found: {template_path}")
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    print(f"Loading Excel file: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb.active
    print(f"Active Sheet: '{sheet.title}'")

    # Read headers from row 1
    headers = [cell.value for cell in sheet[1]]
    if not any(headers):
        print("Error: No headers found in row 1 of the Excel sheet.")
        sys.exit(1)

    # Map column index to variable names (both original and sanitized)
    col_mapping = {}
    for idx, header in enumerate(headers):
        if header is not None:
            raw_key = str(header).strip()
            clean_key = sanitize_variable_name(raw_key)
            col_mapping[idx] = (raw_key, clean_key)

    print("\n--- Available Variables for Word Template ---")
    for idx, (raw_k, clean_k) in col_mapping.items():
        if raw_k == clean_k:
            print(f"  Column {idx + 1}: {{{{ {clean_k} }}}}")
        else:
            print(f"  Column {idx + 1}: {{{{ {clean_k} }}}}  (from '{raw_k}')")
    print("---------------------------------------------\n")

    # Process data rows starting from row 2
    row_count = 0
    generated_count = 0

    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        # Skip completely empty rows
        if not any(val is not None and str(val).strip() != "" for val in row):
            continue

        row_count += 1
        context = {}
        for col_idx, (raw_key, clean_key) in col_mapping.items():
            val = row[col_idx] if col_idx < len(row) else None
            formatted_val = format_cell_value(val)
            # Provide both sanitized key and raw key in context for flexibility
            context[clean_key] = formatted_val
            context[clean_key.lower()] = formatted_val
            if raw_key != clean_key:
                context[raw_key] = formatted_val

        # Determine output file name
        base_name = None
        if filename_column:
            clean_fn_col = sanitize_variable_name(filename_column)
            base_name = context.get(clean_fn_col) or context.get(filename_column)

        if not base_name:
            # Fallback: check first column or use row index
            first_val = next((context.get(clean) for _, clean in col_mapping.values() if context.get(clean)), None)
            base_name = str(first_val) if first_val else f"document_row_{row_idx}"

        # Clean filename of illegal filesystem characters
        safe_base_name = re.sub(r'[\\/*?:"<>|]', '_', str(base_name)).strip()
        output_filename = f"{safe_base_name}_{row_idx}.docx"
        output_filepath = os.path.join(output_dir, output_filename)

        # Render template
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_filepath)
        generated_count += 1
        print(f"[{generated_count}] Created: {output_filepath}")

    print(f"\nDone! Successfully generated {generated_count} documents in '{output_dir}/'.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fill Word templates from Excel data.")
    parser.add_argument(
        "--excel", "-e",
        default="data.xlsx",
        help="Path to Excel data file (.xlsx) (default: data.xlsx)"
    )
    parser.add_argument(
        "--template", "-t",
        default="template.docx",
        help="Path to Word template file (.docx) (default: template.docx)"
    )
    parser.add_argument(
        "--output", "-o",
        default="output",
        help="Output directory path (default: output)"
    )
    parser.add_argument(
        "--name-col", "-n",
        default=None,
        help="Excel column header to use for naming output files (e.g. 'Name' or 'Client_ID')"
    )

    args = parser.parse_args()
    generate_documents(
        excel_path=args.excel,
        template_path=args.template,
        output_dir=args.output,
        filename_column=args.name_col
    )
