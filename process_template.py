"""
A script to process a Word document template, replace placeholders, and convert it to PDF.
"""
import re
import platform
import datetime
from docx import Document
from docx2pdf import convert

def find_placeholders(template_path):
    """
    Finds all placeholders (prefixed with 'v_') in a Word document.
    """
    doc = Document(template_path)
    placeholders = set()

    # Regular expression to find placeholders like v_name, v_date, etc.
    placeholder_regex = re.compile(r'v_[a-zA-Z0-9_]+')

    # Search in paragraphs
    for para in doc.paragraphs:
        found = placeholder_regex.findall(para.text)
        if found:
            for item in found:
                placeholders.add(item)

    # Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                found = placeholder_regex.findall(cell.text)
                if found:
                    for item in found:
                        placeholders.add(item)

    return list(placeholders)

def replace_placeholders_and_convert(template_path, data, output_docx, output_pdf):
    """
    Replaces placeholders in a Word document and converts it to PDF.

    Args:
        template_path (str): The path to the Word template.
        data (dict): A dictionary of placeholder keys and their values.
        output_docx (str): The path to save the processed .docx file.
        output_pdf (str): The path to save the converted .pdf file.
    """
    # First, validate that the user has provided data for all placeholders
    placeholders_in_doc = find_placeholders(template_path)
    missing_keys = [p for p in placeholders_in_doc if p not in data]
    if missing_keys:
        raise ValueError(f"Missing data for the following placeholders: {', '.join(missing_keys)}")

    doc = Document(template_path)

    # --- Placeholder Replacement ---
    # NOTE: This replacement logic is simple and works for placeholders that are
    # contained within a single "run" of text in the Word document. A run is a
    # contiguous stretch of text with the same formatting. If a placeholder is
    # split across different formatting runs (e.g., 'v_user' with 'v_' in bold
    # and 'user' not), this logic will not replace it. A more robust solution
    # would require parsing the underlying XML of the document, which is
    # significantly more complex.

    # Replace in paragraphs
    for para in doc.paragraphs:
        for key, value in data.items():
            if key in para.text:
                for run in para.runs:
                    run.text = run.text.replace(key, str(value))

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in data.items():
                        if key in para.text:
                            for run in para.runs:
                                run.text = run.text.replace(key, str(value))

    doc.save(output_docx)
    print(f"Saved modified document to {output_docx}")

    # Convert to PDF
    # On Linux, this requires LibreOffice to be installed.
    # On Windows, this requires Microsoft Word to be installed.
    try:
        convert(output_docx, output_pdf)
        print(f"Converted {output_docx} to {output_pdf}")
    except Exception as e:
        print(f"Error during PDF conversion: {e}")
        if platform.system() == "Linux":
            print("Please ensure LibreOffice is installed and accessible in the system's PATH.")
        elif platform.system() == "Windows":
            print("Please ensure Microsoft Word is installed.")


if __name__ == '__main__':
    template_file = 'one_year_template.docx'
    output_docx_file = 'one_year_processed.docx'
    output_pdf_file = 'one_year_processed.pdf'

    # --- Sample Usage ---

    # 1. Find all available placeholders in the template
    placeholders = find_placeholders(template_file)
    print("Found placeholders in template:", placeholders)

    # 2. Prepare the data for replacement
    #    In a real application, this data would come from a database, API, or user input.
    sample_data = {
        'v_issued_date': datetime.date.today().strftime("%Y-%m-%d"),
        'v_issued_time': datetime.datetime.now().strftime("%H:%M:%S"),
        'v_division': 'Public Works',
        'v_cost_centre': '54321',
        'v_gl': '09876',
        'v_issued_by': 'System Admin'
    }
    print("\nData to be inserted:", sample_data)

    # 3. Run the replacement and conversion
    try:
        replace_placeholders_and_convert(template_file, sample_data, output_docx_file, output_pdf_file)
        print("\nScript finished successfully.")
        print(f"Generated DOCX file: '{output_docx_file}'")
        print(f"Generated PDF file: '{output_pdf_file}' (if conversion was possible)")

    except ValueError as e:
        print(f"\nError: {e}")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")

    print("\n--- Testing Missing Data Scenario ---")
    # Demonstrate the error handling for missing data
    incomplete_data = {
        'v_issued_date': '2023-01-01'
        # This data is missing other required keys like 'v_division', etc.
    }
    print("\nAttempting to process with incomplete data:", incomplete_data)
    try:
        replace_placeholders_and_convert(template_file, incomplete_data, "temp.docx", "temp.pdf")
    except ValueError as e:
        print(f"Successfully caught expected error: {e}")
