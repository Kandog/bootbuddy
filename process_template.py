"""
A script to process a Word document template, replace placeholders, and convert it to PDF.
This script is designed to be comprehensive, searching for and replacing placeholders
in the main body, headers, and footers of the document.
"""
import re
import platform
import datetime
from docx import Document
from docx2pdf import convert

def find_placeholders(template_path):
    """
    Finds all unique placeholders (prefixed with 'v_') in a Word document,
    searching in paragraphs, tables, headers, and footers.
    """
    doc = Document(template_path)
    placeholders = set()
    placeholder_regex = re.compile(r'v_[a-zA-Z0-9_]+')

    def search_text_in_element(element):
        """Searches for placeholders in a given element (paragraph or cell)."""
        if hasattr(element, 'text'):
            found = placeholder_regex.findall(element.text)
            if found:
                for item in found:
                    placeholders.add(item)

    # Search in main body paragraphs
    for para in doc.paragraphs:
        search_text_in_element(para)

    # Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    search_text_in_element(para)

    # Search in headers and footers
    for section in doc.sections:
        for header in (section.header, section.footer):
            if header:
                for para in header.paragraphs:
                    search_text_in_element(para)
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                search_text_in_element(para)

    return list(placeholders)

def replace_placeholders_and_convert(template_path, data, output_docx, output_pdf):
    """
    Replaces placeholders throughout a Word document (including headers/footers)
    and then converts the document to PDF.
    """
    placeholders_in_doc = find_placeholders(template_path)
    missing_keys = [p for p in placeholders_in_doc if p not in data]
    if missing_keys:
        raise ValueError(f"Missing data for the following placeholders: {', '.join(missing_keys)}")

    doc = Document(template_path)

    def replace_in_element(element):
        """Replaces placeholders in a given element's text runs."""
        for para in element.paragraphs:
            for key, value in data.items():
                if key in para.text:
                    for run in para.runs:
                        run.text = run.text.replace(key, str(value))

    # Replace in main body paragraphs and tables
    for para in doc.paragraphs:
        for key, value in data.items():
            if key in para.text:
                for run in para.runs:
                    run.text = run.text.replace(key, str(value))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_element(cell)

    # Replace in headers and footers
    for section in doc.sections:
        for header in (section.header, section.footer):
            if header:
                for para in header.paragraphs:
                     for key, value in data.items():
                        if key in para.text:
                            for run in para.runs:
                                run.text = run.text.replace(key, str(value))
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            replace_in_element(cell)

    doc.save(output_docx)
    print(f"Saved modified document to {output_docx}")

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

    print("--- Analyzing Template: one_year_template.docx ---")
    found_placeholders = find_placeholders(template_file)
    print("Placeholders found by the script:", found_placeholders)

    # Check for placeholders the user mentioned but were not found
    user_expected_placeholders = ['v_name', 'v_po', 'v_eid', 'v_value', 'v_epy_date']
    not_found = [p for p in user_expected_placeholders if p not in found_placeholders]
    if not_found:
        print("\nNOTE: The following placeholders mentioned by the user were NOT found in the document:")
        print(not_found)
        print("The script will proceed using only the placeholders that were found.")

    # In a real application, data would come from a database, API, etc.
    # This sample data includes keys for all possible placeholders.
    full_sample_data = {
        'v_name': 'John Doe',
        'v_po': 'PO123456',
        'v_eid': 'E98765',
        'v_value': '$250.00',
        'v_epy_date': (datetime.date.today() + datetime.timedelta(days=365)).strftime("%Y-%m-%d"),
        'v_issued_date': datetime.date.today().strftime("%Y-%m-%d"),
        'v_issued_time': datetime.datetime.now().strftime("%H:%M:%S"),
        'v_division': 'Fleet Services',
        'v_cost_centre': '112233',
        'v_gl': '445566',
        'v_issued_by': 'Jane Smith'
    }

    # Filter the data to only include placeholders that actually exist in the template
    data_to_process = {key: full_sample_data[key] for key in found_placeholders if key in full_sample_data}

    print("\n--- Processing Template with Found Data ---")
    print("Data being used for replacement:", data_to_process)

    try:
        replace_placeholders_and_convert(template_file, data_to_process, output_docx_file, output_pdf_file)
        print("\nScript finished successfully.")
    except ValueError as e:
        print(f"\nError: {e}")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
