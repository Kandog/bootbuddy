import json
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.utils import ImageReader
import random
import datetime

def create_voucher(data):
    """
    Generates a PDF voucher based on the provided data.

    Args:
        data (dict): A dictionary containing the voucher information.
    """
    c = canvas.Canvas(f"BootBuddy_Voucher_{data['employee_name'].replace(' ', '_')}.pdf", pagesize=letter)
    width, height = letter

    # Set all colors to black
    c.setFillColor(HexColor("#000000"))
    c.setStrokeColor(HexColor("#000000"))

    # Header
    c.drawImage("resource/cok_logo.png", 1 * inch, 10 * inch, width=1*inch, preserveAspectRatio=True)
    c.drawImage("resource/img_head.png", 2 * inch, 10 * inch, width=5*inch, height=0.5*inch, preserveAspectRatio=True)

    # Employee Information
    c.rect(1 * inch, 8.2 * inch, 4.5 * inch, 1 * inch, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.1 * inch, 9 * inch, "Employee Name:")
    c.drawString(1.1 * inch, 8.7 * inch, "PO Number:")
    c.drawString(1.1 * inch, 8.4 * inch, "Employee ID:")

    c.setFont("Helvetica", 10)
    c.drawRightString(4.4 * inch, 9 * inch, data["employee_name"])
    c.drawRightString(4.4 * inch, 8.7 * inch, data["po_number"])
    c.drawRightString(4.4 * inch, 8.4 * inch, data["employee_id"])

    # Voucher Details
    c.rect(1 * inch, 7.2 * inch, 4.5 * inch, 0.8 * inch, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.1 * inch, 7.8 * inch, "Maximum Value:")
    c.drawString(1.1 * inch, 7.4 * inch, "Expiry Date:")

    c.setFont("Helvetica", 10)
    c.drawRightString(4.4 * inch, 7.8 * inch, data["maximum_value"])
    c.drawRightString(4.4 * inch, 7.4 * inch, data["expiry_date"])

    # Boots Image
    c.drawImage("resource/img_boots.png", 5.8 * inch, 7.3 * inch, width=1.7 * inch, height=1.8 * inch, preserveAspectRatio=True)

    # Authorization Text
    styles = getSampleStyleSheet()
    style = ParagraphStyle(name='Justify', parent=styles['Normal'], alignment=TA_JUSTIFY, leading=14)

    auth_text = "This authorization will be redeemed by the Corporation of the City of Kitchener from a supplier with whom the Corporation has a contract for provision of safety footwear, provided that the authorization has been surrendered by an employee of the Corporation at the time of purchasing CSA approved safety footwear, and that the terms of the contract have been met."
    p = Paragraph(auth_text, style)
    p.wrapOn(c, 6.5 * inch, 1 * inch)
    p.drawOn(c, 1 * inch, 6.2 * inch)

    auth_text_2 = "The employee must provide valid photo identification to the vendor upon purchase of the safety footwear."
    p = Paragraph(auth_text_2, style)
    p.wrapOn(c, 6.5 * inch, 1 * inch)
    p.drawOn(c, 1 * inch, 5.6 * inch)

    note_text = "<b>Note:</b> This voucher must NOT be photocopied and may be used to purchase one or more pair of safety footwear only for the designated employee up to the voucher limit. The voucher may only be used once. Any unauthorized purchase will require the employee to reimburse the Corporation of the City of Kitchener for the unauthorized amount."
    p = Paragraph(note_text, style)
    p.wrapOn(c, 6.5 * inch, 1.5 * inch)
    p.drawOn(c, 1 * inch, 4.3 * inch)


    # Internal Use Only Section
    c.rect(1 * inch, 1.8 * inch, 6.5 * inch, 2.2 * inch, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.1 * inch, 3.8 * inch, "For Internal Use Only:")
    c.line(1 * inch, 3.7 * inch, 7.5 * inch, 3.7 * inch)
    c.line(4.25 * inch, 1.8 * inch, 4.25 * inch, 3.7 * inch)

    # Labels
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.1 * inch, 3.5 * inch, "Date Voucher Issued:")
    c.drawString(1.1 * inch, 3.2 * inch, "Time Issued:")
    c.drawString(1.1 * inch, 2.9 * inch, "Division:")
    c.drawString(1.1 * inch, 2.6 * inch, "Cost Centre:")
    c.drawString(1.1 * inch, 2.3 * inch, "GL:")
    c.drawString(1.1 * inch, 2.0 * inch, "Issued By:")

    # Values
    c.setFont("Helvetica", 10)
    c.drawString(4.35 * inch, 3.5 * inch, data["date_voucher_issued"])
    c.drawString(4.35 * inch, 3.2 * inch, data["time_issued"])
    c.drawString(4.35 * inch, 2.9 * inch, data["division"])
    c.drawString(4.35 * inch, 2.6 * inch, data["cost_centre"])
    c.drawString(4.35 * inch, 2.3 * inch, data["gl"])
    c.drawString(4.35 * inch, 2.0 * inch, data["issued_by"])

    # Horizontal lines
    c.line(1 * inch, 3.4 * inch, 7.5 * inch, 3.4 * inch)
    c.line(1 * inch, 3.1 * inch, 7.5 * inch, 3.1 * inch)
    c.line(1 * inch, 2.8 * inch, 7.5 * inch, 2.8 * inch)
    c.line(1 * inch, 2.5 * inch, 7.5 * inch, 2.5 * inch)
    c.line(1 * inch, 2.2 * inch, 7.5 * inch, 2.2 * inch)

    # Barcodes and Vendor Logos
    c.drawImage("resource/img_mister_safety_shoes.png", 1.1 * inch, 1.2 * inch, width=1.5*inch, preserveAspectRatio=True)
    c.drawImage("resource/img_cok100.png", 1.1 * inch, 0.9 * inch, width=1*inch, preserveAspectRatio=True)

    c.drawImage("resource/img_work_authority.png", 3.5 * inch, 1.2 * inch, width=1.5*inch, preserveAspectRatio=True)
    c.drawImage("resource/img_112934.png", 3.5 * inch, 0.9 * inch, width=1*inch, preserveAspectRatio=True)

    c.drawImage("resource/img_marks.png", 5.9 * inch, 1.2 * inch, width=1.5*inch, preserveAspectRatio=True)
    c.drawImage("resource/img_acct.png", 5.9 * inch, 0.9 * inch, width=1*inch, preserveAspectRatio=True)


    # Footer
    c.setFont("Helvetica", 9)
    c.drawCentredString(4.25 * inch, 0.5 * inch, "Vendors: If you have any questions, call Supply Services, Mon-Fri 8:30am-5pm, (519) 741-2200 ext. 7217")

    c.save()

if __name__ == "__main__":
    # Example Usage
    sample_data = {
        "employee_name": "John Doe",
        "po_number": str(random.randint(1000, 9999)),
        "employee_id": str(random.randint(10000, 99999)),
        "maximum_value": f"${random.uniform(100, 500):.2f}",
        "expiry_date": (datetime.date.today() + datetime.timedelta(days=30)).strftime("%B %d, %Y"),
        "for_internal_use_only": "O:\\Secured\\Safety Footwear - CUPE 68\\",
        "date_voucher_issued": datetime.date.today().strftime("%B %d, %Y"),
        "time_issued": datetime.datetime.now().strftime("%I:%M:%S %p"),
        "division": "CSD-Parks & Cemeteries",
        "cost_centre": str(random.randint(100000, 999999)),
        "gl": str(random.randint(100000, 999999)),
        "issued_by": "Jane Smith"
    }

    create_voucher(sample_data)
    print(f"Generated voucher for {sample_data['employee_name']}")