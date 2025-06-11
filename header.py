from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm

def create_order_header_pdf(filename):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Margins and positions
    left_margin = 20 * mm
    top_margin = height - 20 * mm

    # Draw Title
    c.setFont("Helvetica-BoldOblique", 24)
    c.setFillColor(colors.black)
    c.drawString(left_margin, top_margin, "Balaji Surgicals")

    # Subheading
    c.setFont("Helvetica-Oblique", 16)
    c.setFillColor(colors.black)
    c.drawString(left_margin, top_margin - 20, "Nawabi Road, Haldwani–263 139")

    # Orange Line
    c.setStrokeColor(colors.darkorange)
    c.setLineWidth(2)
    c.line(left_margin, top_margin - 30, width - left_margin, top_margin - 30)

    # Box lines
    y = top_margin - 40
    row_height = 18

    # GST, DL
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, "GSTIN No. 05ADGPA2715B1ZH")
    y -= row_height
    c.drawString(left_margin, y, "D.L.No. 20B:121663, 21B:121664")

    # Contact Info (Right)
    c.setFont("Helvetica", 10)
    c.drawRightString(width - left_margin, top_margin - 40, "Ph.:  97193-04441")
    c.drawRightString(width - left_margin, top_margin - 58, "e-mail : balajisumit@yahoo.co.in")

    # Order Form (Right aligned, bold)
    c.setFont("Helvetica-Bold", 16)
    c.drawRightString(width - left_margin, top_margin, "Order Form")

    # Recipient Info
    y -= 2 * row_height
    c.setFont("Helvetica-BoldOblique", 12)
    c.drawString(left_margin, y, "To,")
    y -= row_height
    c.drawString(left_margin, y, "Company Name")
    y -= row_height
    c.drawString(left_margin, y, "Company Location")
    y -= row_height
    c.setFont("Helvetica", 11)
    c.drawString(left_margin, y, "Email ids")

    # Order Date and No. (Right side)
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(width - left_margin, y + 2 * row_height, "Date : 18/09/24")
    c.drawRightString(width - left_margin, y + row_height, "Order No. : 1")

    # Message
    y -= 2 * row_height
    c.setFont("Helvetica-BoldOblique", 12)
    c.drawString(left_margin, y, "Dear Sir,")
    y -= row_height
    c.drawString(left_margin, y, "Kindly supply the following items:")

    c.save()

create_order_header_pdf("order_header.pdf")