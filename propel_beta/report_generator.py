import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import yellow
from reportlab.lib.colors import Color
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Frame, Paragraph, Table, TableStyle
from .excel_parser import getprojectdetails, read_excel_tasks, convert_tasks_for_gantt, getprojectsections, get_payment_milestones, get_delivery_data
import os
import reportlab.lib.enums as enums
from .cpm import cpmcalc
from .gantt import create_gantt
from PIL import Image



def create_report(project_path, po_data):
    # Create a canvas object
    if os.path.exists(project_path) == False:
        return "No such File/Directory"
    #project_path = 'C:\\Users\\arjun\\OneDrive\\Documents\\UCSD\\SPRING2024\\Propel\\Sample_Project'
    project_details = getprojectdetails(project_path)
    progress_summary_s = getprojectsections(project_path, sheetname='Progress Summary')
    areasofconcern = getprojectsections(project_path, sheetname='Areas of Concern')
    head, tail = os.path.split(project_path)
    output_path = os.path.join(head, "monthly_report.pdf")
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    mustard_color = Color(255 / 255, 204 / 255, 102 / 255)


    # Define header and footer
    def draw_header(c, width, height, left_text, right_text):
        header_font = "Helvetica-Bold"
        header_font_size = 12
        line_y_position = height - 0.6 * inch

        # Draw left-aligned text
        c.setFont(header_font, header_font_size)
        c.drawString(0.5 * inch, height - 0.5 * inch, left_text)

        # Draw right-aligned text
        c.setFont(header_font, header_font_size)
        right_text_width = c.stringWidth(right_text, header_font, header_font_size)
        c.drawString(width - right_text_width - 0.5 * inch, height - 0.5 * inch, right_text)

        # Draw yellow line
        c.setStrokeColor(mustard_color)
        c.setLineWidth(3)
        c.line(0.5 * inch, line_y_position, width - 0.5 * inch, line_y_position)


        # Function to draw footer
    def draw_footer(c, width, height, page_number, total_pages, date_text):
        # Footer text configuration
        footer_font = "Helvetica"
        footer_font_size = 10
        line_y_position = 0.7 * inch

        # Draw yellow line above the footer
        c.setStrokeColor(mustard_color)
        c.setLineWidth(3)
        c.line(0.5 * inch, line_y_position, width - 0.5 * inch, line_y_position)

        # Draw "Page x of y" centered
        footer_text = f"Page {page_number}"
        c.setFont(footer_font, footer_font_size)
        footer_text_width = c.stringWidth(footer_text, footer_font, footer_font_size)
        c.drawCentredString(width / 2.0, 0.5 * inch, footer_text)

        # Draw date right-aligned
        date_text_width = c.stringWidth(date_text, footer_font, footer_font_size)
        c.drawString(width - date_text_width - 0.5 * inch, 0.5 * inch, date_text)

    
    projectname = ''
    projectfullname = ''
    logo1 = ''
    logo2 = ''
    contracttitle = ''
    attention = ''
    cc = ''
    if 'ProjectName' in project_details:
        projectname = project_details['ProjectName']
    if 'Project Full Name' in project_details:
        projectfullname = project_details['Project Full Name']
    if 'Contract Title' in project_details:
        contracttitle = project_details['Contract Title']
    if('Logo1' in project_details):
        logo1 = project_details['Logo1']
    if ('Logo2' in project_details):
        logo2 = project_details['Logo2']
    if ('Attention' in project_details):
        attention = project_details['Attention']
    if ('CC' in project_details):
        cc = project_details['CC']


    total_pages = 4
    date_text = datetime.datetime.now().strftime("%Y-%m-%d")
    curr_page = 1
    # First page content
    draw_header(c, width, height, "Solar Turbines International", projectname + "-Monthly Report")
    draw_footer(c, width, height, curr_page, total_pages, date_text)
    #c.setFont("Helvetica", 12)
    #c.drawString(2 * inch, 7.5 * inch, "Content on the first page.")
    styles = getSampleStyleSheet()
    body_style = styles["Title"]
    frame = Frame(1.2 * inch, 4 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "MONTHLY PROGRESS Report"
    paragraph = Paragraph(text, body_style)
    frame.addFromList([paragraph], c)

    body_style = styles["Title"]
    frame = Frame(1.2 * inch, 3.7 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = projectfullname
    paragraph = Paragraph(text, body_style)
    frame.addFromList([paragraph], c)

    custom_style = ParagraphStyle(
        'CustomStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=16,
        leading=16,
        #textColor=colors.blue,
        alignment= enums.TA_CENTER,
        spaceBefore=12,
        spaceAfter=12,
    )
    frame = Frame(1.2 * inch, 3.4 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "Contract Title: " + contracttitle
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    custom_style = ParagraphStyle(
        'CustomStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=14,
        leading=16,
        # textColor=colors.blue,
        alignment=enums.TA_LEFT,
        spaceBefore=12,
        spaceAfter=12,
    )
    frame = Frame(1 * inch, 1.5 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "Attention: " + attention
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    frame = Frame(1 * inch, 1.2 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "CC: " + cc
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    # Insert Logo 1
    image_path = logo1
    image_x = 1 * inch  # X coordinate
    image_y = 8 * inch  # Y coordinate
    image_width = 3 * inch
    image_height = 2 * inch
    c.drawImage(image_path, image_x, image_y, image_width, image_height)

    # Insert Logo 2
    image_path = logo2
    image_x = 4.5 * inch  # X coordinate
    image_y = 8 * inch  # Y coordinate
    image_width = 3 * inch
    image_height = 2 * inch
    c.drawImage(image_path, image_x, image_y, image_width, image_height)

    # Finalize the first page
    c.showPage()
    curr_page += 1
    draw_header(c, width, height, "Solar Turbines International", projectname + "-Monthly Report")
    draw_footer(c, width, height, curr_page, total_pages, date_text)
    #c.drawString(1*inch, 5*inch, "This is content on a new page.")

    po_date = ''
    contract_duration = ''
    readiness_to_ship_forecast = ''
    contractual_delivery_date = ''
    forecast_delivery_date = ''
    delivery_term = ''
    vendor_location = ''
    scope_of_work = ''
    if 'PO/LOA Date' in project_details:
        po_date = project_details['PO/LOA Date']
    if 'Contract Duration' in project_details:
        contract_duration = project_details['Contract Duration']
    if 'Readiness to Ship Forecast' in project_details:
        readiness_to_ship_forecast = project_details['Readiness to Ship Forecast']
    if 'Contractual Delivery Date' in project_details:
        contractual_delivery_date = project_details['Contractual Delivery Date']
    if 'Forecast Delivery Date' in project_details:
        forecast_delivery_date = project_details['Forecast Delivery Date']
    if 'Delivery Term' in project_details:
        delivery_term = project_details['Delivery Term']
    if 'Vendor/Manufacturing Location' in project_details:
        vendor_location = project_details['Vendor/Manufacturing Location']
    if 'Scope of Work' in project_details:
        scope_of_work = project_details['Scope of Work']

    custom_style = ParagraphStyle(
        'CustomStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        # textColor=colors.blue,
        alignment=enums.TA_LEFT,
        #spaceBefore=12,
        #spaceAfter=12,
    )
    custom_style2 = ParagraphStyle(
        'CustomStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=16,
        # textColor=colors.blue,
        alignment=enums.TA_LEFT,
        # spaceBefore=12,
        # spaceAfter=12,
    )
    
    #SECTION 1.0
    frame = Frame(0.7 * inch, 7.0 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.0 PROJECT SUMMARY"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    frame = Frame(1.3 * inch, 6.7 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.1 PO/LOA DATE:\t" + po_date
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    frame = Frame(1.3 * inch, 6.4 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.2 CONTRACT DURATION:\t" + contract_duration
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    frame = Frame(1.3 * inch, 6.1 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.3 READINESS TO SHIP FORECAST:\t" + readiness_to_ship_forecast
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    frame = Frame(1.3 * inch, 5.8 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.4 CONTRACTUAL DELIVERY DATE:\t" + contractual_delivery_date
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    frame = Frame(1.3 * inch, 5.5 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.5 FORECAST DELIVERY DATE:\t" + forecast_delivery_date
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    frame = Frame(1.3 * inch, 5.2 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.6 DELIVERY TERM:\t" + delivery_term
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    frame = Frame(1.3 * inch, 4.9 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.7 VENDOR/MANUFACTURING LOCATION:\t" + vendor_location
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    frame = Frame(1.3 * inch, 4.6 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "1.8 SCOPE OF WORK:\t" + scope_of_work
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    #SECTION 2.0
    frame = Frame(0.7 * inch, 4.1 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "2.0 HEALTH, SAFETY AND ENVIRONMENT"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    frame = Frame(1.3 * inch, 3.9 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "2.1 No COVID 19 concerns"
    paragraph = Paragraph(text, custom_style2)
    frame.addFromList([paragraph], c)

    #SECTION 3.0
    frame = Frame(0.7 * inch, 3.4 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "3.0 PROGRESS SUMMARY - S-CURVE"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    curr_y = 3.1
    for i, point in enumerate(progress_summary_s):
        paragraph = Paragraph("3." + str(i+1) + " " + point, custom_style2)
        # Calculate the height of the paragraph
        frame_width = width - 2.4 * inch
        paragraph.wrapOn(c, frame_width, height)
        paragraph_height = paragraph.height / inch
        # Create a frame and add the paragraph
        #frame = Frame(1.3 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
        frame = Frame(1.3 * inch, (curr_y * inch) - paragraph_height, frame_width, height - 8 * inch, showBoundary=0)
        frame.addFromList([paragraph], c)
        # Update current y position
        curr_y -=  paragraph_height

    curr_y -= 0.3

    #SECTION 4.0
    frame = Frame(0.7 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "4.0 PROGRESS SUMMARY - PROCUREMENT"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    


    #END OF PAGE 2
    c.showPage()
    curr_page += 1
    draw_header(c, width, height, "Solar Turbines International", projectname + "-Monthly Report")
    draw_footer(c, width, height, curr_page, total_pages, date_text)
    #START OF PAGE 3

    curr_y = 7.0
    #SECTION 5.0
    frame = Frame(0.7 * inch, 7.0 * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "5.0 SCHEDULED DELIVERY"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    curr_y = 6.3
    sched_del_data = get_delivery_data(project_path=project_path)
    # Create a style for paragraphs
    styles = getSampleStyleSheet()
    styleN = styles['BodyText']

    # Convert cell data to Paragraphs for text wrapping
    for i, row in enumerate(sched_del_data):
        for j, cell in enumerate(row):
            sched_del_data[i][j] = Paragraph(cell, styleN)
    # Create the table
    table1 = Table(sched_del_data, colWidths=[0.5 * inch, 2.5 * inch, 0.5 * inch, 1 * inch, 1 * inch, 1*inch, 1*inch])

    # Apply some basic styling
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    # Calculate the height of the table
    table1.wrapOn(c, width, height)
    table_height = table1._height/inch

    # Define starting position for the table
    x = 0.7 * inch
    y = ((curr_y - table_height + 3.0) * inch)  # Adjust the y-coordinate as needed

    # Draw the table on the canvas
    table1.drawOn(c, x, y)


    curr_y = curr_y - table_height - 1
    #SECTION 6.0
    frame = Frame(0.7 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "6.0 PAYMENT MILESTONES"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    curr_y -= 0.2
    '''
    points = po_data[0].split('##')
    for point in points:
        paragraph = Paragraph(point, custom_style2)
        # Calculate the height of the paragraph
        frame_width = width - 2.4 * inch
        paragraph.wrapOn(c, frame_width, height)
        paragraph_height = paragraph.height / inch
        
        # Check if the current y position can accommodate the paragraph
        #if curr_y - paragraph_height < 0:
        #    c.showPage()  # Create a new page if there is not enough space
        #    curr_y = height - 1 * inch  # Reset y position
        
        # Create a frame and add the paragraph
        #frame = Frame(1.3 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
        frame = Frame(1.0 * inch, (curr_y * inch) - paragraph_height, frame_width, height - 8 * inch, showBoundary=0)
        frame.addFromList([paragraph], c)
        
        # Update current y position
        curr_y -=  paragraph_height
    '''

    paymentdata = get_payment_milestones(project_path=project_path)
    
    # Convert cell data to Paragraphs for text wrapping
    for i, row in enumerate(paymentdata):
        for j, cell in enumerate(row):
            paymentdata[i][j] = Paragraph(cell, styleN)
    # Create the table
    table2 = Table(paymentdata, colWidths=[0.5 * inch, 3 * inch, 0.5 * inch, 1 * inch, 1.5 * inch])

    # Apply some basic styling
    table2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    # Calculate the height of the table
    table2.wrapOn(c, width, height)
    table_height = table2._height/inch

    # Define starting position for the table
    x = 0.7 * inch
    y = ((curr_y - table_height + 2.5) * inch)  # Adjust the y-coordinate as needed

    # Draw the table on the canvas
    table2.drawOn(c, x, y)

    #END OF PAGE 3
    c.showPage()
    curr_page += 1
    draw_header(c, width, height, "Solar Turbines International", projectname + "-Monthly Report")
    draw_footer(c, width, height, curr_page, total_pages, date_text)
    #START OF PAGE 4

    curr_y = 7.0
    #SECTION 7.0
    frame = Frame(0.7 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "7.0 AREAS OF CONCERN"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    curr_y = 6.7
    for i, point in enumerate(areasofconcern):
        paragraph = Paragraph("7." + str(i+1) + " " + point, custom_style2)
        # Calculate the height of the paragraph
        frame_width = width - 2.4 * inch
        paragraph.wrapOn(c, frame_width, height)
        paragraph_height = paragraph.height / inch
        # Create a frame and add the paragraph
        #frame = Frame(1.3 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
        frame = Frame(1.3 * inch, (curr_y * inch) - paragraph_height, frame_width, height - 8 * inch, showBoundary=0)
        frame.addFromList([paragraph], c)
        # Update current y position
        curr_y -=  paragraph_height

    curr_y -= 0.5
    #SECTION 8.0
    frame = Frame(0.7 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "8.0 QUALITY / INSPECTION AND TEST"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)
    curr_y -= 0.5

    #SECTION 9.0
    frame = Frame(0.7 * inch, curr_y * inch, width - 2.4 * inch, height - 8 * inch, showBoundary=0)
    text = "9.0 NEXT MONTH ACTIONS"
    paragraph = Paragraph(text, custom_style)
    frame.addFromList([paragraph], c)

    #END OF PAGE 4
    c.showPage()
    curr_page += 1
    draw_header(c, width, height, "Solar Turbines International", projectname + "-Monthly Report")
    draw_footer(c, width, height, curr_page, total_pages, date_text)
    #START OF PAGE 5

    project_tasks = read_excel_tasks(project_path)
    project_tasks = convert_tasks_for_gantt(project_tasks)
    gantt_path = os.path.join(head, "temp_gantt.png")
    create_gantt(project_tasks, gantt_path)
    image_path = gantt_path
    with Image.open(image_path) as img:
        original_width, original_height = img.size
    #Calculate the scaling factor to fit the image within the page dimensions
    width_ratio = width / original_width
    height_ratio = height / original_height
    scaling_factor = min(width_ratio, height_ratio)
    
    # Calculate the new width and height while maintaining aspect ratio
    new_width = original_width * scaling_factor
    new_height = original_height * scaling_factor
    
    # Calculate coordinates to center the image on the page
    image_x = (width - new_width) / 2
    image_y = (height - new_height) / 2
    # Draw the image with the new dimensions, centered on the page
    c.drawImage(image_path, image_x, image_y, width=new_width, height=new_height, preserveAspectRatio=True, anchor='sw')

    # Save the PDF
    c.save()
    return "Successfully generated report: " + output_path

