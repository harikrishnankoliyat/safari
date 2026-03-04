import io
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

def generate_word_quotation(q):
    doc = Document()
    # --- SET GLOBAL FONT TO CALIBRI ---
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # --- 1. A4 PAGE SETUP & NARROW MARGINS ---
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.5)
    
    content_width = Cm(18.46)
    
    # --- 2. THE FIXED HEADER ---
    section.different_first_page_header_footer = True
    header = section.first_page_header
    htable = header.add_table(1, 2, width=content_width)
    
    htable.allow_autofit = False  
    htable.table_layout = 'fixed' 
    htable.columns[0].width = Cm(14.5) 
    htable.columns[1].width = Cm(3.96)
    htable.rows[0].cells[0].width = Cm(14.5)
    htable.rows[0].cells[1].width = Cm(3.96)

    for cell in htable.rows[0].cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for node in ['top', 'left', 'bottom', 'right']:
            node_el = OxmlElement(f'w:{node}')
            node_el.set(qn('w:w'), '0')
            node_el.set(qn('w:type'), 'dxa')
            tcMar.append(node_el)
        tcPr.append(tcMar)
    
    left_cell = htable.rows[0].cells[0]
    left_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    header_text = left_cell.paragraphs[0]
    run_name = header_text.add_run("JIGSAW AFRICA WILDLIFE SAFARIS LTD (JAWS AFRICA)\n")
    run_name.font.bold = True
    run_name.font.size = Pt(17)
    run_reg = header_text.add_run("REGISTRATION PIN: P052221766X\n")
    run_reg.font.bold = True
    run_reg.font.size = Pt(11)
    
    details = (
    "Mercantile House, 2nd floor, Room 230, Koinange street, Nairobi, Kenya\n"
    "Emails: info@jawsafrica.com | Web: www.jawsafrica.com\n"
    "Mobile: +254 719899245 (KE) | +965 94067244 (KW) | +91 9496656977 (IN)"
)
    run_details = header_text.add_run(details)
    run_details.font.size = Pt(11)

    right_cell = htable.rows[0].cells[1]
    right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    logo_para = right_cell.paragraphs[0]
    logo_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
    if os.path.exists("logo.png"):
        logo_para.add_run().add_picture("logo.png", width=Cm(3.5))

    # --- 3. DOUBLE SOLID LINE DIVIDER ---
    line_para = header.add_paragraph()
    p_obj = line_para._p
    pPr = p_obj.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'double') 
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)

    # --- 4. BODY CONTENT (UNDERLINED TITLE WITH GAP) ---
    title = doc.add_paragraph()
    # Adding space_after creates the gap between the title and summary lines
    title.paragraph_format.space_after = Pt(10) 
    
    run = title.add_run(f"Quotation for {q['client']}| {q['country']}")
    run.bold = True
    run.underline = True # Adds the underline to the title
    run.font.size = Pt(14)
    run.font.name = 'Calibri'

    summary_lines = [
        f"TOUR CODE:   {q['code']}",
        f"{q['country'].upper()} SAFARI PRIVATE PACKAGE :  {q['pkg']}", # Dynamic country variable
        f"DATE:   FROM {q['start']} TO {q['end']}"
    ]

    for line in summary_lines:
        p = doc.add_paragraph(line)
        p.paragraph_format.space_after = Pt(2) 
        p.paragraph_format.line_spacing = 1.0

    # --- HELPER FOR STYLED HEADINGS ---
    def add_styled_heading(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Cm(0) 
        
        # Gold background shading
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), 'FFC000') 
        p._p.get_or_add_pPr().append(shading_elm)
        
        run = p.add_run(f" {text}") # Leading space for padding
        run.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    def set_cell_grey(cell):
        """Sets a light grey background for table headers"""
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), 'D9D9D9') 
        cell._tc.get_or_add_tcPr().append(shading_elm)

    # --- 5. ITINERARY TABLE ---
    # --- 5. ITINERARY TABLE ---
    add_styled_heading('ITINERARY PLAN')
    iti_table = doc.add_table(rows=1, cols=6)
    iti_table.style = 'Table Grid'
    iti_table.alignment = WD_ALIGN_PARAGRAPH.LEFT
    iti_table.width = content_width

    # SHIFT TABLE RIGHT: Accessing XML to set left indentation
    tbl_pr = iti_table._element.xpath('w:tblPr')[0]
    tbl_ind = OxmlElement('w:tblInd')
    tbl_ind.set(qn('w:w'), '100') # Value in twentieths of a point (approx 0.44cm)
    tbl_ind.set(qn('w:type'), 'dxa')
    tbl_pr.append(tbl_ind)

    hdrs = ["Day", "From", "To", "Activities", "Accommodation", "Meal Plan"]
    for i, h in enumerate(hdrs):
        cell = iti_table.rows[0].cells[i]
        cell.text = h
        set_cell_grey(cell) # Apply the grey background
        # Set Bold and Calibri for headers
        p = cell.paragraphs[0]
        if p.runs:
            run = p.runs[0]
            run.bold = True
            run.font.name = 'Calibri'

    # Populating ITINERARY data with explicit mapping
    for row_data in q['iti']:
        row_cells = iti_table.add_row().cells
        row_cells[0].text = str(row_data.get("Day", ""))
        row_cells[1].text = str(row_data.get("From", ""))
        row_cells[2].text = str(row_data.get("To", ""))
        row_cells[3].text = str(row_data.get("Activities", ""))
        row_cells[4].text = str(row_data.get("Accommodation", ""))
        row_cells[5].text = str(row_data.get("Meal Plan", ""))
        
        # Ensure Calibri for all table content
        for cell in row_cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = 'Calibri'

    # --- 6. TARIFF TABLE ---
    # --- 6. TARIFF TABLE ---
    add_styled_heading('TARIFF IN USD')
    t_table = doc.add_table(rows=2, cols=3)
    t_table.style = 'Table Grid'
    t_table.alignment = WD_ALIGN_PARAGRAPH.LEFT
    t_table.width = content_width

    # SHIFT TABLE RIGHT: Matching the shift for the itinerary table
    t_tbl_pr = t_table._element.xpath('w:tblPr')[0]
    t_tbl_ind = OxmlElement('w:tblInd')
    t_tbl_ind.set(qn('w:w'), '100') 
    t_tbl_ind.set(qn('w:type'), 'dxa')
    t_tbl_pr.append(t_tbl_ind)

    # Header Row for Tariff with Grey BG
    t_hdrs = ["Item", "Total Cost", "Cost per Person"]
    for i, h in enumerate(t_hdrs):
        cell = t_table.rows[0].cells[i]
        cell.text = h
        set_cell_grey(cell) # Apply grey background
        p = cell.paragraphs[0]
        if p.runs:
            p.runs[0].bold = True
            p.runs[0].font.name = 'Calibri'

    tr = t_table.rows[1].cells
    tr[0].text = f"{q['adults']} PAX Safari"
    tr[1].text = f"{q['total']:,.0f}"
    tr[2].text = f"{q['pp']:,.0f}"
    
    # Ensure Calibri for Tariff data
    for cell in tr:
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].font.name = 'Calibri'
    # --- TARIFF SUMMARY SENTENCE ---
    summary_text = (
        f"This price is for {q['adults']} adults traveling in {q['vehicles']} private "
        f"safari vehicle(s) with accommodation in {q['accommodation_summary']} "
        f"as per the itinerary above."
    )
    
    p_summary = doc.add_paragraph(summary_text)
    p_summary.paragraph_format.space_after = Pt(8)
    p_summary.paragraph_format.left_indent = Cm(0.12)
    
    for run in p_summary.runs:
        run.font.name = 'Calibri'
        run.italic = True 
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(0, 0, 0)

    # --- 7. DETAILED ITINERARY SECTION (JUSTIFIED) ---
    if q.get('detailed_iti'):
        add_styled_heading('DETAILED ITINERARY')
        
        for day_info in q['detailed_iti']:
            # --- Day Title ---
            p_day = doc.add_paragraph()
            p_day.paragraph_format.space_before = Pt(10)
            p_day.paragraph_format.space_after = Pt(2)
            p_day.paragraph_format.left_indent = Cm(0) 
            
            run_day = p_day.add_run(f"{day_info['day']} :-")
            run_day.bold = True
            run_day.underline = True
            run_day.font.name = 'Calibri'
            run_day.font.size = Pt(11)
            
            # --- Day Details (Justified Alignment) ---
            p_details = doc.add_paragraph(day_info['details'])
            p_details.paragraph_format.space_after = Pt(12)
            
            # SET ALIGNMENT TO JUSTIFY
            p_details.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            # Match table border alignment (0.12cm)
            p_details.paragraph_format.left_indent = Cm(0.12)
            p_details.paragraph_format.right_indent = Cm(0)
            
            # Standard line spacing for readability
            p_details.paragraph_format.line_spacing = 1.15
            
            for run in p_details.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(10.5)
                run.font.color.rgb = RGBColor(0, 0, 0)
   

    # --- 7. UPDATED INCLUSIONS ---
    add_styled_heading('INCLUSIONS')
    inclusions = [
        "Mid-Range accommodation with meal plans as stated in the itinerary",
        "Transport in a private vehicle & use of customized 4WD Land cruisers with a Driver/Guide",
        "Game drive time of your choice, we can extend our support for full day game drive at no extra cost",
        "Drinking water on safari and transfers",
        "All National Parks entrance fees and Taxes"
    ]
    for item in inclusions:
        doc.add_paragraph(item, style='List Bullet')
    
    # --- 8. UPDATED EXCLUSIONS ---
    add_styled_heading('EXCLUSIONS')
    exclusions = [
        "International/Domestic Flights/Visa and travel insurance",
        "Yellow Fever Vaccination or travel related test (PCR etc if applicable)",
        "Alcoholic drinks or soft drinks",
        "Items of personal nature",
        "Extra accommodations in Nairobi",
        "Extra meals in any location",
        "Tips 10 USD per day per person for the driver (Optional but highly recommended in camps as well)",
        
    ]
    for item in exclusions:
        doc.add_paragraph(item, style='List Bullet')


    doc.add_paragraph("\n\nThank you for Choosing Jaws Africa").alignment = WD_ALIGN_PARAGRAPH.CENTER

 # --- 8. FOOTER CODE ---

    def setup_styled_footer(footer_obj):
        footer_obj.is_linked_to_previous = False
        f_para = footer_obj.paragraphs[0]
        f_para.clear()
        
        # 1. Create a 1x2 table inside the footer to force margins
        # Using content_width (18.46cm) to match the body content exactly
        f_table = footer_obj.add_table(1, 2, width=content_width)
        f_table.allow_autofit = False
        
        # 2. Set column widths (Website gets 70%, Page Number gets 30%)
        f_table.columns[0].width = Cm(13.0)
        f_table.columns[1].width = Cm(5.46)
        
        # 3. Left Cell: Website URL
        left_cell = f_table.rows[0].cells[0]
        l_para = left_cell.paragraphs[0]
        l_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        l_run = l_para.add_run("www.jawsafrica.com")
        l_run.font.name = 'Calibri'
        l_run.font.size = Pt(10)
        
        # 4. Right Cell: Page Number
        right_cell = f_table.rows[0].cells[1]
        r_para = right_cell.paragraphs[0]
        r_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT # This forces it to the far right margin
        r_para.add_run("Page ")
        
        # --- Dynamic Page Number Field ---
        f_p = r_para._p
        f_run_field = OxmlElement('w:r')
        f_fldChar_start = OxmlElement('w:fldChar')
        f_fldChar_start.set(qn('w:fldCharType'), 'begin')
        f_run_field.append(f_fldChar_start)
        f_p.append(f_run_field)
        
        f_run_instr = OxmlElement('w:r')
        f_instrText = OxmlElement('w:instrText')
        f_instrText.text = "PAGE"
        f_run_instr.append(f_instrText)
        f_p.append(f_run_instr)
        
        f_run_end = OxmlElement('w:r')
        f_fldChar_end = OxmlElement('w:fldChar')
        f_fldChar_end.set(qn('w:fldCharType'), 'end')
        f_run_end.append(f_fldChar_end)
        f_p.append(f_run_end)
        
        # Set font for page number
        for run in r_para.runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(10)

    # Apply to both page versions
    setup_styled_footer(section.first_page_footer)
    setup_styled_footer(section.footer)


    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()