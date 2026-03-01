import io
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def generate_word_quotation(q):
    doc = Document()
    
    # --- 1. A4 PAGE SETUP & NARROW MARGINS ---
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    
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
    run_name.font.size = Pt(15)
    
    details = (
        "REGISTRATION PIN: P052221766X\n"
        "Mercantile House, 2nd floor, Room 230, Koinange street, Nairobi, Kenya\n"
        "PO.Box – 5270-00100\n"
        "+254 719899245 | +965 94067244 | info@jawsafrica.com | www.jawsafrica.com"
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

    # --- 4. BODY CONTENT (TIGHT SPACING) ---
    title = doc.add_paragraph()
    title.paragraph_format.space_after = Pt(0) 
    run = title.add_run(f"Quotation for {q['client']}| {q['country']}")
    run.bold = True
    run.font.size = Pt(14)

    summary_lines = [
        f"TOUR CODE:   {q['code']}",
        f"KENYA SAFARI PRIVATE PACKAGE :  {q['pkg']}",
        f"DATE:   FROM {q['start']} TO {q['end']}"
    ]

    for line in summary_lines:
        p = doc.add_paragraph(line)
        p.paragraph_format.space_after = Pt(2) 
        p.paragraph_format.line_spacing = 1.0

    # --- 5. ITINERARY TABLE ---
    doc.add_heading('ITINERARY PLAN', level=2)
    iti_table = doc.add_table(rows=1, cols=6)
    iti_table.style = 'Table Grid'
    iti_table.width = content_width
    
    hdrs = ["Day", "From", "To", "Activities", "Accommodation", "Meal Plan"]
    for i, h in enumerate(hdrs):
        iti_table.rows[0].cells[i].text = h
        iti_table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
    
    for row in q['iti']:
        cells = iti_table.add_row().cells
        for i, key in enumerate(hdrs):
            cells[i].text = str(row.get(key, ""))
    
    # --- 6. TARIFF TABLE ---
    doc.add_heading('TARIFF IN USD', level=2)
    t_table = doc.add_table(rows=2, cols=3)
    t_table.style = 'Table Grid'
    t_table.rows[0].cells[0].text, t_table.rows[0].cells[1].text, t_table.rows[0].cells[2].text = "Item", "Total Cost", "Cost per Person"
    
    tr = t_table.rows[1].cells
    tr[0].text = f"{q['adults']} PAX Safari"
    tr[1].text = f"{q['total']:,.0f}"
    tr[2].text = f"{q['pp']:,.0f}"

    # --- 7. UPDATED INCLUSIONS ---
    doc.add_heading('INCLUSIONS', level=2)
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
    doc.add_heading('EXCLUSIONS', level=2)
    exclusions = [
        "International/Domestic Flights/Visa and travel insurance",
        "Yellow Fever Vaccination or travel related test (PCR etc if applicable)",
        "Alcoholic drinks or soft drinks",
        "Items of personal nature",
        "Extra accommodations in Nairobi",
        "Extra meals in any location",
        "Tips 10 USD per day per person for the driver (Optional but highly recommended in camps as well)",
        "Hot Air Balloon (450 USD/Adult – Early reservation required)"
    ]
    for item in exclusions:
        doc.add_paragraph(item, style='List Bullet')

    doc.add_heading('TERMS & PAYMENT POLICY', level=2)
    doc.add_paragraph("Deposit: 30% required at booking. Final Payment: 30 days before start.", style='List Bullet')

    doc.add_paragraph("\nThank you for Choosing Jaws Africa").alignment = WD_ALIGN_PARAGRAPH.CENTER

    footer = section.footer
    f_para = footer.paragraphs[0]
    f_para.text = "www.jawsafrica.com"
    f_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()