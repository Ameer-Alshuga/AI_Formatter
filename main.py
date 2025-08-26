# main.py - The Intelligent Cleaning & Conversion Engine

import re
from bs4 import BeautifulSoup, Comment
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_DIRECTION
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

### UPDATED & SMARTER CLEANER FUNCTION ###
def clean_html(soup):
    # Remove script and style elements
    for script_or_style in soup(['script', 'style']):
        script_or_style.decompose()

    # Remove elements that are clearly UI buttons or junk
    for element in soup.find_all(text=re.compile(r'IGNORE_WHEN_COPYING|downloadcontent_copy|expand_less')):
        element.find_parent().decompose()

    # Remove all framework-specific attributes
    for tag in soup.find_all(True):
        tag.attrs = {key: val for key, val in tag.attrs.items() if not key.startswith('_ngcontent')}

    # Unwrap all custom tags like 'ms-cmark-node' and simple spans
    for custom_tag in soup.find_all(['ms-cmark-node', 'span']):
        custom_tag.unwrap()

    # Remove all HTML comments
    for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
        comment.extract()
        
    return soup

def is_rtl(text):
    return re.search(r'[\u0600-\u06FF]', text) is not None

def handle_paragraph(tag, doc):
    p = doc.add_paragraph()
    paragraph_text = tag.get_text(strip=True)
    if is_rtl(paragraph_text):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    for child in tag.children:
        if child.name in ['strong', 'b']:
            p.add_run(child.get_text()).bold = True
        elif child.name in ['em', 'i']:
            p.add_run(child.get_text()).italic = True
        elif isinstance(child, str):
            p.add_run(str(child))
        else:
            p.add_run(child.get_text())

# ... (handle_table and handle_code_block functions remain the same) ...
def handle_table(tag, doc):
    rows_data = []
    for row in tag.find_all('tr'):
        cols_data = [cell.get_text(strip=True) for cell in row.find_all(['th', 'td'])]
        rows_data.append(cols_data)
    if not rows_data: return
    table = doc.add_table(rows=len(rows_data), cols=len(rows_data[0]))
    table.style = 'Table Grid'
    if any(is_rtl(cell) for cell in rows_data[0]):
        table.direction = WD_TABLE_DIRECTION.RTL
    for i, row_data in enumerate(rows_data):
        row_cells = table.rows[i].cells
        for j, cell_text in enumerate(row_data):
            row_cells[j].text = cell_text
            if i == 0:
                for paragraph in row_cells[j].paragraphs:
                    for run in paragraph.runs: run.font.bold = True
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

def handle_code_block(tag, doc):
    code_text = tag.get_text()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    run = cell.paragraphs[0].add_run(code_text)
    run.font.name = 'Courier New'
    run.font.size = Pt(10)
    shading_xml = parse_xml(r'<w:shd {} w:fill="F0F0F0"/>'.format(nsdecls('w')))
    cell._tc.get_or_add_tcPr().append(shading_xml)
    doc.add_paragraph()


### UPDATED MAIN FUNCTION ###
def convert_html_to_docx(html_content, docx_path):
    # STEP 0: Check if file is writable before doing anything
    try:
        with open(docx_path, 'a') as f:
            pass
    except IOError:
        print(f"Permission denied to write to {docx_path}. It's likely open.")
        return False # Return False to signal failure

    doc = docx.Document()
    soup = BeautifulSoup(html_content, 'lxml')
    if not soup.body: return True

    cleaned_soup = clean_html(soup)

    for tag in cleaned_soup.body.find_all(['h1', 'h2', 'h3', 'p', 'div', 'ul', 'ol', 'pre', 'table']):
        text = tag.get_text(strip=True)
        if not text: continue
            
        if tag.name in ['h1', 'h2', 'h3']:
            level = int(tag.name[1])
            p = doc.add_heading(text, level=level)
            if is_rtl(text): p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif tag.name in ['p', 'div']:
            handle_paragraph(tag, doc)
        elif tag.name in ['ul', 'ol']:
            style = 'List Bullet' if tag.name == 'ul' else 'List Number'
            for li in tag.find_all('li', recursive=False):
                doc.add_paragraph(li.get_text(strip=True), style=style)
        elif tag.name == 'pre':
            handle_code_block(tag, doc)
        elif tag.name == 'table':
            handle_table(tag, doc)

    doc.save(docx_path)
    print(f"File converted successfully! Saved to: {docx_path}")
    return True # Return True to signal success