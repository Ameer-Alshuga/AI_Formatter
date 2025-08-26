# main.py - The Conversion Engine

from bs4 import BeautifulSoup
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_DIRECTION
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

# جميع دوال المعالجة (handle_paragraph, handle_table, handle_code_block)
# تبقى كما هي تماماً. لقد حذفتها من هنا للاختصار،
# لكن تأكد من أنها لا تزال موجودة في ملفك.
# تحتاج فقط إلى استبدال دالة convert_html_to_docx 
# وجزء الاختبار في نهاية الملف.

def handle_paragraph(tag, doc):
    """Handles paragraph tags (<p>) and supports bold and italic text within them."""
    p = doc.add_paragraph()
    if tag.get('dir') == 'rtl':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for child in tag.children:
        if child.name in ['strong', 'b']:
            p.add_run(child.get_text()).bold = True
        elif child.name in ['em', 'i']:
            p.add_run(child.get_text()).italic = True
        else:
            p.add_run(str(child))

def handle_table(tag, doc):
    """Processes HTML tables and converts them to Word tables, with RTL support."""
    rows_data = []
    for row in tag.find_all('tr'):
        cols_data = [cell.get_text(strip=True) for cell in row.find_all(['th', 'td'])]
        rows_data.append(cols_data)
    if not rows_data: return
    table = doc.add_table(rows=len(rows_data), cols=len(rows_data[0]))
    table.style = 'Table Grid'
    if tag.get('dir') == 'rtl':
        table.direction = WD_TABLE_DIRECTION.RTL
    for i, row_data in enumerate(rows_data):
        row_cells = table.rows[i].cells
        for j, cell_text in enumerate(row_data):
            row_cells[j].text = cell_text
            if i == 0:
                for p in row_cells[j].paragraphs:
                    for run in p.runs: run.font.bold = True
    doc.add_paragraph()

def handle_code_block(tag, doc):
    """Handles code blocks by placing them in a single-cell table with a grey background."""
    code_text = tag.get_text()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    run = cell.paragraphs[0].add_run(code_text)
    run.font.name = 'Courier New'
    run.font.size = Pt(10)
    shading_xml = parse_xml(r'<w:shd {} w:fill="F0F0F0"/>'.format(nsdecls('w')))
    cell._tc.get_or_add_tcPr().append(shading_xml)
    doc.add_paragraph()


# --- الدالة المعدلة ---
def convert_html_to_docx(html_content, docx_path):
    """
    The main function that converts an HTML content string into a DOCX file.
    """
    doc = docx.Document()
    soup = BeautifulSoup(html_content, 'lxml')

    if not soup.body:
        return # اخرج إذا كان الـ HTML غير صالح أو لا يحتوي على body

    for tag in soup.body.find_all(recursive=False):
        if tag.name == 'h1':
            p = doc.add_heading(tag.get_text(strip=True), level=1)
            if tag.get('dir') == 'rtl': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif tag.name == 'h2':
            p = doc.add_heading(tag.get_text(strip=True), level=2)
            if tag.get('dir') == 'rtl': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif tag.name == 'p':
            handle_paragraph(tag, doc)
        elif tag.name == 'ul':
            for li in tag.find_all('li'):
                doc.add_paragraph(li.get_text(strip=True), style='List Bullet')
        elif tag.name == 'pre':
            handle_code_block(tag, doc)
        elif tag.name == 'table':
            handle_table(tag, doc)

    doc.save(docx_path)
    print(f"File converted successfully! Saved to: {docx_path}")

# --- جزء الاختبار المعدل ---
# هذا الجزء الآن يقرأ الملف أولاً ويمرر المحتوى كنص،
# محاكياً كيفية استخدام مراقب الحافظة له.
if __name__ == '__main__':
    # هذا الجزء مخصص فقط لاختبار main.py مباشرة
    print("Testing the conversion engine...")
    with open('sample.html', 'r', encoding='utf-8') as f:
        sample_html_content = f.read()
    
    convert_html_to_docx(sample_html_content, 'output.docx')