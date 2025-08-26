from bs4 import BeautifulSoup
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_DIRECTION
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

def handle_paragraph(tag, doc):
    """
    Handles paragraph tags (<p>) and supports bold and italic text within them.
    """
    p = doc.add_paragraph()
    
    # Check for right-to-left text direction (for Arabic)
    if tag.get('dir') == 'rtl':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # Iterate over each child element within the paragraph tag
    for child in tag.children:
        if child.name == 'strong' or child.name == 'b':
            p.add_run(child.get_text()).bold = True
        elif child.name == 'em' or child.name == 'i':
            p.add_run(child.get_text()).italic = True
        else:
            # This handles plain text nodes
            p.add_run(str(child))


def handle_table(tag, doc):
    """
    Processes HTML tables and converts them to Word tables, with RTL support.
    """
    rows_data = []
    # Extract all rows from the table
    for row in tag.find_all('tr'):
        cols_data = []
        # Extract all cells (th or td) from each row
        for cell in row.find_all(['th', 'td']):
            cols_data.append(cell.get_text(strip=True))
        rows_data.append(cols_data)

    if not rows_data:
        return # Don't create an empty table

    # Create a table in Word with the appropriate number of rows and columns
    num_rows = len(rows_data)
    num_cols = len(rows_data[0]) if num_rows > 0 else 0
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid' # Apply a basic table style

    # Check for table direction (for Arabic)
    if tag.get('dir') == 'rtl':
        table.direction = WD_TABLE_DIRECTION.RTL

    # Populate the table cells with data
    for i, row_data in enumerate(rows_data):
        row_cells = table.rows[i].cells
        for j, cell_text in enumerate(row_data):
            row_cells[j].text = cell_text
            # Make the first row (headers) bold
            if i == 0:
                for paragraph in row_cells[j].paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
    
    # Add a paragraph after the table for better spacing
    doc.add_paragraph()


def handle_code_block(tag, doc):
    """
    Handles code blocks by placing them in a single-cell table with a grey background.
    """
    code_text = tag.get_text()
    
    # Create a table with one row and one column
    table = doc.add_table(rows=1, cols=1)
    
    # Get the single cell from the table
    cell = table.cell(0, 0)
    
    # Add the code text to the cell
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(code_text)
    
    # Apply monospace font formatting
    run.font.name = 'Courier New'
    run.font.size = Pt(10)

    # --- Section for cell background color ---
    # This code modifies the cell's underlying XML to add a background color
    shading_xml = parse_xml(r'<w:shd {} w:fill="F0F0F0"/>'.format(nsdecls('w'))) # F0F0F0 is light grey
    cell._tc.get_or_add_tcPr().append(shading_xml)
    
    # Add a paragraph after the code block for spacing
    doc.add_paragraph()


def convert_html_to_docx(html_path, docx_path):
    """
    The main function that reads an HTML file and converts it to a DOCX file.
    """
    # Create a new Word document
    doc = docx.Document()

    # Read the content of the HTML file
    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html_content, 'lxml')

    # Iterate over all top-level tags within the HTML body
    for tag in soup.body.find_all(recursive=False):
        if tag.name == 'h1':
            p = doc.add_heading(tag.get_text(strip=True), level=1)
            if tag.get('dir') == 'rtl':
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        elif tag.name == 'h2':
            p = doc.add_heading(tag.get_text(strip=True), level=2)
            if tag.get('dir') == 'rtl':
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        elif tag.name == 'p':
            handle_paragraph(tag, doc)

        elif tag.name == 'ul':
            for li in tag.find_all('li'):
                doc.add_paragraph(li.get_text(strip=True), style='List Bullet')
        
        elif tag.name == 'pre':
            # Call the new dedicated function for code blocks
            handle_code_block(tag, doc)

        elif tag.name == 'table':
            handle_table(tag, doc)

    # Save the final document
    doc.save(docx_path)
    print(f"File converted successfully! Saved to: {docx_path}")


# Script entry point
if __name__ == '__main__':
    html_file = 'sample.html'
    docx_file = 'output.docx'
    convert_html_to_docx(html_file, docx_file)