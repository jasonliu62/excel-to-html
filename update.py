import os
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QLabel
from PyQt5.QtCore import Qt
import sys
import re

class DocxProcessor:
    def __init__(self):
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'tbl': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }

    def extract_docx_to_xml(self, docx_path):
        """Extract DOCX file to XML files"""
        extract_dir = os.path.splitext(docx_path)[0] + '_extracted'
        if not os.path.exists(extract_dir):
            os.makedirs(extract_dir)
        
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        return extract_dir
    
    def process_plain_text(self, docx_path):
        """Extract and "parse plain text" from DOCX file, returning a list of paragraphs."""
        """Note that this method will only process text that is in a <w:p> element. Additional methods may be needed to handle headers, footers, textboxes/shapes, comments,"""
        """endnotes, fieldcodes, special elements, etc."""
        extract_dir = self.extract_docx_to_xml(docx_path)
        document_xml = os.path.join(extract_dir, 'word', 'document.xml')
        tree = ET.parse(document_xml)
        root = tree.getroot()
        ns = self.namespaces

        paragraphs = []
        for p in root.findall('.//w:p', ns):
            p_pr = p.find('w:pPr', ns)
            # Extract pagagraph style
            if p_pr is not None:
                # Call helper function here
                style = {}
            style_str = '; '.join(f'{k}: {v}' for k, v in style.items())
            paragraph = [f'<p style="{style_str}">']
            for run in p.findall('.//w:r', ns):
                # Further processing of CSS on runs may be needed here
                run_style = self._get_run_style(run, ns)
      
                runpr = run.find('w:rPr', ns)
                is_bold = runpr is not None and runpr.find('w:b', ns) is not None
                run_text = ''
                for child in list(run):
                    tag = child.tag
                    if tag == f'{{{ns["w"]}}}t':
                        #TODO: Put run_text in a span and extract styles and apply them

                        run_text += self.clean_cell_text(child.text or '')
                    elif tag == f'{{{ns["w"]}}}br':
                        run_text += '<br/>'
                if run_text:
                    if is_bold:
                        run_text = f'<b>{run_text}</b>'
                    if run_style:
                        run_text = f'<span style="{run_style}">{run_text}</span>'
                    paragraph.append(run_text)
                # note nesting may need to be properly handled in future
            paragraph.append('</p>')
            text = ''.join(paragraph)
            text = text.replace('–', '&#8211;').replace('—', '&#8212;')
            paragraphs.append(text)
        return '\n\n'.join(paragraphs)            

    def process_table(self, docx_path):
        """Extract and parse document.xml directly, build HTML table with dynamic structure and inline styles, mapping Word widths to percentages if possible, and only outputting styles/attributes present in the XML."""
        extract_dir = self.extract_docx_to_xml(docx_path)
        document_xml = os.path.join(extract_dir, 'word', 'document.xml')
        tree = ET.parse(document_xml)
        root = tree.getroot()
        ns = self.namespaces

        html_tables = []
        for tbl in root.findall('.//w:tbl', ns):
            tbl_pr = tbl.find('w:tblPr', ns)
            total_width_twips = None
            if tbl_pr is not None:
                tblw = tbl_pr.find('w:tblW', ns)
                if tblw is not None and tblw.get(f'{{{ns["w"]}}}type') == 'dxa':
                    w = tblw.get(f'{{{ns["w"]}}}w')
                    if w and w.isdigit():
                        total_width_twips = int(w)

            html_table = [
                '<table cellpadding="0" cellspacing="0" style="font: 10pt Times New Roman, Times, Serif; border-collapse: collapse; width: 100%">'
            ]
            for tr_idx, tr in enumerate(tbl.findall('w:tr', ns)):
                row_style = self._get_row_style(tr, ns)
                # Always add vertical-align: bottom for every row
                if row_style:
                    row_style = f'vertical-align: bottom; {row_style}'
                else:
                    row_style = 'vertical-align: bottom;'
                tcs = tr.findall('w:tc', ns)
                row_cells = []
                last_cell_double_underline = False
                for tc_idx, tc in enumerate(tcs):
                    cell_text = self._get_cell_text(tc, ns)
                    row_cells.append(cell_text)
                # Check if all cells are empty
                all_empty = all(cell.strip() == '' for cell in row_cells)
                # Add min-height if all cells are empty and fill with &nbsp;
                tr_style = row_style
                if all_empty:
                    tr_style += ' min-height: 12pt;'
                    row_cells = ['&#160;' for _ in row_cells]
                html_table.append(f'<tr{f" style=\"{tr_style}\"" if tr_style else ""}>')
                for tc_idx, tc in enumerate(tcs):
                    cell_text = row_cells[tc_idx]
                    cell_style, colspan = self._get_cell_style(tc, ns, total_width_twips, tc_idx, cell_text)
                    tag = 'td'
                    attrs = []
                    if colspan > 1:
                        attrs.append(f'colspan="{colspan}"')
                    if cell_style:
                        attrs.append(f'style="{cell_style}"')
                    attr_str = ' '.join(attrs)
                    html_table.append(f'<{tag} {attr_str}>{cell_text}</{tag}>')
                    # Check for double underline in last cell
                    if tc_idx == len(tcs) - 1:
                        props = tc.find('w:tcPr', ns)
                        if props is not None:
                            borders = props.find('w:tcBorders', ns)
                            if borders is not None:
                                bottom = borders.find('w:bottom', ns)
                                if bottom is not None and bottom.get(f'{{{ns["w"]}}}val') == 'double':
                                    last_cell_double_underline = True
                # Add extra <td> with double border if needed
                if last_cell_double_underline:
                    html_table.append('<td style="border-bottom: Black 2.5pt double;"></td>')
                html_table.append('</tr>')
            html_table.append('</table>')
            html_tables.append('\n'.join(html_table))
        return '\n\n'.join(html_tables)
    
    # def _get_paragraph_style(self, p, ns):
    #     props = p.find('w:pPr', ns)
    #     if props is not None and props.find

    def _get_run_style(self, r, ns):
        props = r.find('w:rPr', ns)
        style = []
        if props is not None:
            # Check for vanish
            if props.find('w:vanish', ns) is not None:
                style.append('display: none;')
            
            # Check for font
            rfonts = props.find('w:rFonts', ns)
            if rfonts is not None:
                font_css = []
                east_asian = rfonts.get(f'{{{ns["w"]}}}eastAsia')
                ascii = rfonts.get(f'{{{ns["w"]}}}ascii')
                if ascii:
                    font_css.append(f'font-family: {ascii};')
                if east_asian:
                    font_css.append(f'font-variant-east-asian: {east_asian};')
                #TODO: add support for hansi, cs
                if font_css:
                    style.append(' '.join(font_css))
            
            # Check for bold
            if props.find('w:b', ns) is not None:
                style.append('font-weight: bold;')
            # Check for italic
            if props.find('w:i', ns) is not None:
                style.append('font-style: italic;')
            # Check for underline
            u = props.find('w:u', ns)
            if u is not None:
                val = u.get(f'{{{ns["w"]}}}val')
                if val == 'single':
                    style.append('text-decoration: underline;')
                elif val == 'double':
                    style.append('text-decoration: underline double;')
        return ' '.join(style)

    def _get_row_style(self, tr, ns):
        props = tr.find('w:trPr', ns)
        style = []
        if props is not None:
            shd = props.find('w:shd', ns)
            if shd is not None:
                fill = shd.get(f'{{{ns["w"]}}}fill')
                if fill and fill != 'auto' and fill != 'FFFFFF':
                    style.append(f'background-color: #{fill};')
            # vertical-align is always bottom in your sample, but only add if present
            # (Word rarely stores this, so usually omitted)
        return ' '.join(style)

    def _get_cell_style(self, tc, ns, total_width_twips=None, tc_idx=0, cell_text=None):
        props = tc.find('w:tcPr', ns)
        style = []
        colspan = 1
        width_percent = None
        text_align_found = False
        bold_found = False
        # Alignment: Try to get from w:jc in cell or paragraph
        align = None
        if props is not None:
            # Colspan
            gridspan = props.find('w:gridSpan', ns)
            if gridspan is not None:
                colspan = int(gridspan.get(f'{{{ns["w"]}}}val', '1'))
            # Width
            tcw = props.find('w:tcW', ns)
            if tcw is not None:
                w = tcw.get(f'{{{ns["w"]}}}w')
                if w and w.isdigit() and int(w) > 0:
                    w_twips = int(w)
                    if total_width_twips and total_width_twips > 0:
                        width_percent = 100 * w_twips / total_width_twips
                        style.append(f'width: {width_percent:.0f}%;')
                    else:
                        width_pt = w_twips / 20
                        style.append(f'width: {width_pt:.1f}pt;')
            # Background
            shd = props.find('w:shd', ns)
            if shd is not None:
                fill = shd.get(f'{{{ns["w"]}}}fill')
                if fill and fill != 'auto' and fill != 'FFFFFF':
                    style.append(f'background-color: #{fill};')
            # Borders
            borders = props.find('w:tcBorders', ns)
            if borders is not None:
                for side in ['top', 'bottom', 'left', 'right']:
                    el = borders.find(f'w:{side}', ns)
                    if el is not None:
                        val = el.get(f'{{{ns["w"]}}}val')
                        sz = el.get(f'{{{ns["w"]}}}sz')
                        color = el.get(f'{{{ns["w"]}}}color', '000000')
                        if val in ['single', 'double']:
                            thickness = '1pt' if val == 'single' else '2.5pt double'
                            style.append(f'border-{side}: Black {thickness} solid;')
            # Alignment from cell properties
            jc = props.find('w:jc', ns)
            if jc is not None:
                align = jc.get(f'{{{ns["w"]}}}val', None)
        # If not found in cell, look for first paragraph alignment
        if align is None:
            first_p = tc.find('w:p', ns)
            if first_p is not None:
                ppr = first_p.find('w:pPr', ns)
                if ppr is not None:
                    jc = ppr.find('w:jc', ns)
                    if jc is not None:
                        align = jc.get(f'{{{ns["w"]}}}val', None)
        # Map Word alignment to CSS
        if align is not None:
            if align == 'center':
                style.append('text-align: center;')
                text_align_found = True
            elif align == 'left':
                style.append('text-align: left;')
                text_align_found = True
            elif align == 'right':
                style.append('text-align: right;')
                text_align_found = True
            elif align == 'both':
                style.append('text-align: justify;')
                text_align_found = True
        # Padding (tcMar)
        if props is not None:
            tcmar = props.find('w:tcMar', ns)
            if tcmar is not None:
                for side in ['top', 'bottom', 'left', 'right']:
                    mar = tcmar.find(f'w:{side}', ns)
                    if mar is not None:
                        w = mar.get(f'{{{ns["w"]}}}w')
                        if w and w.isdigit():
                            pt = int(w) / 20
                            style.append(f'padding-{side}: {pt:.1f}pt;')
            # Check for bold in cell properties
            cell_bold = props.find('w:b', ns)
            if cell_bold is not None:
                bold_found = True

        # Check for bold in run properties
        for run in tc.findall('.//w:r', ns):
            runpr = run.find('w:rPr', ns)
            if runpr is not None and runpr.find('w:b', ns) is not None:
                bold_found = True
                break

        # If any part of the cell is bold, apply font-weight: bold
        if bold_found:
            if not any(s.startswith('font-weight:') for s in style):
                style.append('font-weight: bold;')

        # Force text-align: left for cells with only ')' or '$'
        if cell_text is not None and cell_text.strip() in [')', '$']:
            style = [s for s in style if not s.startswith('text-align:')]
            style.append('text-align: left;')
        # Default text-align: right if not specified, except for first column
        elif not text_align_found:
            if tc_idx == 0:
                style.append('text-align: left;')
            else:
                style.append('text-align: right;')
        return ' '.join(style), colspan

    def _get_cell_text(self, tc, ns):
        # Walk through each <w:r> in order, and for each run, process its children in order
        texts = []
        for run in tc.findall('.//w:r', ns):
            runpr = run.find('w:rPr', ns)
            is_bold = runpr is not None and runpr.find('w:b', ns) is not None
            run_text = ''
            for child in list(run):
                tag = child.tag
                if tag == f'{{{ns["w"]}}}t':
                    run_text += child.text or ''
                elif tag == f'{{{ns["w"]}}}br':
                    run_text += '<br/>'
            if run_text:
                texts.append(run_text)
        text = self.clean_cell_text(''.join(texts))
        # Replace single hyphen or en dash with en dash entity
        if text.strip() in ['-', '–']:
            text = '&#8211;'
        return text

    def clean_cell_text(self, text):
        """Clean and format cell text"""
        if not text:
            return ''
        
        # Remove extra whitespace
        text = re.sub(r'\s+', ' ', text.strip())
        
        # Handle special characters
        text = text.replace('<br/>', '___BR___')
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('___BR___', '<br/>')
        
        return text

    def extract_table_data(self, docx_path):
        """Extract table data from DOCX file"""
        doc = Document(docx_path)
        tables_data = []
        
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = [self.clean_cell_text(cell.text) for cell in row.cells]
                table_data.append(row_data)
            tables_data.append(table_data)
        
        return tables_data

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.processor = DocxProcessor()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('DOCX Table Converter')
        self.setGeometry(100, 100, 400, 200)

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Add label
        self.label = QLabel('Select a DOCX file to convert tables')
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        # Add button
        self.button = QPushButton('Select DOCX File')
        self.button.clicked.connect(self.select_file)
        layout.addWidget(self.button)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select DOCX File",
            "",
            "DOCX Files (*.docx)"
        )
        
        if file_name:
            try:
                # Process the table
                html_content = self.processor.process_plain_text(file_name)
                
                # Save to output.html
                with open('output.html', 'w', encoding='utf-8') as f:
                    f.write(html_content)
                
                self.label.setText('Conversion completed! Check output.html')
            except Exception as e:
                self.label.setText(f'Error: {str(e)}')

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main() 