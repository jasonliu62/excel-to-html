import os
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
from table import TableProcessor
from text import TextProcessor

class DocxProcessor:
    def __init__(self):
        self.table_processor = TableProcessor()
        self.text_processor = TextProcessor()
        self.namespaces = {
            # TODO: This dictionary is hardcoded, it may be necessary to use the docx XML  to process all namespaces
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'tbl': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
             'v': "urn:schemas-microsoft-com:vml",
            'o': "urn:schemas-microsoft-com:office:office",
        }

    def extract_docx_to_xml(self, docx_path):
        """Extract DOCX file to XML files"""
        extract_dir = os.path.splitext(docx_path)[0] + '_extracted'
        if not os.path.exists(extract_dir):
            os.makedirs(extract_dir)
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        return extract_dir
    

    def process_docx(self, docx_path, content_type='auto'):
        """
        Process DOCX file based on content type
        content_type can be: 'auto', 'table', 'text'
        """
        extract_dir = self.extract_docx_to_xml(docx_path)
        document_xml = os.path.join(extract_dir, 'word', 'document.xml')
        tree = ET.parse(document_xml)
        root = tree.getroot()
        ns = self.namespaces

        if content_type == 'auto':
            # Parse the entire XML file and route each paragraph or table to the appropriate processor
            html_parts = []
            body = root.find(f'{{{ns["w"]}}}body', self.namespaces)

            # Define business logic for handling nested lists
            list_stack = [] # Stack to manage nested lists consisting of a tuple (list_tag, ilvl)
            prev_ilvl = -1
            prev_list_tag = None


            for element in list(body):
                if element.tag == f'{{{ns["w"]}}}p':

                    # TODO: List handling logic
                    # It is pretty complex and will need to be repeated in cells
                    # Further research to see if this can be refactored into a separate reusable function

                    if self.text_processor.is_list_paragraph(element, ns):
                        ilvl = self.text_processor.get_list_level(element, ns)
                        list_tag = self.text_processor.get_list_tag(element, ns)
                        while (prev_ilvl < ilvl):
                            html_parts.append(f'<{list_tag}>')
                            list_stack.append((list_tag, ilvl))
                            prev_ilvl += 1
                            prev_list_tag = list_tag
                        while (prev_ilvl > ilvl):
                            tag, _ = list_stack.pop()
                            html_parts.append(f'</{tag}>')
                            prev_ilvl -= 1
                        if prev_list_tag != None and prev_list_tag != list_tag:
                            if list_stack:
                                tag, _ = list_stack.pop()
                                html_parts.append(f'</{tag}>')
                            html_parts.append(f'<{list_tag}>')
                            list_stack.append((list_tag, ilvl))
                            prev_list_tag = list_tag
                        li_content = []
                        for child in list(element):
                            if child.tag == f'{{{ns["w"]}}}r':
                                li_content.append(self.text_processor.process_run(child, ns))
                            elif child.tag == f'{{{ns["w"]}}}hyperlink':
                                li_content.append(self.text_processor.process_hyperlink(child, ns, extract_dir))
                      
                        html_parts.append(f'<li>{"".join(li_content)}</li>')
                        continue
                    else:
                        while list_stack:
                            tag, _ = list_stack.pop()
                            html_parts.append(f'</{tag}>')
                        prev_ilvl = -1
                        prev_list_tag = None
                        html_parts.append(self.text_processor.process_paragraph(element, ns, extract_dir))
                elif element.tag == f'{{{ns["tbl"]}}}tbl':
                    # Assuming list starts outside of table and ends before table starts
                    # May need more robust logic if this assumption does not hold and tables can be inside lists
                    while list_stack:
                        tag, _ = list_stack.pop()
                        html_parts.append(f'</{tag}>')
                    prev_ilvl = -1
                    prev_list_tag = None
                    html_parts.append(self.table_processor.process_table_element(element, ns, extract_dir))
            while list_stack:
                tag, _ = list_stack.pop()
                html_parts.append(f'</{tag}>')
            prev_ilvl = -1
            prev_list_tag = None
            return '\n'.join(html_parts)
        elif content_type == 'table':
            return self.table_processor.process_table(extract_dir)
        elif content_type == 'text':
            return self.text_processor.process_text(extract_dir)
        else:
            raise ValueError("Invalid content type. Must be 'auto', 'table', or 'text'") 