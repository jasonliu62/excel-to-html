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
            for element in list(body):
                if element.tag == f'{{{ns["w"]}}}p':
                    html_parts.append(self.text_processor.process_paragraph(element, ns))
                elif element.tag == f'{{{ns["tbl"]}}}tbl':
                    html_parts.append(self.table_processor.process_table_element(element, ns))
            return '\n'.join(html_parts)
        elif content_type == 'table':
            return self.table_processor.process_table(extract_dir)
        elif content_type == 'text':
            return self.text_processor.process_text(extract_dir)
        else:
            raise ValueError("Invalid content type. Must be 'auto', 'table', or 'text'") 