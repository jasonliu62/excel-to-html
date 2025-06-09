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

    def get_table_style(self):
        """Get the table style matching the template"""
        return '''
        <style>
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 20px 0;
                font-family: Arial, sans-serif;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 12px;
                text-align: left;
            }
            th {
                background-color: #f5f5f5;
                font-weight: bold;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            tr:hover {
                background-color: #f5f5f5;
            }
        </style>
        '''

    def process_table(self, docx_path):
        """Process table from DOCX file and convert to HTML format"""
        doc = Document(docx_path)
        html_tables = []
        
        # Add style section
        html_tables.append(self.get_table_style())
        
        for table in doc.tables:
            html_table = ['<table>']
            
            # Process header row
            if table.rows:
                header_row = table.rows[0]
                html_table.append('<thead>')
                html_table.append('<tr>')
                for cell in header_row.cells:
                    # Clean and format cell text
                    cell_text = self.clean_cell_text(cell.text)
                    html_table.append(f'<th>{cell_text}</th>')
                html_table.append('</tr>')
                html_table.append('</thead>')
            
            # Process data rows
            html_table.append('<tbody>')
            for row in table.rows[1:]:
                html_table.append('<tr>')
                for cell in row.cells:
                    # Clean and format cell text
                    cell_text = self.clean_cell_text(cell.text)
                    html_table.append(f'<td>{cell_text}</td>')
                html_table.append('</tr>')
            html_table.append('</tbody>')
            
            html_table.append('</table>')
            html_tables.append('\n'.join(html_table))
        
        return '\n\n'.join(html_tables)

    def clean_cell_text(self, text):
        """Clean and format cell text"""
        if not text:
            return ''
        
        # Remove extra whitespace
        text = re.sub(r'\s+', ' ', text.strip())
        
        # Handle special characters
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        
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
                html_content = self.processor.process_table(file_name)
                
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