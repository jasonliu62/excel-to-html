import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog, 
                            QVBoxLayout, QWidget, QLabel, QProgressBar, QMessageBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from update import DocxProcessor

class ConversionWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, file_path, content_type):
        super().__init__()
        self.file_path = file_path
        self.content_type = content_type
        self.processor = DocxProcessor()

    def run(self):
        try:
            self.progress.emit(10)
            # Process the document
            html_content = self.processor.process_docx(self.file_path, self.content_type)
            self.progress.emit(70)

            # Save output
            with open('output.html', 'w', encoding='utf-8') as f:
                f.write(html_content)
            self.progress.emit(100)
            
            self.finished.emit('Conversion completed successfully!')
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('DOCX to HTML Converter')
        self.setGeometry(100, 100, 500, 400)

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Add title label
        title_label = QLabel('DOCX to HTML Converter')
        title_label.setStyleSheet('font-size: 16px; font-weight: bold; margin: 10px;')
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # Add description label
        desc_label = QLabel('Convert DOCX files to HTML format')
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setStyleSheet('margin: 10px;')
        layout.addWidget(desc_label)

        # Add buttons
        button_style = '''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border: none;
                border-radius: 5px;
                font-size: 14px;
                margin: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        '''

        self.auto_button = QPushButton('Convert Auto-Detect')
        self.auto_button.setStyleSheet(button_style)
        self.auto_button.clicked.connect(lambda: self.select_file('auto'))
        layout.addWidget(self.auto_button)

        self.table_button = QPushButton('Convert Tables Only')
        self.table_button.setStyleSheet(button_style)
        self.table_button.clicked.connect(lambda: self.select_file('table'))
        layout.addWidget(self.table_button)

        self.text_button = QPushButton('Convert Text Only')
        self.text_button.setStyleSheet(button_style)
        self.text_button.clicked.connect(lambda: self.select_file('text'))
        layout.addWidget(self.text_button)

        # Add status label
        self.status_label = QLabel('No file selected')
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet('margin: 10px;')
        layout.addWidget(self.status_label)

        # Add progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Add some spacing at the bottom
        layout.addStretch()

    def select_file(self, content_type):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select DOCX File",
            "",
            "DOCX Files (*.docx)"
        )
        
        if file_name:
            self.status_label.setText('Processing...')
            self.auto_button.setEnabled(False)
            self.table_button.setEnabled(False)
            self.text_button.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)

            # Create and start worker thread
            self.worker = ConversionWorker(file_name, content_type)
            self.worker.finished.connect(self.conversion_finished)
            self.worker.error.connect(self.conversion_error)
            self.worker.progress.connect(self.update_progress)
            self.worker.start()

    def conversion_finished(self, message):
        self.status_label.setText(message)
        self.auto_button.setEnabled(True)
        self.table_button.setEnabled(True)
        self.text_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        QMessageBox.information(self, 'Success', 'Conversion completed! Check output.html')

    def conversion_error(self, error_message):
        self.status_label.setText(f'Error: {error_message}')
        self.auto_button.setEnabled(True)
        self.table_button.setEnabled(True)
        self.text_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, 'Error', f'An error occurred: {error_message}')

    def update_progress(self, value):
        self.progress_bar.setValue(value)

def main():
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
