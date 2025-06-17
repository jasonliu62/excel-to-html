# DOCX to HTML Converter

A Python-based application that converts Microsoft Word (DOCX) files to HTML format with precise formatting preservation. The converter handles both text and table content, maintaining styles, alignments, and formatting from the original document.

## Features

- **Multiple Conversion Modes**:
  - Auto-Detect: Automatically detects and converts both text and tables
  - Tables Only: Converts only table content
  - Text Only: Converts only text content

- **Table Formatting**:
  - Preserves table structure and layout
  - Maintains cell alignments and borders
  - Handles merged cells and complex table structures
  - Preserves text formatting within table cells

- **Text Formatting**:
  - Maintains font styles and sizes
  - Preserves text alignment and spacing
  - Handles special characters and symbols
  - Supports text decorations (bold, italic, underline)

- **User-Friendly Interface**:
  - Simple and intuitive GUI
  - Progress tracking
  - Error handling and notifications
  - File selection dialog

## Requirements

- Python 3.x
- PyQt5
- python-docx
- Required Python packages:
  ```
  PyQt5
  python-docx
  ```

## Installation

1. Clone the repository:
   ```bash
   git clone [repository-url]
   cd excel-to-html
   ```

2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```

2. Using the GUI:
   - Click "Convert Auto-Detect" to convert both text and tables
   - Click "Convert Tables Only" to convert only table content
   - Click "Convert Text Only" to convert only text content
   - Select your DOCX file when prompted
   - Wait for the conversion to complete
   - The converted HTML will be saved as 'output.html'

## Project Structure

- `main.py`: Main application file with GUI implementation
- `update.py`: Core document processing logic
- `table.py`: Table-specific processing and conversion
- `text.py`: Text-specific processing and conversion

## Output

The converter generates an HTML file with:
- Preserved document structure
- Maintained formatting and styles
- Clean, semantic HTML markup
- Inline CSS for styling

## Notes

- The converter preserves Word's default left alignment when no explicit alignment is specified
- Table borders are handled with special attention to empty cells vs. content cells
- The output HTML is optimized for readability and browser compatibility

## License

MIT