from unidecode import unidecode

def clean_text(text):
    """Clean and format text for HTML output"""
    if text is None or text.strip() == '':
        return '&#160;'
    text = unidecode(text)
    text = text.replace('<br/>', '___BR___')
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('___BR___', '<br/>')
    return text 