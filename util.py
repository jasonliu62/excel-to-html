def clean_text(text):
    """Clean and format text for HTML output"""
    if not text:
        return ''
    text = text.replace('<br/>', '___BR___')
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('___BR___', '<br/>')
    return text 