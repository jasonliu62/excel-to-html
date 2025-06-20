import os
import xml.etree.ElementTree as ET
import re
from util import clean_text

class TextProcessor:
    def __init__(self):
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'tbl': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }

    def process_text(self, extract_dir):
        document_xml = os.path.join(extract_dir, 'word', 'document.xml')
        tree = ET.parse(document_xml)
        root = tree.getroot()
        body = root.find('w:body', self.namespaces)
        ns = self.namespaces

        nodes = []
        for child in list(body):
            tag = child.tag
            if tag == f'{{{ns["w"]}}}p':
                para_html = self.process_paragraph(child, ns, extract_dir)
                nodes.append(para_html)
            elif tag == f'{{{ns["w"]}}}sectPr':
                # Skip section properties
                continue
            # TODO: Handle headers, footers, textboxes/shapes, comments, endnotes, fieldcodes, special elements, etc.
        return '\n\n'.join(nodes)

    def is_list_paragraph(self, p, ns):
        p_pr = p.find('w:pPr', ns)
        if p_pr is not None and p_pr.find('w:numPr', ns) is not None:
            return True
        return False
    
    def get_list_level(self, p, ns):
        p_pr = p.find('w:pPr', ns)
        if p_pr is not None and p_pr.find('w:numPr', ns) is not None:
            numPr = p_pr.find('w:numPr', ns)
            ilvl = numPr.find('w:ilvl', ns)
            if ilvl is not None:
                return int(ilvl.get(f'{{{ns["w"]}}}val', '0'))
        return 0
    
    def get_list_tag(self, p, ns):
        # TODO: Determine if the list is ordered or unordered based on numbering.xml
        # For now, treat all lists as unordered lists
        return 'ul'
    
    def process_paragraph(self, p, ns, extract_dir):
        p_pr = p.find('w:pPr', ns)
        style = self._get_paragraph_style(p, ns) if p_pr is not None else ''
        paragraph = [f'<p style="{style}">']
        for child in list(p):
            tag = child.tag
            if tag == f'{{{ns["w"]}}}pPr':
                continue
            if tag == f'{{{ns["w"]}}}r':
                run_html = self.process_run(child, ns)
                paragraph.append(run_html)
            elif tag == f'{{{ns["w"]}}}hyperlink':
                hyperlink_html = self.process_hyperlink(child, ns, extract_dir)
                paragraph.append(hyperlink_html)
        paragraph.append('</p>')
        text = ''.join(paragraph)
        return text

    def process_hyperlink(self, hyperlink, ns, extract_dir):
        # TODO: Generate link URL from r:ID in <w:hyperlink>, find the link in document.xml.rels
        document_xml_rels = os.path.join(extract_dir, 'word', '_rels', 'document.xml.rels')
        rels_tree = ET.parse(document_xml_rels)
        rels_root = rels_tree.getroot()
        relationships = rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
        r_id = hyperlink.get(f'{{{ns["r"]}}}id')
        link = ''
        for rel in relationships:
            if rel.get('Id') == r_id:
                link = rel.get('Target')
                break
        html = ''
        for run in hyperlink.findall('w:r', ns):
            html += self.process_run(run, ns)
        html = f'<a href="{link}">{html}</a>'
        return html

    def process_run(self, run, ns):
        # Refactor code from process_plain_text into this function
        run_style = self._get_run_style(run, ns)
        runpr = run.find('w:rPr', ns)
        is_bold = runpr is not None and runpr.find('w:b', ns) is not None
        run_text = ''
        for child in list(run):
            tag = child.tag
            if tag == f'{{{ns["w"]}}}rPr':
                continue
            if tag == f'{{{ns["w"]}}}t':
                run_text += clean_text(child.text or '')
            elif tag == f'{{{ns["w"]}}}br':
                run_text += '<br/>'
        if is_bold:
            run_text = f'<b>{run_text}</b>'
        if run_style:
            run_text = f'<span style="{run_style}">{run_text}</span>'
        return run_text

    def _get_paragraph_style(self, p, ns):
        props = p.find('w:pPr', ns)
        style = []
        if props is not None:
            pstyle = props.find('w:pStyle', ns)
            if pstyle is not None:
                # TODO: Handle pStyle mapping to CSS
                pass
            rPr = props.find('w:rPr', ns)
            if rPr is not None:
                run_style = self._get_run_style(rPr, ns)
                if run_style:
                    style.append(run_style)
            jc = props.find('w:jc', ns)
            if jc is not None:
                align = jc.get(f'{{{ns["w"]}}}val', None)
                if align == 'center':
                    style.append('text-align: center;')
                elif align == 'left':
                    style.append('text-align: left;')
                elif align == 'right':
                    style.append('text-align: right;')
                elif align == 'both':
                    style.append('text-align: justify;')
                elif align == 'start':
                    style.append('text-align: start;')
                elif align == 'end':
                    style.append('text-align: end;')
                else:
                    style.append('text-align: justify;')
            spacing = props.find('w:spacing', ns)
            if spacing is not None:
                before = spacing.get(f'{{{ns["w"]}}}before')
                after = spacing.get(f'{{{ns["w"]}}}after')
                if before and before.isdigit():
                    style.append(f'margin-top: {int(before) / 20.0:.1f}pt;')
                if after and after.isdigit():
                    style.append(f'margin-bottom: {int(after) / 20.0:.1f}pt;')
                line_rule = spacing.get(f'{{{ns["w"]}}}lineRule')
                line = spacing.get(f'{{{ns["w"]}}}line')
                if line_rule == 'exact' and line and line.isdigit():
                    style.append(f'line-height: {int(line) / 240.0:.1f}pt;')
                elif line_rule == 'atLeast' and line and line.isdigit():
                    style.append(f'min-height: {int(line) / 240.0:.1f}pt;')
                elif line and line.isdigit():
                    style.append(f'line-height: {int(line) / 240.0:.1f};')
            ind = props.find('w:ind', ns)
            if ind is not None:
                left = ind.get(f'{{{ns["w"]}}}left')
                right = ind.get(f'{{{ns["w"]}}}right')
                first_line = ind.get(f'{{{ns["w"]}}}firstLine')
                hanging = ind.get(f'{{{ns["w"]}}}hanging')
                if left and left.isdigit():
                    style.append(f'margin-left: {int(left) / 20.0:.1f}pt;')
                if right and right.isdigit():
                    style.append(f'margin-right: {int(right) / 20.0:.1f}pt;')
                if first_line and first_line.isdigit():
                    style.append(f'text-indent: {int(first_line) / 20.0:.1f}pt;')
                if hanging and hanging.isdigit():
                    style.append(f'text-indent: -{int(hanging) / 20.0:.1f}pt; margin-left: {int(hanging) / 20.0:.1f}pt;')
            context_spacing = props.find('w:contextualSpacing', ns)
            if context_spacing is not None:
                val = context_spacing.get(f'{{{ns["w"]}}}val')
                if val == 'true':
                    style.append('margin-top: 0; margin-bottom: 0;')
            page_break_before = props.find('w:pageBreakBefore', ns)
            if page_break_before is not None:
                val = page_break_before.get(f'{{{ns["w"]}}}val')
                if val == 'true':
                    style.append('page-break-before: always;')
            borders = props.find('w:pBdr', ns)
            if borders is not None:
                for side in ['top', 'bottom', 'left', 'right']:
                    el = borders.find(f'w:{side}', ns)
                    if el is not None:
                        val = el.get(f'{{{ns["w"]}}}val')
                        sz = el.get(f'{{{ns["w"]}}}sz', '0')
                        color = el.get(f'{{{ns["w"]}}}color', '000000')
                        space = el.get(f'{{{ns["w"]}}}space', '0')
                        css_style = 'solid' if val == 'single' else 'double' if val == 'double' else 'none'
                        style.append(f'border-{side}: {int(sz)/8.0:.2f}pt {css_style} #{color};')
                        if int(space) > 0:
                            style.append(f'margin-{side}: {int(space)}pt;')
            shd = props.find('w:shd', ns)
            if shd is not None:
                fill = shd.get(f'{{{ns["w"]}}}fill')
                if fill and fill != 'auto' and fill != 'FFFFFF':
                    style.append(f'background-color: #{fill};')
            suppressAutoHyphens = props.find('w:suppressAutoHyphens', ns)
            if suppressAutoHyphens is not None:
                val = suppressAutoHyphens.get(f'{{{ns["w"]}}}val')
                if val == 'true':
                    style.append('hyphens: none;')
        return ' '.join(style)

    def _get_run_style(self, r, ns):
        props = r.find('w:rPr', ns)
        style = []
        if props is not None:
            if props.find('w:vanish', ns) is not None:
                style.append('display: none;')
            rfonts = props.find('w:rFonts', ns)
            if rfonts is not None:
                font_css = []
                east_asian = rfonts.get(f'{{{ns["w"]}}}eastAsia')
                ascii = rfonts.get(f'{{{ns["w"]}}}ascii')
                if ascii:
                    font_css.append(f'font-family: {ascii};')
                if east_asian:
                    font_css.append(f'font-variant-east-asian: {east_asian};')
                if font_css:
                    style.append(' '.join(font_css))
            if props.find('w:sz', ns) is not None:
                sz = props.find('w:sz', ns).get(f'{{{ns["w"]}}}val')
                if sz and sz.isdigit():
                    style.append(f'font-size: {int(sz) / 2.0:.1f}pt;')
            color = props.find('w:color', ns)
            if color is not None:
                val = color.get(f'{{{ns["w"]}}}val')
                if val and val != 'auto':
                    style.append(f'color: #{val};')
            if props.find('w:caps', ns) is not None:
                style.append('text-transform: uppercase;')
            if props.find('w:smallCaps', ns) is not None:
                style.append('font-variant: small-caps;')
            if props.find('w:strike', ns) is not None:
                style.append('text-decoration: line-through;')
            if props.find('w:dstrike', ns) is not None:
                style.append('text-decoration: line-through double;')
            if props.find('w:outline', ns) is not None:
                style.append('text-decoration: underline;')
            if props.find('w:shadow', ns) is not None:
                style.append('text-shadow: 1px 1px 2px #888888;')
            if props.find('w:emboss', ns) is not None:
                style.append('text-shadow: 1px 1px 0 #fff, 2px 2px 2px #888;')
            if props.find('w:imprint', ns) is not None:
                style.append('text-shadow: 1px 1px 0 #fff, -1px -1px 1px #888;')
            v_align = props.find('w:vAlign', ns)
            if v_align is not None:
                val = v_align.get(f'{{{ns["w"]}}}val')
                if val == 'top':
                    style.append('vertical-align: top;')
                elif val == 'center':
                    style.append('vertical-align: middle;')
                elif val == 'bottom':
                    style.append('vertical-align: bottom;')
            if props.find('w:b', ns) is not None:
                style.append('font-weight: bold;')
            if props.find('w:i', ns) is not None:
                style.append('font-style: italic;')
            u = props.find('w:u', ns)
            if u is not None:
                val = u.get(f'{{{ns["w"]}}}val')
                if val == 'single':
                    style.append('text-decoration: underline;')
                elif val == 'double':
                    style.append('text-decoration: underline double;')
        return ' '.join(style) 