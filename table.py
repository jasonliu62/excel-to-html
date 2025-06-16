import os
import zipfile
import xml.etree.ElementTree as ET
import re
from util import clean_text

class TableProcessor:
    def __init__(self):
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'tbl': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }

    def process_table(self, docx_path):
        """Extract and parse document.xml directly, build HTML table with dynamic structure and inline styles"""
        document_xml = os.path.join(docx_path, 'word', 'document.xml')
        tree = ET.parse(document_xml)
        root = tree.getroot()
        ns = self.namespaces

        html_tables = []
        for tbl in root.findall('.//w:tbl', ns):
            tbl_pr = tbl.find('w:tblPr', ns)
            total_width_twips = None
            tbl_cellmar = {}
            if tbl_pr is not None:
                tblw = tbl_pr.find('w:tblW', ns)
                if tblw is not None and tblw.get(f'{{{ns["w"]}}}type') == 'dxa':
                    w = tblw.get(f'{{{ns["w"]}}}w')
                    if w and w.isdigit():
                        total_width_twips = int(w)
                # Parse table-wide default cell margins
                tblcellmar = tbl_pr.find('w:tblCellMar', ns)
                if tblcellmar is not None:
                    for side in ['top', 'bottom', 'left', 'right']:
                        mar = tblcellmar.find(f'w:{side}', ns)
                        if mar is not None:
                            w_val = mar.get(f'{{{ns["w"]}}}w')
                            if w_val and w_val.isdigit():
                                tbl_cellmar[side] = int(w_val) / 20.0  # pt

            html_table = [
                '<table cellpadding="0" cellspacing="0" style="font: 10pt Times New Roman, Times, Serif; border-collapse: collapse; width: 100%">'
            ]
            for tr_idx, tr in enumerate(tbl.findall('w:tr', ns)):
                row_style = self._get_row_style(tr, ns)
                # Add row height if present
                tr_pr = tr.find('w:trPr', ns)
                if tr_pr is not None:
                    tr_height = tr_pr.find('w:trHeight', ns)
                    if tr_height is not None:
                        val = tr_height.get(f'{{{ns["w"]}}}val')
                        if val and val.isdigit():
                            pt = int(val) / 20.0
                            row_style += f' min-height: {pt:.1f}pt;'
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
                all_empty = all(cell.strip() == '' for cell in row_cells)
                tr_style = row_style
                if all_empty:
                    tr_style += ' min-height: 12pt;'
                    row_cells = ['&#160;' for _ in row_cells]
                html_table.append(f'<tr{f" style=\"{tr_style}\"" if tr_style else ""}>')
                for tc_idx, tc in enumerate(tcs):
                    cell_text = row_cells[tc_idx]
                    cell_style, colspan = self._get_cell_style(tc, ns, total_width_twips, tc_idx, cell_text, tbl_cellmar)
                    tag = 'td'
                    attrs = []
                    if colspan > 1:
                        attrs.append(f'colspan="{colspan}"')
                    if cell_style:
                        attrs.append(f'style="{cell_style}"')
                    attr_str = ' '.join(attrs)
                    html_table.append(f'<{tag} {attr_str}>{cell_text}</{tag}>')
                    if tc_idx == len(tcs) - 1:
                        props = tc.find('w:tcPr', ns)
                        if props is not None:
                            borders = props.find('w:tcBorders', ns)
                            if borders is not None:
                                bottom = borders.find('w:bottom', ns)
                                if bottom is not None and bottom.get(f'{{{ns["w"]}}}val') == 'double':
                                    last_cell_double_underline = True
                if last_cell_double_underline:
                    html_table.append('<td style="border-bottom: Black 2.5pt double;"></td>')
                html_table.append('</tr>')
            html_table.append('</table>')
            html_tables.append('\n'.join(html_table))
        return '\n\n'.join(html_tables)
    
    def process_table_element(self, tbl, ns):
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
        return '\n\n'.join(html_table)

    def _get_row_style(self, tr, ns):
        props = tr.find('w:trPr', ns)
        style = []
        if props is not None:
            shd = props.find('w:shd', ns)
            if shd is not None:
                fill = shd.get(f'{{{ns["w"]}}}fill')
                if fill and fill != 'auto' and fill != 'FFFFFF':
                    style.append(f'background-color: #{fill};')
        return ' '.join(style)

    def _get_cell_text(self, tc, ns):
        # Output plain text unless inline style is needed
        html = []
        for p in tc.findall('w:p', ns):
            para_text = []
            for child in list(p):
                tag = child.tag
                if tag == f'{{{ns["w"]}}}pPr':
                    continue
                if tag == f'{{{ns["w"]}}}r':
                    run_style = self._get_run_style(child, ns)
                    run_text = ''
                    for rchild in list(child):
                        rtag = rchild.tag
                        if rtag == f'{{{ns["w"]}}}t':
                            run_text += clean_text(rchild.text or '')
                        elif rtag == f'{{{ns["w"]}}}br':
                            run_text += '<br/>'
                    if run_text == '':
                        run_text = '&#160;'
                    # Only wrap in <span> if there is actual style
                    if run_style and run_text.strip() != '&#160;':
                        para_text.append(f'<span style="{run_style}">{run_text}</span>')
                    else:
                        para_text.append(run_text)
                elif tag == f'{{{ns["w"]}}}hyperlink':
                    for run in child.findall('w:r', ns):
                        run_style = self._get_run_style(run, ns)
                        run_text = ''
                        for rchild in list(run):
                            rtag = rchild.tag
                            if rtag == f'{{{ns["w"]}}}t':
                                run_text += clean_text(rchild.text or '')
                            elif rtag == f'{{{ns["w"]}}}br':
                                run_text += '<br/>'
                        if run_text == '':
                            run_text = '&#160;'
                        if run_style and run_text.strip() != '&#160;':
                            para_text.append(f'<span style="{run_style}">{run_text}</span>')
                        else:
                            para_text.append(run_text)
            html.append(''.join(para_text))
        text = ''.join(html)
        text = text.replace('–', '&#8211;').replace('—', '&#8212;')
        # Only output &#160; for empty
        if text.strip() == '' or text.strip() == '<span></span>' or text.strip() == '<span>&#160;</span>':
            text = '&#160;'
        return text

    def _get_cell_style(self, tc, ns, total_width_twips=None, tc_idx=0, cell_text=None, tbl_cellmar=None):
        props = tc.find('w:tcPr', ns)
        style_parts = []
        colspan = 1
        width_style = None
        bgcolor_style = None
        align_style = None
        border_style = []
        padding_style = None
        align = None
        if props is not None:
            gridspan = props.find('w:gridSpan', ns)
            if gridspan is not None:
                colspan = int(gridspan.get(f'{{{ns["w"]}}}val', '1'))
            tcw = props.find('w:tcW', ns)
            if tcw is not None:
                w = tcw.get(f'{{{ns["w"]}}}w')
                w_type = tcw.get(f'{{{ns["w"]}}}type')
                if w and w.isdigit() and int(w) > 0:
                    if w_type == 'pct':
                        width_percent = int(w) / 50.0
                        width_style = f'width: {width_percent:.0f}%;'
                    elif total_width_twips and total_width_twips > 0 and w_type == 'dxa':
                        w_twips = int(w)
                        width_percent = 100 * w_twips / total_width_twips
                        width_style = f'width: {width_percent:.0f}%;'
                    else:
                        width_pt = int(w) / 20
                        width_style = f'width: {width_pt:.1f}pt;'
            shd = props.find('w:shd', ns)
            if shd is not None:
                fill = shd.get(f'{{{ns["w"]}}}fill')
                if fill and fill != 'auto' and fill != 'FFFFFF':
                    bgcolor_style = f'background-color: #{fill};'
            jc = props.find('w:jc', ns)
            if jc is not None:
                align = jc.get(f'{{{ns["w"]}}}val', None)
                if align == 'center':
                    align_style = 'text-align: center;'
                elif align == 'left':
                    align_style = 'text-align: left;'
                elif align == 'right':
                    align_style = 'text-align: right;'
                elif align == 'both':
                    align_style = 'text-align: justify;'
                elif align == 'start':
                    align_style = 'text-align: start;'
                elif align == 'end':
                    align_style = 'text-align: end;'
            borders = props.find('w:tcBorders', ns)
            if borders is not None:
                for side in ['top', 'bottom', 'left', 'right']:
                    el = borders.find(f'w:{side}', ns)
                    if el is not None:
                        val = el.get(f'{{{ns["w"]}}}val')
                        sz = el.get(f'{{{ns["w"]}}}sz', '0')
                        color = el.get(f'{{{ns["w"]}}}color', '000000')
                        css_style = 'solid' if val == 'single' else 'double' if val == 'double' else 'none'
                        # Always use 2.5pt for double borders
                        if val == 'double':
                            border_style.append(f'border-{side}: 2.5pt double #{color};')
                        elif val and val != 'none':
                            border_style.append(f'border-{side}: {int(sz)/8.0:.2f}pt {css_style} #{color};')
            # Only include padding-bottom if present in DOCX
            tcMar = props.find('w:tcMar', ns)
            if tcMar is not None:
                bottom = tcMar.find('w:bottom', ns)
                if bottom is not None:
                    w_val = bottom.get(f'{{{ns["w"]}}}w')
                    if w_val and w_val.isdigit():
                        pt = int(w_val) / 20.0
                        if pt > 0:
                            padding_style = f'padding-bottom: {pt:.1f}pt;'
        if align is None:
            first_p = tc.find('w:p', ns)
            if first_p is not None:
                ppr = first_p.find('w:pPr', ns)
                if ppr is not None:
                    jc = ppr.find('w:jc', ns)
                    if jc is not None:
                        align = jc.get(f'{{{ns["w"]}}}val', None)
                        if align == 'center':
                            align_style = 'text-align: center;'
                        elif align == 'left':
                            align_style = 'text-align: left;'
                        elif align == 'right':
                            align_style = 'text-align: right;'
                        elif align == 'both':
                            align_style = 'text-align: justify;'
                        elif align == 'start':
                            align_style = 'text-align: start;'
                        elif align == 'end':
                            align_style = 'text-align: end;'
        # Default alignment based on cell content
        if align_style is None:
            if cell_text is not None:
                # Remove HTML tags and whitespace to check actual content
                clean_text = re.sub(r'<[^>]+>', '', cell_text).strip()
                if clean_text == '$' or clean_text == ')':
                    align_style = 'text-align: left;'
                elif tc_idx == 0:
                    align_style = 'text-align: left;'
                else:
                    align_style = 'text-align: right;'
        # Compose style in the order: width, background-color, text-align, border, padding-bottom
        if width_style:
            style_parts.append(width_style)
        if bgcolor_style:
            style_parts.append(bgcolor_style)
        if align_style:
            style_parts.append(align_style)
        if border_style:
            style_parts.extend(border_style)
        if padding_style:
            style_parts.append(padding_style)
        return ' '.join(style_parts), colspan

    def _get_paragraph_style(self, p, ns):
        props = p.find('w:pPr', ns)
        style = []
        if props is not None:
            pstyle = props.find('w:pStyle', ns)
            if pstyle is not None:
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
        props = r.find('w:rPr', ns) if hasattr(r, 'find') else r
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