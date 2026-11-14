"""Small native OMML builder. No evaluation of formula text or external conversion."""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_TAB_ALIGNMENT
from lxml import etree
from report_format import paragraph


def math_node(value, depth=0):
    if depth > 32:
        raise ValueError('Equation nesting exceeds 32')
    if isinstance(value, str) and value:
        r = OxmlElement('m:r')
        t = OxmlElement('m:t'); t.text = value; r.append(t)
        return [r]
    if isinstance(value, list) and value:
        return [node for item in value for node in math_node(item, depth + 1)]
    kinds = {'fraction': ('m:f', ('num', 'den')),
             'subscript': ('m:sSub', ('e', 'sub')),
             'superscript': ('m:sSup', ('e', 'sup')),
             'root': ('m:rad', ('deg', 'e'))}
    if not isinstance(value, dict) or value.get('kind') not in kinds:
        raise ValueError('Use equation strings, lists, fraction/subscript/superscript/root nodes')
    tag, slots = kinds[value['kind']]
    node = OxmlElement(tag)
    for slot in slots:
        child = OxmlElement('m:' + slot)
        for part in math_node(value[slot], depth + 1): child.append(part)
        node.append(child)
    return [node]


def equation_element(block):
    if ('math' in block) == ('omml' in block):
        raise ValueError('Equation requires exactly one of math or omml')
    if 'omml' in block:
        raw = block['omml']
        if not isinstance(raw, str) or '<!DOCTYPE' in raw.upper() or '<!ENTITY' in raw.upper():
            raise ValueError('Equation XML must not contain DTDs or entities')
        root = etree.fromstring(raw.encode('utf-8'), etree.XMLParser(resolve_entities=False, no_network=True))
        if root.tag != qn('m:oMath') or any(e.tag.startswith('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}') and e.tag != qn('w:rPr') and e.getparent().tag != qn('w:rPr') for e in root.iter()):
            raise ValueError('Expected a standalone m:oMath object without Word document content')
        for element in root.iter():
            if element.tag.startswith('{') and not element.tag.startswith(('{' + qn('m:oMath')[1:].split('}')[0] + '}', '{' + qn('w:rPr')[1:].split('}')[0] + '}')):
                raise ValueError('Unsupported namespace in OMML')
        return root
    root = OxmlElement('m:oMath')
    for part in math_node(block['math']): root.append(part)
    return root


def add_equation(doc, block):
    p = paragraph(doc, '', 'Equation')
    width = doc.sections[-1].page_width - doc.sections[-1].left_margin - doc.sections[-1].right_margin
    p.paragraph_format.tab_stops.add_tab_stop(int(width / 2), WD_TAB_ALIGNMENT.CENTER)
    p.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
    p.add_run('\t'); p._p.append(equation_element(block)); p.add_run('\t(' + block['number'] + ')')
    p.paragraph_format.keep_with_next = bool(block.get('definitions'))
    for definition in block.get('definitions', []):
        symbol, text = definition['symbol'], definition['text']
        d = paragraph(doc, symbol + '——\t' + text, 'Definition')
        # Explicit hanging indent aligns wrapped definitions after the dash.
        from docx.shared import Pt
        indent = (len(symbol) + 3) * 12
        d.paragraph_format.left_indent = Pt(indent)
        d.paragraph_format.first_line_indent = Pt(-indent)
        d.paragraph_format.tab_stops.add_tab_stop(Pt(indent))
    return p
