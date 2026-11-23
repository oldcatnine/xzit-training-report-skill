"""Shared body/bibliography discovery, including conservative legacy style support."""
import re
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

CITE = re.compile(r'\[(\d+(?:\s*[-,，]\s*\d+)*)\]')
CITE_GROUP = re.compile(CITE.pattern + r'(?:[ \t]*' + CITE.pattern + r')*')


def role(p, mapping=None):
    name = p.style.name
    if name in (mapping or {}):
        return mapping[name]
    if name.startswith('XZIT'):
        return name[4:]
    match = re.fullmatch(r'(?:Heading|标题)\s*([1-4])', name, re.I)
    if match:
        return 'H' + match[1]
    if re.fullmatch(r'toc\s+[1-3]', name, re.I):
        return 'TOC'
    return None


def paragraphs(doc):
    for element in doc.element.body:
        for pe in ([element] if element.tag == qn('w:p') else element.xpath('.//w:p')):
            yield Paragraph(pe, doc)


def discover(doc, mapping=None):
    """Return body paragraphs, numbered bibliography, and explicit discovery errors.

    A plain-text bibliography heading is accepted, but body start must be a real
    heading (built-in, XZIT or mapped). Unrecognized documents fail visibly.
    """
    body, refs, errors = [], {}, []
    main = in_refs = found_heading = False
    for p in paragraphs(doc):
        text = p.text.strip()
        kind = role(p, mapping)
        if kind == 'TOC' or (kind or '').startswith('TOC'):
            continue
        if re.fullmatch(r'(?:\d+[.、\s]*)?(?:参考文献|References|Bibliography)', text, re.I):
            in_refs = found_heading = True
            main = False
            continue
        if in_refs:
            if kind == 'H1':
                in_refs = False
                continue
            match = re.match(r'^\[(\d+)\]\s*(.+)', text)
            if match:
                n = int(match[1])
                if n in refs:
                    errors.append('Duplicate bibliography number ' + str(n))
                refs[n] = p
            continue
        if kind == 'H1' and not found_heading and re.match(r'^(?:第[一二三四五六七八九十\d]+章|\d+(?:[.、\s]|$))', text):
            main = True
        if main:
            body.append(p)
    if not body:
        errors.append('Body start not recognized; map the first numbered chapter style to H1')
    if not found_heading or not refs:
        errors.append('Numbered bibliography under a References/参考文献 heading not recognized')
    return body, refs, errors


def numbers(label):
    result = []
    for part in re.split(r'[,，]', label.strip('[]')):
        bounds = part.split('-')
        lo, hi = (int(bounds[0]), int(bounds[-1]))
        if lo < 1 or hi < lo or hi - lo > 1000:
            raise ValueError('Invalid citation range: ' + label)
        result.extend(range(lo, hi + 1))
    return result


def citation_numbers(text):
    return [number for match in CITE.finditer(text) for number in numbers(match[0])]


def citation_label(values):
    values = sorted(set(values))
    parts = []
    i = 0
    while i < len(values):
        end = i
        while end+1 < len(values) and values[end+1] == values[end]+1:
            end += 1
        if end-i >= 2:
            parts.append(f'{values[i]}-{values[end]}')
        else:
            parts.extend(str(n) for n in values[i:end+1])
        i = end+1
    return '[' + ','.join(parts) + ']'


def suggest_roles(doc, mapping=None):
    """Suggest per-paragraph roles and only offer homogeneous style mappings."""
    groups = {}
    for index, p in enumerate(paragraphs(doc), 1):
        text = p.text.strip()
        if not text or role(p, mapping):
            continue
        name = p.style.name
        suggestion, reason = None, 'Insufficient structural evidence'
        if re.match(r'^续表\s*\d+[-–.]\d+', text):
            suggestion, reason = 'Continuation', 'Numbered continuation label'
        elif re.match(r'^[图表]\s*\d+[-–.]\d+(?:\s|[\u3400-\u9fff])', text):
            suggestion, reason = 'Caption', 'Numbered figure/table caption'
        elif p._p.xpath('.//m:oMath'):
            if re.fullmatch(r'\(\d+-\d+\)', text):
                suggestion, reason = 'Equation', 'Native math object and equation number'
        elif p._p.xpath('ancestor::w:tc'):
            suggestion, reason = 'Table', 'Inside a table cell; exclude cover metadata before mapping'
        else:
            props = [p._p.pPr, p.style.element.find(qn('w:pPr'))]
            levels = [level.get(qn('w:val')) for pr in props if pr is not None
                      for level in pr.findall(qn('w:outlineLvl'))]
            if levels and levels[0] in ('0','1','2','3'):
                suggestion, reason = f'H{int(levels[0])+1}', 'Explicit Word outline level'
            elif len(text) <= 55 and (match := re.match(r'^(\d+(?:\.\d+){1,3})\s+\S',text)):
                suggestion, reason = f'H{match[1].count(".")+1}', 'Numbered outline pattern'
        groups.setdefault(name, []).append({'paragraph':index,'preview':text[:100],
            'role':suggestion,'reason':reason})
    proposed = {}
    for name, items in groups.items():
        kinds = {item['role'] for item in items}
        if len(kinds) == 1 and kinds.isdisjoint({None,'Table'}):
            proposed[name] = items[0]['role']
    return {'proposed_role_map':proposed,'unmapped_styles':groups,'status':'pending' if groups else 'not_applicable'}
