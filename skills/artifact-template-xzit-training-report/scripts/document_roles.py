"""Shared body/bibliography discovery, including conservative legacy style support."""
import re
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

CITE = re.compile(r'\[(\d+(?:\s*[-,，]\s*\d+)*)\]')


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
