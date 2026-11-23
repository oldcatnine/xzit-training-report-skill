"""Compare author content across formatting and Word saves."""
import hashlib
import re
from itertools import zip_longest
from docx import Document
from docx.oxml.ns import qn
from document_roles import CITE_GROUP, discover, citation_numbers, citation_label, paragraphs, role
from reference_format import styled_reference

MATH_NS = qn('m:oMath').split('}')[0]+'}'
MATH_OPERATORS = {'chr','begChr','endChr','sepChr','type','pos','degHide','subHide','supHide'}


def canonical_citation(match):
    try:
        return citation_label(citation_numbers(match[0]))
    except ValueError:
        return match[0]


def snapshot(path, mapping=None):
    doc = Document(path)
    body, refs, _ = discover(doc, mapping)
    citation_paragraphs = {p._p for p in body}
    reference_paragraphs = {p._p for p in refs.values()}
    result = []
    for position, p in enumerate(paragraphs(doc), 1):
        kind = role(p, mapping) or ''
        if re.fullmatch(r'TOC[1-9]?', kind) or kind == 'TOCField':
            continue
        # Paragraph.text includes hyperlinks but omits native math and markup.
        # Read authored XML text including text boxes and tracked change text.
        texts = []
        for node in p._p.xpath('.//w:t | .//w:delText | .//w:tab | .//w:br | .//w:cr'):
            if next(node.iterancestors(qn('w:p')), None) is not p._p:
                continue
            if any(a.tag in (qn('w:pPr'),qn('w:rPr')) for a in node.iterancestors()):
                continue
            if node.tag == qn('w:br') and node.get(qn('w:type')) in ('page','column'):
                continue
            if node.tag in (qn('w:t'),qn('w:delText')):
                texts.append(node.text or '')
            else:
                texts.append('\t' if node.tag == qn('w:tab') else '\n')
        content = ''.join(texts)
        if p._p in citation_paragraphs:
            content = CITE_GROUP.sub(canonical_citation, content)
        if p._p in reference_paragraphs:
            content = styled_reference(content)
        images = []
        for node in p._p.xpath('.//a:blip'):
            for attr in ('r:embed','r:link'):
                rid = node.get(qn(attr))
                if rid:
                    rel = doc.part.rels.get(rid)
                    images.append(rel.target_ref if rel and rel.is_external else
                        hashlib.sha256(rel.target_part.blob).hexdigest() if rel else 'missing:'+rid)
        # VML images use r:id. Compare resolved content rather than mutable rIds.
        for node in p._p.iter(qn('w:pict')):
            for image in node.iter('{urn:schemas-microsoft-com:vml}imagedata'):
                rid = image.get(qn('r:id'))
                rel = doc.part.rels.get(rid)
                images.append(rel.target_ref if rel and rel.is_external else
                    hashlib.sha256(rel.target_part.blob).hexdigest() if rel else 'missing:'+str(rid))
        math = []
        for formula in p._p.xpath('.//m:oMath'):
            parts = []
            for node in formula.iter():
                local = node.tag.split('}')[-1]
                if not node.tag.startswith(MATH_NS) or local in ('oMath','r'):
                    continue
                # Preserve visible operator choices; ignore Word's default
                # property materialization and math run styling.
                if local in MATH_OPERATORS and node.get(qn('m:val')) is not None:
                    parts.append((local, node.get(qn('m:val'))))
                    continue
                if any(a.tag.split('}')[-1].endswith('Pr') for a in node.iterancestors() if a is not p._p):
                    continue
                if local == 't' and parts and parts[-1][0] == 't':
                    parts[-1] = ('t',parts[-1][1] + (node.text or ''))
                elif local == 't':
                    parts.append(('t',node.text or ''))
                elif not local.endswith('Pr'):
                    parts.append((local, tuple(a.tag.split('}')[-1] for a in node.iterancestors()
                        if a is not formula and a.tag.startswith(MATH_NS))))
            math.append(parts)
        coords = tuple((a.tag.split('}')[-1], len(list(a.itersiblings(preceding=True,tag=a.tag)))+1)
                       for a in reversed(list(p._p.iterancestors())) if a.tag in {qn('w:tbl'),qn('w:tr'),qn('w:tc')})
        if content or images or math or coords or kind in ('Figure','Equation'):
            result.append((position, (content,images,math,coords)))
    return result


def compare(path, before, mapping=None):
    for old, new in zip_longest(before, snapshot(path, mapping)):
        if old is None or new is None or old[1] != new[1]:
            labels = ('text','images','equations','table position')
            changed = [label for label,a,b in zip(labels,old[1],new[1]) if a!=b] if old and new else ['paragraph count']
            return {'status':'fail','before_paragraph':old[0] if old else None,
                    'after_paragraph':new[0] if new else None,'changed':changed,
                    'before_text':old[1][0][:120] if old else None,'after_text':new[1][0][:120] if new else None}
    return {'status':'pass','compared_paragraphs':len(before)}
