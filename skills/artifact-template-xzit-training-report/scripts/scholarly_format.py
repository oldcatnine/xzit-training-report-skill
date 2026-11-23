"""Compact linked citations, bibliography presentation and three-line tables.

Edits only plain text runs; fields, equations, images and existing links remain.
"""
import argparse,json,re
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from report_format import child
from document_roles import discover, role, CITE_GROUP, citation_numbers, citation_label
from reference_format import DOI_SPAN, DOI_URL, author_spans

PREFIX='xzit_ref_'
TABLE_OUTER_EIGHTHS = 12
TABLE_HEADER_EIGHTHS = 4

def text_run(original,text,superscript=False):
    r=deepcopy(original)
    for e in list(r):
        if e.tag!=qn('w:rPr'):r.remove(e)
    t=OxmlElement('w:t');t.set(qn('xml:space'),'preserve');t.text=text;r.append(t)
    if superscript:
        rp=r.get_or_add_rPr()
        child(rp,'w:vertAlign',val='superscript')
        child(rp,'w:color',val='000000');child(rp,'w:u',val='none')
    return r

def hyperlink(original,text,anchor=None,superscript=False):
    link=OxmlElement('w:hyperlink');link.set(qn('w:history'),'1')
    if anchor:
        link.set(qn('w:anchor'),anchor)
        link.set(qn('w:tooltip'),'跳转到参考文献'+text)
    r=text_run(original,text,superscript)
    child(r.get_or_add_rPr(),'w:color',val='000000');child(r.get_or_add_rPr(),'w:u',val='none')
    link.append(r);return link

def replace_plain_spans(p,regex,make_nodes):
    """Match across adjacent plain runs, retaining formatting outside matches."""
    groups=[];group=[]
    for e in p._p:
        plain=e.tag==qn('w:r') and all(x.tag in (qn('w:rPr'),qn('w:t')) for x in e)
        # Word spellchecking inserts proofErr markers inside DOI strings.
        # Treat those zero-width hints as transparent; they may be regenerated.
        if plain or e.tag==qn('w:proofErr'):group.append(e)
        else:
            if group:groups.append(group);group=[]
    if group:groups.append(group)
    replaced=0
    for group in reversed(groups):
        segments=[];at=0
        for e in group:
            text=''.join(x.text or '' for x in e.findall(qn('w:t')))
            segments.append((at,at+len(text),e,text));at+=len(text)
        content=''.join(x[3] for x in segments);matches=list(regex.finditer(content))
        if not matches:continue
        nodes=[]
        def plain_slice(start,end):
            for lo,hi,e,text in segments:
                a,b=max(start,lo),min(end,hi)
                if a<b:nodes.append(text_run(e,text[a-lo:b-lo]))
        cursor=0
        for match in matches:
            plain_slice(cursor,match.start())
            original=next(e for lo,hi,e,text in segments if lo<=match.start()<hi)
            nodes.extend(make_nodes(original,match));cursor=match.end();replaced+=1
        plain_slice(cursor,len(content))
        parent=group[0].getparent();index=parent.index(group[0])
        for e in group:parent.remove(e)
        for e in nodes:parent.insert(index,e);index+=1
    return replaced


def unwrap(link):
    parent = link.getparent()
    index = parent.index(link)
    for run in list(link):
        parent.insert(index,run)
        index += 1
    parent.remove(link)


def unwrap_managed_fields(p, matches):
    """Expose cached text for Word HYPERLINK fields owned by these rules."""
    for field in list(p._p.xpath('./w:fldSimple')):
        if matches(field.get(qn('w:instr'),'')):
            unwrap(field)
    begin = separator = None
    depth = 0
    instruction = []
    nested = False
    for i, node in enumerate(children := list(p._p)):
        kind = node.find(qn('w:fldChar')) if node.tag == qn('w:r') else None
        marker = kind.get(qn('w:fldCharType')) if kind is not None else None
        if marker == 'begin':
            if depth == 0:
                begin, separator, instruction, nested = i, None, [], False
            else:
                nested = True
            depth += 1
        elif marker == 'separate' and depth == 1:
            separator = i
        elif marker == 'end' and depth:
            depth -= 1
            if depth == 0 and separator is not None and not nested and matches(''.join(instruction)):
                for part in children[begin:separator+1] + [node]:
                    p._p.remove(part)
        elif depth == 1 and separator is None:
            instruction.extend(part.text or '' for part in node.iter(qn('w:instrText')))


def is_doi_field(instruction):
    match = re.match(r'^\s*HYPERLINK\s+"?([^"\s]+)',instruction,re.I)
    return bool(match and DOI_URL.match(match[1]))


def is_citation_field(instruction):
    return bool(re.match(r'^\s*HYPERLINK\s+\\l\s+"?'+PREFIX+r'\d+(?=["\s]|$)',instruction,re.I))


def adjust_author_case(p):
    nodes = p._p.xpath('./w:r/w:t | ./w:hyperlink/w:r/w:t')
    plain = ''.join(n.text or '' for n in nodes)
    if plain != p.text:
        return
    spans = author_spans(plain)
    cursor = 0
    for node in nodes:
        value = node.text or ''
        end = cursor+len(value)
        for start,stop,replacement in reversed(spans):
            lo,hi = max(start,cursor),min(stop,end)
            if lo < hi:
                value = value[:lo-cursor]+replacement[lo-start:hi-start]+value[hi-cursor:]
        node.text = value
        cursor = end


def linked_citation(original, values):
    label = citation_label(values)
    if len(set(values)) == 1:
        return [hyperlink(original,label,anchor=PREFIX+str(values[0]),superscript=True)]
    parts = re.split(r'(\d+)',label)
    return [hyperlink(original,part,anchor=PREFIX+part,superscript=True) if part.isdecimal()
            else text_run(original,part,True) for part in parts if part]

def format_document(doc,role_map=None):
    body,refs,errors=discover(doc,role_map)
    if errors:raise ValueError('; '.join(errors))
    body_elements={p._p for p in body}
    ref_elements={p._p for p in refs.values()}
    ids=[int(x) for x in doc.element.xpath('.//w:bookmarkStart/@w:id') if str(x).isdigit()]
    next_id=max(ids,default=-1)+1
    for n,p in refs.items():
        name=PREFIX+str(n)
        existing=doc.element.xpath('.//w:bookmarkStart[@w:name="'+name+'"]')
        if not existing:
            start=OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(next_id));start.set(qn('w:name'),name)
            end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(next_id));next_id+=1
            p._p.insert(1 if p._p.pPr is not None else 0,start);p._p.append(end)
        elif len(existing)!=1 or existing[0].getparent() is not p._p:raise ValueError('Conflicting bibliography bookmark '+name)
    citations=0;removed_dois=0
    unmapped_doi_links=set()
    for e in doc.element.body:
        ps=[e] if e.tag==qn('w:p') else e.xpath('.//w:p')
        for pe in ps:
            p=Paragraph(pe,doc)
            if pe in ref_elements:
                unwrap_managed_fields(p,is_doi_field)
                for link in list(p._p.xpath('./w:hyperlink[@r:id]')):
                    rid = link.get(qn('r:id'))
                    target = p.part.rels.get(rid)
                    if target and DOI_URL.match(target.target_ref):
                        unwrap(link)
                        unmapped_doi_links.add(rid)
                removed_dois += replace_plain_spans(p,DOI_SPAN,lambda *_: [])
                adjust_author_case(p)
            elif pe in body_elements:
                unwrap_managed_fields(p,is_citation_field)
                for link in list(p._p.xpath('./w:hyperlink')):
                    if link.get(qn('w:anchor'),'').startswith(PREFIX):
                        unwrap(link)
                def citation_nodes(original,m):
                    values=citation_numbers(m[0])
                    if any(n not in refs for n in values):raise ValueError('Citation has no bibliography entry: '+m[0])
                    return linked_citation(original,values)
                citations+=replace_plain_spans(p,CITE_GROUP,citation_nodes)
    remaining_rel_ids=set(doc.element.xpath('.//@r:id'))
    for rid in unmapped_doi_links - remaining_rel_ids:
        doc.part.drop_rel(rid)
    table_count=0
    for table in doc.tables:
        if not any(role(p,role_map)=='Table' for row in table.rows for cell in row.cells for p in cell.paragraphs):continue
        table_count+=1
        borders=child(table._tbl.tblPr,'w:tblBorders')
        for side in ('top','bottom','left','right','insideH','insideV'):
            child(borders,'w:'+side,val='single' if side in ('top','bottom') else 'nil',sz=TABLE_OUTER_EIGHTHS,color='000000')
        for i,row in enumerate(table.rows):
            for cell in row.cells:
                cb=child(cell._tc.get_or_add_tcPr(),'w:tcBorders')
                for side in ('top','bottom','left','right'):child(cb,'w:'+side,val='nil')
                if i==0:
                    child(cb,'w:top',val='single',sz=TABLE_OUTER_EIGHTHS,color='000000')
                    child(cb,'w:bottom',val='single',sz=TABLE_HEADER_EIGHTHS,color='000000')
                if i==len(table.rows)-1:child(cb,'w:bottom',val='single',sz=TABLE_OUTER_EIGHTHS,color='000000')
    return {'new_citation_spans':citations,'removed_doi_spans':removed_dois,'bibliography_entries':len(refs),
            'three_line_tables':table_count,'outer_line_width_pt':TABLE_OUTER_EIGHTHS/8,
            'header_line_width_pt':TABLE_HEADER_EIGHTHS/8}

def audit_links(doc,role_map=None):
    failures=[];links=doc.element.xpath('.//w:hyperlink[starts-with(@w:anchor,"'+PREFIX+'")]')
    bookmarks=doc.element.xpath('.//w:bookmarkStart/@w:name')
    if not links:failures.append('No linked superscript citations found')
    for link in links:
        if link.get(qn('w:anchor')) not in bookmarks:failures.append('Broken citation bookmark')
        elif bookmarks.count(link.get(qn('w:anchor')))!=1:failures.append('Duplicate citation bookmark')
        label=''.join(link.xpath('.//w:t/text()'))
        if not re.fullmatch(r'(?:\[\d+\]|\d+)',label) or link.get(qn('w:anchor'))!=PREFIX+label.strip('[]'):failures.append('Citation text and target disagree')
        for r in link.xpath('./w:r'):
            if not r.xpath('./w:rPr/w:vertAlign[@w:val="superscript"]'):failures.append('Citation not superscript')
    body,refs,errors=discover(doc,role_map)
    failures.extend(errors)
    for p in body:
        offsets = {}
        position = 0
        for item in p.iter_inner_content():
            end = position+len(item.text)
            anchor = item._element.get(qn('w:anchor'),'')
            if anchor.startswith(PREFIX):
                offsets.update((index,anchor) for index in range(position,end))
            position = end
        for match in CITE_GROUP.finditer(p.text):
            try:
                expected = citation_label(citation_numbers(match[0]))
            except ValueError:
                failures.append('Invalid citation: '+match[0])
                continue
            if match[0] != expected:
                failures.append('Citation should use compact notation: '+expected)
            for digits in re.finditer(r'\d+',match[0]):
                if offsets.get(match.start()+digits.start()) != PREFIX+digits[0]:
                    failures.append('Unlinked citation number: '+digits[0])
    for p in refs.values():
        targets = [p.part.rels.get(e.get(qn('r:id'))) for e in p._p.xpath('./w:hyperlink[@r:id]')]
        fields = p._p.xpath('./w:fldSimple/@w:instr | ./w:r/w:instrText/text()')
        if DOI_SPAN.search(p.text) or any(target and DOI_URL.match(target.target_ref) for target in targets) or any(is_doi_field(field) for field in fields):
            failures.append('Remove DOI from bibliography: '+p.text[:70])
        if author_spans(p.text):
            failures.append('Use capitalization for English author names: '+p.text[:70])
    return {'status':'fail' if failures else 'pass','linked_citation_count':len(links),'failures':failures}

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('input',type=Path);ap.add_argument('output',type=Path);ap.add_argument('--role-map',type=Path);a=ap.parse_args()
    if a.output.exists():ap.error('Use a new output path')
    mapping=json.loads(a.role_map.read_text(encoding='utf-8-sig')) if a.role_map else None
    d=Document(a.input);result=format_document(d,mapping);d.save(a.output)
    result['link_audit']=audit_links(Document(a.output),mapping);print(json.dumps(result,ensure_ascii=False,indent=2))
    raise SystemExit(result['link_audit']['status']=='fail')
