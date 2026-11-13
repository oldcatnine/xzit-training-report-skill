"""Superscript linked citations, DOI links and reference-matched table borders.

Edits only plain text runs; fields, equations, images and existing links remain.
"""
import argparse,json,re
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.text.paragraph import Paragraph
from report_format import child

CITE=re.compile(r'\[(\d+(?:\s*[-,，]\s*\d+)*)\]')
PREFIX='xzit_ref_'
TABLE_LINE_EIGHTHS=12  # 1.5pt: observed source sample table cell borders.

def numbers(label):
    result=[]
    for part in re.split(r'[,，]',label.strip('[]')):
        bounds=part.split('-')
        if len(bounds)==1:result.append(int(bounds[0]))
        else:
            lo,hi=map(int,bounds)
            if hi<lo or hi-lo>1000:raise ValueError('Invalid citation range')
            result.extend(range(lo,hi+1))
    return result

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

def hyperlink(original,text,anchor=None,relationship=None,superscript=False):
    link=OxmlElement('w:hyperlink');link.set(qn('w:history'),'1')
    if anchor:
        link.set(qn('w:anchor'),anchor)
        link.set(qn('w:tooltip'),'跳转到参考文献'+text)
    if relationship:link.set(qn('r:id'),relationship)
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

def format_document(doc):
    refs={}
    for p in doc.paragraphs:
        if p.style.name=='XZITReference' and (m:=re.match(r'^\[(\d+)\]',p.text.strip())):refs[int(m[1])]=p
    if not refs:raise ValueError('No mapped bibliography entries')
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
    citations=0;doi_links=0;main=False;in_refs=False
    for e in doc.element.body:
        ps=[e] if e.tag==qn('w:p') else e.xpath('.//w:p')
        for pe in ps:
            p=Paragraph(pe,doc);text=p.text.strip()
            if p.style.name=='XZITH1' and re.match(r'^1\s',text):main=True
            if p.style.name=='XZITH1' and text=='参考文献':in_refs=True
            elif in_refs and p.style.name=='XZITH1':in_refs=False;main=False
            if in_refs and p.style.name=='XZITReference':
                # Repair a previously truncated DOI link without changing text.
                expected=re.search(r'DOI\s*:\s*(10\.\d{4,9}/\S+)',p.text,re.I)
                expected_url='https://doi.org/'+expected[1].rstrip('.,;，。') if expected else None
                for link in list(p._p.xpath('./w:hyperlink[@r:id]')):
                    rid=link.get(qn('r:id'));url=p.part.rels[rid].target_ref
                    if url.startswith('https://doi.org/') and expected_url and url!=expected_url:
                        at=p._p.index(link)
                        for r in list(link):p._p.insert(at,r);at+=1
                        p._p.remove(link)
                def doi_node(original,m):
                    doi=m[0].rstrip('.,;，。')
                    rid=doc.part.relate_to('https://doi.org/'+doi,RT.HYPERLINK,is_external=True)
                    nodes=[hyperlink(original,doi,relationship=rid)]
                    if len(doi)<len(m[0]):nodes.append(text_run(original,m[0][len(doi):]))
                    return nodes
                doi_links+=replace_plain_spans(p,re.compile(r'10\.\d{4,9}/[^\s]+'),doi_node)
            elif main and not p.style.name.lower().startswith('toc '):
                def citation_nodes(original,m):
                    values=numbers(m[0])
                    if any(n not in refs for n in values):raise ValueError('Citation has no bibliography entry: '+m[0])
                    return [hyperlink(original,f'[{n}]',anchor=PREFIX+str(n),superscript=True) for n in values]
                citations+=replace_plain_spans(p,CITE,citation_nodes)
                # Preserve and restyle links inserted by an earlier invocation.
                for link in p._p.xpath('./w:hyperlink'):
                    if link.get(qn('w:anchor'),'').startswith(PREFIX):
                        for r in link.xpath('./w:r'):child(r.get_or_add_rPr(),'w:vertAlign',val='superscript')
    table_count=0
    for table in doc.tables:
        if not any(p.style.name=='XZITTable' for row in table.rows for cell in row.cells for p in cell.paragraphs):continue
        table_count+=1
        borders=child(table._tbl.tblPr,'w:tblBorders')
        for side in ('top','bottom','left','right','insideH','insideV'):
            child(borders,'w:'+side,val='single' if side in ('top','bottom') else 'nil',sz=TABLE_LINE_EIGHTHS,color='000000')
        for i,row in enumerate(table.rows):
            for cell in row.cells:
                cb=child(cell._tc.get_or_add_tcPr(),'w:tcBorders')
                for side in ('top','bottom','left','right'):child(cb,'w:'+side,val='nil')
                if i==0:
                    child(cb,'w:top',val='single',sz=TABLE_LINE_EIGHTHS,color='000000')
                    child(cb,'w:bottom',val='single',sz=TABLE_LINE_EIGHTHS,color='000000')
                if i==len(table.rows)-1:child(cb,'w:bottom',val='single',sz=TABLE_LINE_EIGHTHS,color='000000')
    return {'new_citation_spans':citations,'new_doi_links':doi_links,'bibliography_entries':len(refs),'three_line_tables':table_count,'line_width_pt':TABLE_LINE_EIGHTHS/8}

def audit_links(doc):
    failures=[];links=doc.element.xpath('.//w:hyperlink[starts-with(@w:anchor,"'+PREFIX+'")]')
    bookmarks=set(doc.element.xpath('.//w:bookmarkStart/@w:name'))
    if not links:failures.append('No linked superscript citations found')
    for link in links:
        if link.get(qn('w:anchor')) not in bookmarks:failures.append('Broken citation bookmark')
        for r in link.xpath('./w:r'):
            if not r.xpath('./w:rPr/w:vertAlign[@w:val="superscript"]'):failures.append('Citation not superscript')
    doi_count=0
    for p in doc.paragraphs:
        if p.style.name!='XZITReference':continue
        m=re.search(r'DOI\s*:\s*(10\.\d{4,9}/\S+)',p.text,re.I)
        if not m:continue
        expected='https://doi.org/'+m[1].rstrip('.,;，。')
        targets=[p.part.rels[e.get(qn('r:id'))].target_ref for e in p._p.xpath('./w:hyperlink[@r:id]')]
        if expected not in targets:failures.append('DOI hyperlink does not cover complete DOI: '+m[1])
        else:doi_count+=1
    return {'status':'fail' if failures else 'pass','linked_citation_count':len(links),'complete_doi_link_count':doi_count,'failures':failures}

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('input',type=Path);ap.add_argument('output',type=Path);a=ap.parse_args()
    if a.output.exists():ap.error('Use a new output path')
    d=Document(a.input);result=format_document(d);d.save(a.output)
    result['link_audit']=audit_links(Document(a.output));print(json.dumps(result,ensure_ascii=False,indent=2))
