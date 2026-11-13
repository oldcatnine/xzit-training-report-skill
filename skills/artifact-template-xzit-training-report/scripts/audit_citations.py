"""Audit numbered citations and bibliography in body order, excluding TOC cache."""
import argparse,json,re
from collections import Counter
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph

def audit(path,minimum=0):
    doc=Document(path);entries={};cited=[];locations={};fail=[];main=False;in_refs=False
    for pos,e in enumerate(doc.element.body,1):
        ps=[e] if e.tag.endswith('}p') else e.xpath('.//w:p')
        for item in ps:
            p=Paragraph(item,doc);t=p.text.strip()
            if p.style.name=='XZITH1' and re.match(r'^1\s',t):main=True
            if t=='参考文献' and p.style.name=='XZITH1':in_refs=True;continue
            if in_refs and p.style.name=='XZITH1':in_refs=False;main=False
            if in_refs:
                m=re.match(r'^\[(\d+)\]\s*(.+)',t)
                if m:
                    n=int(m[1])
                    if n in entries:fail.append('Duplicate bibliography number '+str(n))
                    entries[n]=m[2]
                continue
            if not main or p.style.name.lower().startswith('toc '):continue
            for m in re.finditer(r'\[(\d+(?:\s*[-,，]\s*\d+)*)\]',t):
                for part in re.split(r'[,，]',m[1]):
                    ns=part.split('-');values=range(int(ns[0]),int(ns[1])+1) if len(ns)==2 else [int(ns[0])]
                    for n in values:cited.append(n);locations.setdefault(n,[]).append({'body_element':pos,'context':t[:90]})
    nums=set(entries);used=set(cited)
    if len(nums)<minimum:fail.append(f'References {len(nums)} below task minimum {minimum}')
    if nums!=set(range(1,len(nums)+1)):fail.append('Bibliography is not continuous from 1')
    if used-nums:fail.append('Missing bibliography for '+str(sorted(used-nums)))
    if nums-used:fail.append('Uncited bibliography '+str(sorted(nums-used)))
    first=list(dict.fromkeys(cited))
    if first!=sorted(first):fail.append('First citation order differs from numbering')
    dois=[re.search(r'DOI\s*:\s*(\S+)',v,re.I) for v in entries.values()]
    values=[m[1].rstrip('.').lower() for m in dois if m]
    if len(values)!=len(set(values)):fail.append('Duplicate DOI entries')
    return {'status':'fail' if fail else 'pass','task_minimum':minimum,'reference_count':len(nums),'distinct_cited_count':len(used),'first_occurrence_order':first,'citation_counts':dict(Counter(cited)),'failures':fail,'locations':locations,'scope':'Checks numbering and correspondence only; factual support and source metadata still require review.'}

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('docx',type=Path);ap.add_argument('--min-references',type=int,default=0);ap.add_argument('--output',type=Path);a=ap.parse_args()
    if a.min_references<0:ap.error('Minimum cannot be negative')
    if a.output and a.output.exists():ap.error('Refusing to overwrite output')
    result=audit(a.docx,a.min_references);payload=json.dumps(result,ensure_ascii=False,indent=2)
    if a.output:a.output.write_text(payload,encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k!='locations'},ensure_ascii=False,indent=2));raise SystemExit(bool(result['failures']))
