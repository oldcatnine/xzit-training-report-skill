"""Apply explicitly mapped paragraph roles on a COPY, preserving semantic objects."""
import argparse, json, re
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from document_roles import paragraphs, role, discover
from report_format import ROLES, install_styles


def normalize(doc, mapping):
    _, references, _ = discover(doc,mapping)
    ref_elements={p._p for p in references.values()}
    assignments = []
    for p in paragraphs(doc):
        kind=role(p,mapping)
        if p._p in ref_elements:kind='Reference'
        elif re.fullmatch(r'(?:\d+[.、\s]*)?(?:参考文献|References|Bibliography)',p.text.strip(),re.I):kind='H1'
        assignments.append((p,kind))
    install_styles(doc)
    changed, unknown = 0, []
    for p, kind in assignments:
        if kind not in ROLES:
            if p.text.strip(): unknown.append(p.text[:80])
            continue
        # Do not guess cover slots, drawings or definition hanging indents.
        p.style = 'XZIT' + kind
        pp = p._p.pPr
        if pp is not None:
            for tag in ('spacing','jc','outlineLvl') + (() if kind=='Definition' else ('ind',)):
                for node in list(pp.findall(qn('w:'+tag))): pp.remove(node)
        for r in p._p.xpath('./w:r | ./w:hyperlink/w:r'):
            rp = r.find(qn('w:rPr'))
            if rp is not None:
                for tag in ('rFonts','sz','szCs'):
                    for node in list(rp.findall(qn('w:'+tag))):rp.remove(node)
        changed += 1
    return {'mapped_paragraphs': changed, 'unmapped_paragraphs': unknown}


if __name__ == '__main__':
    ap=argparse.ArgumentParser();ap.add_argument('input',type=Path);ap.add_argument('output',type=Path)
    ap.add_argument('--role-map',type=Path,required=True);a=ap.parse_args()
    if a.output.exists():ap.error('Use a new output path')
    d=Document(a.input)
    result=normalize(d,json.loads(a.role_map.read_text(encoding='utf-8-sig')))
    d.save(a.output);print(json.dumps(result,ensure_ascii=False,indent=2))
