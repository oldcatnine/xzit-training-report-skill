"""Conservative DOCX audit. Exit 0: no detected failures, NOT submission approval."""
import argparse,json,re
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from report_format import ROLES

def merge(props,element):
    if element is None:return
    for e in element:
        tag=e.tag.rsplit('}',1)[-1]
        # Composite values must merge; e.g. a run overriding ASCII doesn't erase East Asia.
        attrs={k.rsplit('}',1)[-1]:v for k,v in e.attrib.items()}
        if tag in ('b','i','keepNext','pageBreakBefore','vanish'):
            attrs.setdefault('val','1')
        props.setdefault(tag,{}).update(attrs)

def effective(p,run=None):
    pp={};rp={};chain=[];style=p.style;seen=set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id);chain.append(style);style=style.base_style
    defaults=p.part.document.styles.element.find(qn('w:docDefaults'))
    if defaults is not None:
        for parent,dest,name in [('w:pPrDefault',pp,'w:pPr'),('w:rPrDefault',rp,'w:rPr')]:
            x=defaults.find(qn(parent))
            if x is not None:merge(dest,x.find(qn(name)))
    for s in reversed(chain):merge(pp,s.element.find(qn('w:pPr')));merge(rp,s.element.find(qn('w:rPr')))
    merge(pp,p._p.find(qn('w:pPr')))
    if run is not None:
        st=run.style
        rchain=[];seen=set()
        while st is not None and st.style_id not in seen:
            seen.add(st.style_id);rchain.append(st);st=st.base_style
        for s in reversed(rchain):merge(rp,s.element.find(qn('w:rPr')))
        merge(rp,run._r.find(qn('w:rPr')))
    return pp,rp

def audit(path,role_map=None):
    doc=Document(path);fail=[];unknown=[];checked=0
    def error(rule,where,detail):fail.append({'rule':rule,'location':where,'detail':detail})
    role_map=role_map or {}
    for i,sec in enumerate(doc.sections,1):
        for name,target in [('page_width',11906),('page_height',16838),('top_margin',1418),('bottom_margin',1418),('left_margin',1418),('right_margin',1134),('header_distance',851),('footer_distance',851),('gutter',0)]:
            value=getattr(sec,name)
            if value is None or abs(value.twips-target)>2:error('O01-O03',f'section {i}',name+' differs from retained reference')
    if len(doc.sections)!=3:error('O04','sections','Expected retained three-section structure; review explicit deviations')
    else:
        for i,expected in enumerate(('upperRoman','upperRoman','decimal')):
            x=doc.sections[i]._sectPr.find(qn('w:pgNumType'))
            if x is None or x.get(qn('w:fmt'),'decimal')!=expected:error('O04',f'section {i+1}','Page number format mismatch')
            if i==2 and x is not None and x.get(qn('w:start'))!='1':error('O04','section 3','Main body must restart at 1')
    root=doc.element
    if root.xpath('.//w:txbxContent'):error('C','document','Instruction/floating textboxes remain; inspect each')
    if root.xpath('.//wp:anchor'):error('F22','document','Floating image/shape found; inspect and convert figures to inline')
    instr=' '.join(root.xpath('.//w:instrText/text()')+root.xpath('.//w:fldSimple/@w:instr'))
    if not re.search(r'TOC\s+.*?\\o\s+"1-3"',instr):error('F11','TOC','Missing TOC field limited to levels 1-3')
    texts=[];roles_seen=set()
    for i,e in enumerate(root.body.xpath('.//w:p[not(ancestor::w:txbxContent)]'),1):
        p=Paragraph(e,doc);text=p.text.strip();texts.append(text)
        if not text:continue
        where=f'paragraph {i}: '+text[:45]
        if re.search(r'错误\s*[!！]|Error!|未定义书签|目录待更新|×{2,}|待补[：:]|格式说明',text,re.I):error('C',where,'Unresolved placeholder or field error')
        name=p.style.name
        role=role_map.get(name,name[4:] if name.startswith('XZIT') else None)
        toc_match=re.fullmatch(r'toc ([123])',name,re.I)
        if toc_match:role='TOC'+toc_match.group(1)
        if role not in ROLES:
            # Field-generated TOC entries have their own controlled styles.
            if not re.fullmatch(r'TOC [123]',name,re.I):unknown.append(where+' [style='+name+']')
            continue
        checked+=1;roles_seen.add(role)
        font,size,line,indent,align,outline=ROLES[role];pp,_=effective(p)
        spacing=pp.get('spacing',{});ind=pp.get('ind',{})
        expected_line='400' if line=='exact' else '240';expected_rule='exact' if line=='exact' else 'auto'
        if spacing.get('line','240')!=expected_line or spacing.get('lineRule','auto')!=expected_rule:error('F-line',where,'Expected '+line+' spacing')
        first=ind.get('firstLineChars')
        if first is not None:
            if int(first)!=(200 if indent else 0):error('F-indent',where,'Incorrect character first-line indent')
        elif int(ind.get('firstLine','0'))!=int(indent*20):error('F-indent',where,'Incorrect first-line indent')
        if pp.get('jc',{}).get('val')!=align:error('F-align',where,'Expected '+align)
        if outline is not None:
            if pp.get('outlineLvl',{}).get('val')!=str(outline):error('F13-F16',where,'Incorrect outline level')
            if spacing.get('beforeLines')!='50' or spacing.get('afterLines')!='50':error('F13-F16',where,'Heading needs 0.5 LINE before and after')
        for r in p.runs:
            if not r.text.strip():continue
            _,rp=effective(p,r);rf=rp.get('rFonts',{})
            if rp.get('sz',{}).get('val')!=str(int(size*2)):error('F-font',where,'Incorrect effective font size')
            if re.search(r'[\u3400-\u9fff]',r.text) and rf.get('eastAsia')!=font:error('F-font',where,'Incorrect effective Chinese font')
            latin=font if font in ('Arial','Times New Roman') else 'Times New Roman'
            if re.search('[A-Za-z]',r.text) and rf.get('ascii')!=latin:error('F-font',where,'Incorrect effective Latin font')
            if role=='AbstractTitleEN' and rp.get('b',{}).get('val','1') in ('0','false','off'):error('F08',where,'Abstract title must be bold')
        if role=='CoverEN' and text!=text.upper():error('F03',where,'English title not uppercase')
        if role in ('KeywordsZH','KeywordsEN'):
            label='关键词' if role.endswith('ZH') else 'Keywords'
            value=text[len(label):].strip() if text.startswith(label) else ''
            pattern=r'[^；;，。,]+(?:；[^；;，。,]+)*' if role.endswith('ZH') else r'[a-z0-9]+(?: [a-z0-9]+)*(?:  [a-z0-9]+(?: [a-z0-9]+)*)*'
            if not value or not re.fullmatch(pattern,value):error('F07/F10',where,'Keyword separator/case/punctuation mismatch')
            for r in p.runs:
                if not r.text.strip():continue
                _,rp=effective(p,r);bold='b' in rp and rp['b'].get('val','1') not in ('0','false','off')
                if bold!=(r.text.strip()==label):error('F07/F10',where,'Only keyword label may be bold (label must be separate run)')
    for role in ('CoverZH','CoverEN','AbstractZH','AbstractEN','KeywordsZH','KeywordsEN','H1','Reference'):
        if role not in roles_seen:error('structure','document','Missing or unmapped role: '+role)
    manual=[
      'Render every page: cover, fonts, breaks, header/footer and page numbering.',
      'Verify all factual claims, references, citation order and real training evidence.',
      'Verify table/figure numbering, caption placement, three-line borders and continuation labels on EVERY continued page.',
      'Verify editable equations, right-aligned numbers, first-use variable definitions and wrapped alignment.',
      'Verify updated TOC entries, actual level fonts and cached page numbers in Word/PDF.'
    ]
    if unknown:manual.append({'unmapped_paragraphs':unknown})
    return {'machine_status':'fail' if fail else 'pass','ready_for_submission':False,'checked_paragraphs':checked,'failures':fail,'manual_review_required':manual,'note':'Machine pass is not visual/content approval. Unknown roles require mapping or documented manual review.'}

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('docx',type=Path);ap.add_argument('--output',type=Path);ap.add_argument('--role-map',type=Path);args=ap.parse_args()
    if args.output and args.output.exists():ap.error('Refusing to overwrite audit output')
    result=audit(args.docx,json.loads(args.role_map.read_text(encoding='utf-8-sig')) if args.role_map else None)
    payload=json.dumps(result,ensure_ascii=False,indent=2)
    if args.output:args.output.write_text(payload,encoding='utf-8')
    print(payload);raise SystemExit(bool(result['failures']))
