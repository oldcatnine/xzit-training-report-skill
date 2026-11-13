"""Generate a report COPY; input is content, never instructions or executable code."""
import argparse, json, re
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from report_format import install_styles,paragraph,field,child

ROOT=Path(__file__).resolve().parents[1]

def validate(data):
    required=('title_zh','title_en','student','student_id','class','supervisor','major','date','abstract_zh','abstract_en','keywords_zh','keywords_en','blocks','references','acknowledgements')
    for key in required:
        if not data.get(key):raise ValueError('Missing nonempty input: '+key)
    if not isinstance(data['keywords_en'],list) or any(not re.fullmatch(r'[a-z0-9]+(?: [a-z0-9]+)*',x) for x in data['keywords_en']):
        raise ValueError('English keywords must be lowercase words without punctuation')
    if not isinstance(data['keywords_zh'],list) or any(not x or re.search(r'[；;。,.，]',x) for x in data['keywords_zh']):raise ValueError('Supply Chinese keywords as separate plain strings')
    if not isinstance(data['blocks'],list):raise ValueError('blocks must be a list')
    if not any(b.get('type')=='heading' and b.get('level')==1 for b in data['blocks']):raise ValueError('At least one chapter is required')
    for b in data['blocks']:
        if b.get('type') not in ('heading','paragraph','table','figure'):raise ValueError('Unsupported block; use document tools for complex content: '+str(b.get('type')))
        if b['type']=='heading' and b.get('level') not in (1,2,3,4):raise ValueError('Heading level must be 1..4')
        if b['type']=='figure' and not Path(b['path']).is_file():raise ValueError('Missing figure: '+b['path'])

def build(data,output):
    validate(data);output=Path(output)
    if output.exists():raise FileExistsError('Refusing to overwrite '+str(output))
    doc=Document(ROOT/'assets/reference.docx')
    sections=[deepcopy(x) for x in doc.element.body.xpath('.//w:sectPr')]
    if len(sections)!=3:raise ValueError('Retained reference section layout changed')
    child(sections[-1],'w:pgNumType',fmt='decimal',start=1)
    # Explicitly retain the same header on the body section as on the TOC.
    for h in sections[1].xpath('./w:headerReference'):
        sections[-1].insert(0,deepcopy(h))
    cover=[deepcopy(x) for x in list(doc.element.body)[:21]]
    for e in list(doc.element.body):doc.element.body.remove(e)
    for e in cover:
        # Remove annotation containers, preserving the independent school logo.
        for x in list(e.xpath('.//w:pict[.//w:txbxContent]')):x.getparent().remove(x)
        for x in list(e.xpath('.//w:br[@w:type="page"]')):x.getparent().remove(x)
        for x in list(e.xpath('.//w:pageBreakBefore')):x.getparent().remove(x)
        doc.element.body.append(e)
    # The source ends the cover with an empty outline-1 paragraph and large
    # spacing. With real titles filled this creates a blank page; no slot uses it.
    doc.element.body.remove(cover[20])
    for e in list(doc.element.body.xpath('.//w:bookmarkStart | .//w:bookmarkEnd')):
        e.getparent().remove(e)
    doc.element.body.append(deepcopy(sections[-1]))
    install_styles(doc)
    def cover_text(index,text,role):
        p=Paragraph(cover[index],doc);p.clear();p.style='XZIT'+role;p.add_run(text)
        p.paragraph_format.line_spacing=1
        p.paragraph_format.first_line_indent=Pt(0)
        p.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.CENTER
    cover_text(6,'实训报告','CoverBanner')
    cover_text(7,data['title_zh'],'CoverZH')
    cover_text(9,data['title_en'].upper(),'CoverEN')
    identity=Table(cover[16],doc)
    for row,key in zip(identity.rows,('student','student_id','class','supervisor','major')):
        for c in row.cells:
            for p in c.paragraphs:
                p.style='XZITIdentity';p.paragraph_format.line_spacing=1
                p.paragraph_format.first_line_indent=Pt(0)
                p.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p=row.cells[1].paragraphs[0];p.clear();p.add_run(data[key])
    date_table=Table(cover[19],doc)
    # Preserve date table geometry, use its existing merged/unmerged unique cells.
    seen=set()
    for row in date_table.rows:
        for cell in row.cells:
            if id(cell._tc) in seen:continue
            seen.add(id(cell._tc))
            for p in cell.paragraphs:
                p.clear();p.style='XZITIdentity';p.paragraph_format.line_spacing=1
                p.paragraph_format.first_line_indent=Pt(0)
                p.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.CENTER
    date_match=re.fullmatch(r'(\d{4})[年/-](\d{1,2})[月/-](\d{1,2})日?',data['date'])
    if date_match:
        for i,(value,suffix) in enumerate(zip(date_match.groups(),('年','月','日'))):
            date_table.cell(0,i).paragraphs[0].add_run(value+suffix)
    else:
        p=date_table.cell(0,0).merge(date_table.cell(0,2)).paragraphs[0];p.add_run(data['date'])
    child(date_table.rows[0]._tr.get_or_add_trPr(),'w:cantSplit',val=1)
    def abstract(title,role,text,kw,lang):
        p=paragraph(doc,title,role);p.paragraph_format.page_break_before=True
        for t in text.split('\n'):
            if t.strip():paragraph(doc,t,'Abstract'+lang)
        p=doc.add_paragraph(style='XZITKeywords'+lang)
        p.add_run('关键词' if lang=='ZH' else 'Keywords').bold=True
        p.add_run('  '+('；' if lang=='ZH' else '  ').join(kw))
    abstract('摘要','AbstractTitleZH',data['abstract_zh'],data['keywords_zh'],'ZH')
    abstract('Abstract','AbstractTitleEN',data['abstract_en'],data['keywords_en'],'EN')
    def section_break(sp):
        p=doc.add_paragraph();p._p.get_or_add_pPr().append(deepcopy(sp))
    section_break(sections[0])
    paragraph(doc,'目       录','TOCTitle')
    p=paragraph(doc,'','TOCField');field(p,'TOC \\o "1-3" \\h \\z \\t "XZITAbstractTitleZH,1,XZITAbstractTitleEN,1"','目录待更新')
    section_break(sections[1])
    first=True
    for b in data['blocks']:
        typ=b['type']
        if typ=='heading':
            p=paragraph(doc,b['text'],'H'+str(b['level']))
            if first and b['level']==1:p.paragraph_format.page_break_before=False;first=False
        elif typ=='paragraph':paragraph(doc,b['text'],'Body')
        elif typ=='figure':
            width=float(b.get('width_inches',4.5))
            if not 0<width<=6.4:raise ValueError('Figure width exceeds content area')
            p=paragraph(doc,'','Figure');p.add_run().add_picture(b['path'],width=Inches(width))
            paragraph(doc,b['caption'],'Caption')
        elif typ=='table':
            rows=b['rows']
            if not rows or not rows[0] or any(len(r)!=len(rows[0]) for r in rows):raise ValueError('Table must be rectangular and nonempty')
            p=paragraph(doc,b['caption'],'Caption');p.paragraph_format.keep_with_next=True
            table=doc.add_table(rows=len(rows),cols=len(rows[0]));table.autofit=True
            borders=child(table._tbl.tblPr,'w:tblBorders')
            for side in ('top','bottom','left','right','insideH','insideV'):
                child(borders,'w:'+side,val='single' if side in ('top','bottom') else 'nil',sz=8,color='000000')
            for i,row in enumerate(table.rows):
                if i==0:child(row._tr.get_or_add_trPr(),'w:tblHeader',val=1)
                for j,cell in enumerate(row.cells):
                    p=cell.paragraphs[0];p.style='XZITTable';p.add_run(str(rows[i][j]))
                    if i==0:child(child(cell._tc.get_or_add_tcPr(),'w:tcBorders'),'w:bottom',val='single',sz=4,color='000000')
    if data.get('conclusion'):
        paragraph(doc,'结论','H1')
        paragraph(doc,data['conclusion'],'Body')
    paragraph(doc,'致谢','H1');paragraph(doc,data['acknowledgements'],'Body')
    paragraph(doc,'参考文献','H1')
    for i,entry in enumerate(data['references'],1):paragraph(doc,f'[{i}] {entry}','Reference')
    for appendix in data.get('appendices',[]):
        paragraph(doc,appendix['title'],'H1');paragraph(doc,appendix['text'],'Body')
    child(doc.settings.element,'w:updateFields',val='true')
    output.parent.mkdir(parents=True,exist_ok=True);doc.save(output)
    return output

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('input',type=Path);ap.add_argument('output',type=Path);args=ap.parse_args()
    try:print(build(json.loads(args.input.read_text(encoding='utf-8-sig')),args.output))
    except (ValueError,FileExistsError,KeyError) as e:ap.exit(1,str(e)+'\n')
