"""Shared format definitions; all sizes are points unless stated otherwise."""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE

ROLES = {
 'Body': ('宋体',12,'exact',24,'both',None),
 'AbstractZH': ('宋体',12,'exact',24,'both',None),
 'AbstractEN': ('Times New Roman',12,'exact',24,'left',None),
 'AbstractTitleZH': ('黑体',18,'single',0,'center',None),
 'AbstractTitleEN': ('Times New Roman',18,'single',0,'center',None),
 'KeywordsZH': ('宋体',12,'exact',0,'left',None),
 'KeywordsEN': ('Times New Roman',12,'exact',0,'left',None),
 'H1': ('黑体',18,'single',0,'center',0),
 'H2': ('黑体',15,'single',0,'left',1),
 'H3': ('黑体',14,'single',0,'left',2),
 'H4': ('黑体',12,'single',0,'left',3),
 'Caption': ('宋体',10.5,'single',0,'center',None),
 'Table': ('宋体',10.5,'single',0,'center',None),
 'Reference': ('宋体',10.5,'exact',0,'left',None),
 'CoverZH': ('黑体',22,'single',0,'center',None),
 'CoverEN': ('Arial',22,'single',0,'center',None),
 'CoverBanner': ('黑体',42,'single',0,'center',None),
 'Identity': ('黑体',15,'single',0,'center',None),
 'TOCTitle': ('黑体',16,'exact',0,'center',None),
 'TOCField': ('宋体',12,'exact',0,'left',None),
 'Figure': ('宋体',10.5,'single',0,'center',None),
 'TOC1': ('黑体',12,'exact',0,'left',None),
 'TOC2': ('宋体',12,'exact',0,'left',None),
 'TOC3': ('宋体',12,'exact',0,'left',None),
 'Equation': ('Times New Roman',12,'single',0,'left',None),
 'Definition': ('宋体',12,'exact',0,'left',None),
}

def child(parent, tag, **attrs):
    e=parent.find(qn(tag))
    if e is None: e=OxmlElement(tag); parent.append(e)
    for key,value in attrs.items(): e.set(qn('w:'+key),str(value))
    return e

def install_styles(doc):
    for role,(font,size,line,indent,align,outline) in ROLES.items():
        name='XZIT'+role
        s=doc.styles[name] if name in doc.styles else doc.styles.add_style(name,WD_STYLE_TYPE.PARAGRAPH)
        s.base_style=doc.styles['Normal']
        s.font.name=font; s.font.size=Pt(size); s.font.bold=role in ('AbstractTitleEN','TOCTitle')
        s.font.color.rgb=None
        child(s.element.get_or_add_rPr(),'w:rFonts',eastAsia=font,ascii=(font if font in ('Arial','Times New Roman') else 'Times New Roman'),hAnsi=(font if font in ('Arial','Times New Roman') else 'Times New Roman'))
        pf=s.paragraph_format
        pf.alignment={'center':WD_ALIGN_PARAGRAPH.CENTER,'left':WD_ALIGN_PARAGRAPH.LEFT,'both':WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
        pf.first_line_indent=Pt(indent);pf.left_indent=Pt(0);pf.right_indent=Pt(0)
        pf.space_before=Pt(0);pf.space_after=Pt(0)
        if line=='exact':pf.line_spacing=Pt(20)
        else:pf.line_spacing=1
        pp=s.element.get_or_add_pPr()
        if outline is not None:
            child(pp,'w:outlineLvl',val=outline)
            child(pp,'w:spacing',beforeLines=50,afterLines=50)
            pf.keep_with_next=True
            pf.page_break_before=outline==0
        else:child(pp,'w:outlineLvl',val=9)
    for level in (1,2,3):
        name='toc '+str(level)
        s=doc.styles[name] if name in doc.styles else doc.styles.add_style(name,WD_STYLE_TYPE.PARAGRAPH,builtin=True)
        s.base_style=doc.styles['XZITTOCField'];s.font.name='黑体' if level==1 else '宋体';s.font.size=Pt(12)
        child(s.element.get_or_add_rPr(),'w:rFonts',eastAsia='黑体' if level==1 else '宋体',ascii='Times New Roman',hAnsi='Times New Roman')
        s.paragraph_format.left_indent=Pt((level-1)*9)

def paragraph(doc,text,role):
    return doc.add_paragraph(str(text),style='XZIT'+role)

def field(p,instruction,placeholder=''):
    for kind in ('begin','instruction','separate','text','end'):
        r=OxmlElement('w:r')
        if kind=='instruction':
            e=OxmlElement('w:instrText');e.set(qn('xml:space'),'preserve');e.text=' '+instruction+' '
        elif kind=='text':e=OxmlElement('w:t');e.text=placeholder
        else:e=OxmlElement('w:fldChar');e.set(qn('w:fldCharType'),kind)
        r.append(e);p._p.append(r)
