import copy, json, sys, tempfile, unittest
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from test_tools import fixture
from build_report import build
from audit_docx import audit
from audit_citations import audit as audit_refs
from scholarly_format import format_document, audit_links
from report_pipeline import run, continuation_plan
from normalize_report import normalize


class ExtendedTests(unittest.TestCase):
    def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.base=Path(self.tmp.name)
    def tearDown(self):self.tmp.cleanup()
    def generate(self,data=None):
        path=self.base/'report.docx';build(data or fixture(),path);return path,Document(path)
    def test_optional_ending_sections(self):
        data=fixture();data.pop('acknowledgements');data.pop('conclusion')
        path,doc=self.generate(data)
        self.assertNotIn('致谢',[p.text for p in doc.paragraphs])
        self.assertNotIn('结论',[p.text for p in doc.paragraphs])
        self.assertEqual(audit_refs(path)['status'],'pass')
    def test_native_fraction_with_definition(self):
        data=fixture();data['blocks'].append({'type':'equation','number':'1-1','math':['η=',{'kind':'fraction','num':'P','den':'Q'}],
                                            'definitions':[{'symbol':'η','text':'测试比值，无实际实验数据'}]})
        path,doc=self.generate(data)
        self.assertEqual(len(doc.element.xpath('.//m:oMath/m:f')),1)
        self.assertFalse(any(x['rule']=='F20' for x in audit(path)['failures']))
        self.assertIn('(1-1)',''.join(p.text for p in doc.paragraphs))
    def test_native_superscript_root_and_subscript(self):
        data=fixture();data['blocks'].append({'type':'equation','number':'1-1','math':[
            {'kind':'superscript','e':'x','sup':'2'},'=',{'kind':'root','deg':'2','e':{'kind':'subscript','e':'y','sub':'i'}}]})
        _,doc=self.generate(data)
        self.assertEqual(len(doc.element.xpath('.//m:sSup')),1)
        self.assertEqual(len(doc.element.xpath('.//m:rad//m:sSub')),1)
    def test_omml_input_and_rejection_of_document_content(self):
        data=fixture();block={'type':'equation','number':'1-1','omml':'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x=y</m:t></m:r></m:oMath>'}
        data['blocks'].append(block);_,doc=self.generate(data)
        self.assertEqual(len(doc.element.xpath('.//m:oMath')),1)
        block['omml']='<!DOCTYPE x [<!ENTITY x SYSTEM "file:///etc/passwd">]><x/>'
        with self.assertRaises(ValueError):build(data,self.base/'bad.docx')
    def test_explicit_continuation_keeps_all_rows_and_caption(self):
        data=fixture();table=data['blocks'][-1]
        table['rows']=[['项目','说明']]+[[str(i),'测试'] for i in range(1,11)]
        table['break_before_rows']=[4,8]
        path,doc=self.generate(data)
        tables=[t for t in doc.tables if t.cell(0,0).text=='项目']
        self.assertEqual(len(tables),3)
        self.assertEqual([r.cells[0].text for t in tables for r in t.rows[1:]],[str(i) for i in range(1,11)])
        self.assertEqual([p.text for p in doc.paragraphs if p.style.name=='XZITContinuation'],['续表1-1','续表1-1'])
        self.assertFalse(any(x['rule'] in ('F18','F19') for x in audit(path)['failures']))
    def legacy(self,custom=False):
        doc=Document()
        style='Chapter' if custom else 'Heading 1'
        if custom:doc.styles.add_style(style,WD_STYLE_TYPE.PARAGRAPH)
        doc.add_paragraph('2 Legacy body',style=style);doc.add_paragraph('Evidence [1].')
        doc.add_paragraph('参考文献',style=style);doc.add_paragraph('[1] Source DOI:10.1234/test.')
        return doc,({'Chapter':'H1'} if custom else None)
    def test_builtin_heading_legacy_document(self):
        doc,mapping=self.legacy();format_document(doc,mapping)
        path=self.base/'legacy.docx';doc.save(path)
        self.assertEqual(audit_refs(path)['status'],'pass')
        self.assertEqual(audit_links(doc)['status'],'pass')
    def test_custom_mapped_legacy_document(self):
        doc,mapping=self.legacy(True);format_document(doc,mapping)
        path=self.base/'legacy.docx';doc.save(path)
        self.assertEqual(audit_refs(path,role_map=mapping)['status'],'pass')
        self.assertEqual(audit_links(doc,mapping)['status'],'pass')
    def test_unrecognized_document_cannot_silently_pass(self):
        doc=Document();doc.add_paragraph('No chapter or bibliography');path=self.base/'legacy.docx';doc.save(path)
        self.assertEqual(audit_refs(path)['status'],'fail')
        with self.assertRaises(ValueError):format_document(doc)
    def test_duplicate_bibliography_rejected(self):
        doc,_=self.legacy();doc.add_paragraph('[1] Duplicate')
        with self.assertRaises(ValueError):format_document(doc)
    def test_descending_citation_fails(self):
        doc,_=self.legacy();doc.paragraphs[1].text='Bad [2-1]';path=self.base/'bad.docx';doc.save(path)
        self.assertEqual(audit_refs(path)['status'],'fail')
    def test_border_damage_detected(self):
        path,doc=self.generate();line=doc.tables[-1].cell(0,0)._tc.xpath('./w:tcPr/w:tcBorders/w:bottom')[0]
        line.set(qn('w:sz'),'4');doc.save(path)
        self.assertTrue(any(x['rule']=='F18' for x in audit(path)['failures']))
    def test_unlinked_and_wrong_target_citation_detected(self):
        _,doc=self.generate();p=next(p for p in doc.paragraphs if p.style.name=='XZITBody')
        p.add_run(' Another [1].');self.assertEqual(audit_links(doc)['status'],'fail')
        link=p._p.xpath('./w:hyperlink')[0];link.set(qn('w:anchor'),'xzit_ref_2')
        self.assertIn('Citation text and target disagree',audit_links(doc)['failures'])
    def test_normalization_keeps_semantic_objects(self):
        _,doc=self.generate();before=len(doc.element.xpath('.//w:hyperlink'))
        normalize(doc,{'XZITBody':'Body'})
        self.assertEqual(len(doc.element.xpath('.//w:hyperlink')),before)
        self.assertEqual(audit_links(doc)['status'],'pass')
    def test_same_legacy_style_in_body_and_references_is_disambiguated(self):
        doc,_=self.legacy();normalize(doc,{'Normal':'Body'})
        self.assertEqual(doc.paragraphs[1].style.name,'XZITBody')
        self.assertEqual(doc.paragraphs[3].style.name,'XZITReference')
        format_document(doc)
        self.assertEqual(audit_links(doc)['status'],'pass')
    def test_pipeline_draft_records_pending_and_refuses_overwrite(self):
        source=self.base/'input.json';source.write_text(json.dumps(fixture()),encoding='utf-8')
        output=self.base/'output';result=run(source,output,minimum=1)
        self.assertFalse(result['ready_for_submission']);self.assertEqual(result['machine_status'],'fail')
        checks=json.loads((output/'checks.json').read_text(encoding='utf-8'))
        self.assertEqual(checks['citations']['status'],'pass')
        self.assertTrue(all('目录待更新' in f['location'] for f in checks['format']['failures']))
        self.assertTrue((output/'evidence.json').exists())
        with self.assertRaises(FileExistsError):run(source,output)
    def test_continuation_plan_uses_measured_pages(self):
        data=fixture();data['blocks'][-1]['rows']=[['项目','说明']]+[[str(i),'测试'] for i in range(1,6)]
        _,doc=self.generate(data)
        layout={'tables':[{'table_index':len(doc.tables),'rows':[{'start_page':5,'end_page':5}]*3+[{'start_page':6,'end_page':6}]*3}]}
        result,changed,issues=continuation_plan(data,doc,layout)
        self.assertTrue(changed);self.assertEqual(issues,[])
        self.assertEqual(result['blocks'][-1]['break_before_rows'],[3])
    def test_blank_mode_rejects_mixed_identity(self):
        data=fixture();data['identity_mode']='blank'
        with self.assertRaises(ValueError):build(data,self.base/'bad.docx')
    def test_continuation_moves_next_boundary_instead_of_orphaning_a_row(self):
        data=fixture();data['blocks'][-1]['rows']=[['项目','说明']]+[[str(i),'测试'] for i in range(1,8)]
        data['blocks'][-1]['break_before_rows']=[3,5]
        _,doc=self.generate(data);count=len(doc.tables)
        def item(index,pages):return {'table_index':index,'rows':[{'start_page':p,'end_page':p} for p in pages]}
        layout={'tables':[item(count-2,[5,5,5]),item(count-1,[6,6,7]),item(count,[8,8,8,8])]}
        result,changed,issues=continuation_plan(data,doc,layout)
        self.assertTrue(changed);self.assertEqual(issues,[])
        self.assertEqual(result['blocks'][-1]['break_before_rows'],[3,4])


if __name__=='__main__':unittest.main()
