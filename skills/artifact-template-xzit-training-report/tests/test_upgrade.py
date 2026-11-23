"""Regression tests for report upgrades using synthetic documents."""
import json
from pathlib import Path
import sys
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from build_report import build
from content_guard import compare, snapshot
from document_roles import suggest_roles
from equations import equation_element
from metadata_audit import audit_metadata
from pdf_audit import inspect_pdf
from report_pipeline import run


class ReportUpgradeTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.folder = Path(tmp.name)
        self.data = json.loads((ROOT/'examples/report-input.json').read_text(encoding='utf-8'))

    def build(self, **options):
        return build(self.data,self.folder/'source.docx',**options)

    def test_new_numeric_spaces_and_course_size(self):
        self.data['blocks'][1]['text'] = '测量 15 kg，记录第　３次。Model 3 与型号\xa04。[1]'
        path = self.build(overrides={'Body':{'size_pt':12.5,'source':'测试任务书第2页'}})
        doc = Document(path)
        paragraph = next(p for p in doc.paragraphs if p.text.startswith('测量'))
        self.assertEqual(paragraph.text,'测量15 kg，记录第３次。Model 3 与型号\xa04。[1]')
        for tag in ('autoSpaceDE','autoSpaceDN'):
            self.assertEqual(paragraph._p.pPr.find(qn('w:'+tag)).get(qn('w:val')),'0')
        self.assertEqual(paragraph.style.font.size.pt,12.5)
        self.assertIn('1 工具演示',[p.text for p in doc.paragraphs])

    def test_read_only_audit_does_not_write_report(self):
        source = self.build()
        before = source.read_bytes()
        output = self.folder/'audit'
        run(source,output,check_only=True)
        self.assertEqual(source.read_bytes(),before)
        self.assertEqual({p.name for p in output.iterdir()},{'inspection.json','inspection.md'})
        result = json.loads((output/'inspection.json').read_text(encoding='utf-8'))
        self.assertEqual(result['input_sha256'],result['input_sha256_after'])
        with self.assertRaises(ValueError):
            run(source,self.folder/'rejected',check_only=True,word=True)

    def test_content_guard_blocks_promotion_of_changed_copy(self):
        source = self.build()
        from report_pipeline import format_document
        def mutate(doc,mapping):
            result = format_document(doc,mapping)
            doc.tables[-1].cell(1,1).text = 'Unauthorized change'
            return result
        destination = self.folder/'modified'
        with patch('report_pipeline.format_document',side_effect=mutate):
            with self.assertRaisesRegex(RuntimeError,'Content changed'):
                run(source,destination)
        self.assertFalse((destination/'report.docx').exists())
        record = json.loads((destination/'run.json').read_text(encoding='utf-8'))
        self.assertIn('text',record['content_guard']['formatted']['changed'])

    def test_guard_detects_equation_object_disappearance(self):
        source = self.build()
        doc = Document(source)
        formula = doc.element.xpath('.//m:oMath')[0]
        formula.getparent().remove(formula)
        changed = self.folder/'changed.docx'
        doc.save(changed)
        self.assertEqual(compare(changed,snapshot(source))['changed'],['equations'])

    def test_guard_detects_removed_figure(self):
        self.data['blocks'].append({'type':'figure','path':str(ROOT/'assets/preview.png'),'caption':'图1-1 示例'})
        source = self.build()
        doc = Document(source)
        drawing = doc.element.xpath('.//w:drawing')[-1]
        drawing.getparent().remove(drawing)
        changed = self.folder/'no-figure.docx'
        doc.save(changed)
        self.assertEqual(compare(changed,snapshot(source))['changed'],['images'])

    def test_mixed_style_does_not_propose_blanket_mapping(self):
        doc = Document()
        doc.styles.add_style('Custom',WD_STYLE_TYPE.PARAGRAPH)
        doc.add_paragraph('1.1 测试标题','Custom')
        doc.add_paragraph('这里是正文。','Custom')
        result = suggest_roles(doc)
        self.assertNotIn('Custom',result['proposed_role_map'])
        self.assertEqual(result['unmapped_styles']['Custom'][0]['role'],'H2')

    def test_pdf_catches_wrong_toc_page(self):
        doc = Document()
        doc.styles.add_style('TOC 1',WD_STYLE_TYPE.PARAGRAPH)
        doc.add_paragraph('1.1 测试标题\t2','TOC 1')
        pages = [{'lines':['1.1 测试标题 ..... 2'],'printed':('roman',1)},
                 {'lines':['1.1 测试标题'],'printed':('decimal',1)}]
        with patch('pdf_audit.read_pages',return_value=pages):
            report = inspect_pdf(doc,self.folder/'would-be.pdf')
        self.assertEqual(report['status'],'fail')
        self.assertIn('cached',report['issues'][0])

    def test_crossref_matches_using_only_doi(self):
        references = {1:SimpleNamespace(text='[1] Taylor. A Test Title. 2020. DOI:10.1234/example.'),
                      2:SimpleNamespace(text='[2] No identifier')}
        def fetch(doi):
            self.assertEqual(doi,'10.1234/example')
            return {'title':['A Test Title'],'author':[{'family':'Taylor'}],
                    'published':{'date-parts':[[2020]]}}
        result = audit_metadata(references,fetch)
        self.assertEqual([i['status'] for i in result['items']],['matched','pending'])

    def test_latex_requires_pandoc_and_only_one_input(self):
        with patch('equations.shutil.which',return_value=None):
            with self.assertRaisesRegex(ValueError,'Pandoc'):
                equation_element({'latex':r'\frac{x}{y}'})
        with self.assertRaisesRegex(ValueError,'exactly one'):
            equation_element({'latex':'x','math':'x'})
