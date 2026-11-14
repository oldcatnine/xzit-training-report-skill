# 徐州工程学院实训报告格式整理以及生成skill

依据用户保存的《实训报告模板.doc》制作的报告生成与检查技能。格式优先；以范本明确说明为依据，不声称是学校发布的最新统一标准。

**[详细使用文档](skills/artifact-template-xzit-training-report/references/usage-guide.md)**：环境准备、自然语言调用、完整示例、全部命令参数、旧稿映射、公式、续表、输出解释与常见问题。

配套：[可运行演示输入](skills/artifact-template-xzit-training-report/examples/report-input.json) · [输入字段规范](skills/artifact-template-xzit-training-report/references/input-schema.md) · [验证记录](VALIDATION.md)。演示内容及文献均非真实研究，不能直接提交。

## 使用

将 `skills/artifact-template-xzit-training-report` 放进个人技能目录（Codex：`~/.codex/skills/`），重新加载后输入：

> $artifact-template-xzit-training-report 根据我提供的新文献编写实训报告，严格执行保存的格式规范，生成 Word、PDF 和检查记录。

支持新报告、现有报告格式检查、只整理文献和提纲。姓名、学号、班级、指导教师、实训经历必须由用户提供；缺少时保留待补草稿，不编造。

已补充本机Word工作流、用户指定身份留空模式，以及按本次任务篇数运行的引用检查。用`audit_citations.py report.docx --min-references 10`可检查实际引用是否满足该次任务；数字由任务决定。

## 内容

正文引文现自动设置为可跳转的右上角上标；文后DOI使用完整超链接。正文三线表按保留范本使用1.5磅线条。现有报告可运行`scholarly_format.py input.docx output.docx`修正；Word小黑方点先按非打印格式标记检查，不删除分页设置。文字需逐段审读，少用套话与空泛自我说明。

- [格式与证据](skills/artifact-template-xzit-training-report/references/format-requirements.md)：23项明确格式规则与实际版式观测值。
- [文献工作流](skills/artifact-template-xzit-training-report/references/literature-workflow.md)：学习三份用户范文的结构，支持光伏回收及寿命预测任务。
- `report_pipeline.py`：统一生成/整理、Word字段更新、基于实际分页的自动续表、PDF/PNG导出与检查汇总。
- `build_report.py`：克隆范本生成封面、摘要、目录、正文、三线表、嵌入图、原生可编辑公式与参考文献；结论、致谢和附录可选。
- `normalize_report.py`：按显式角色映射整理旧稿，保留公式、图片、域、链接和上下标；未知段落列为待核。
- `audit_docx.py`：检查实际生效格式、节设置、关键词和残留占位；机器检查不代替引文核验与逐页视觉验收。
- `refresh_fields.ps1`：用 Windows Word 更新字段，另存DOCX和PDF。

依赖 Python 3.10+、python-docx、lxml，安装见 `requirements.txt`。从仓库根目录运行：

```sh
python -m unittest discover -s skills/artifact-template-xzit-training-report/tests -v
python skills/artifact-template-xzit-training-report/scripts/build_report.py input.json output.docx
python skills/artifact-template-xzit-training-report/scripts/audit_docx.py output.docx --output audit.json
python skills/artifact-template-xzit-training-report/scripts/report_pipeline.py input.json new-output --word --auto-continue --render --min-references 10
```

输入见 [schema](skills/artifact-template-xzit-training-report/references/input-schema.md)，统一入口及旧稿模式见[使用说明](skills/artifact-template-xzit-training-report/references/pipeline.md)。脚本拒绝覆盖已有输出；高级公式可传OMML，交叉引用等未覆盖功能用文档工具补齐并渲染检查。Word/Poppler为可选本机依赖；未更新目录的起稿不能称为终稿。文献证据记录不等于真实性核实，最终仍需来源与逐页核验。

GitHub参考：[format-thesis-docx](https://github.com/wksudud/format-thesis-docx)、[lab-report-skill](https://github.com/WinterMoonlight/lab-report-skill)，仅借鉴流程，无代码复制，详见技能内调研记录。

原创文字和脚本 MIT；`assets/reference.docx`及`assets/preview.png`保留原权利，不属于MIT授权，不代表学校背书。用户论文全文不随技能公开。
