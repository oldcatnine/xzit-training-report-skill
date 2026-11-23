# 详细使用文档

本技能依据用户保存的徐州工程学院实训报告范本制作，不代表学校最新统一政策。教师任务书与用户本次明确要求优先于工具默认设置。演示身份、内容和文献不能直接用于提交。

## 目录

- [功能与使用方式](#功能与使用方式)
- [首轮选择处理方式](#首轮选择处理方式)
- [环境准备与安装](#环境准备与安装)
- [自然语言调用示例](#自然语言调用示例)
- [从JSON生成新报告](#从json生成新报告)
- [统一入口参数](#统一入口参数)
- [输入内容与编号](#输入内容与编号)
- [可编辑公式](#可编辑公式)
- [表格和续表](#表格和续表)
- [整理已有DOCX](#整理已有docx)
- [参考文献与证据记录](#参考文献与证据记录)
- [输出与状态说明](#输出与状态说明)
- [单独运行工具](#单独运行工具)
- [常见问题](#常见问题)
- [交付前核验](#交付前核验)

## 功能与使用方式

| 目的 | 使用方式 | 自动处理范围 |
|---|---|---|
| 根据真实资料编写新报告 | 在Codex调用技能，提供任务书与文献 | 助手阅读、组织和写作，再用脚本排版与检查 |
| 已准备全部内容，生成Word | JSON + `report_pipeline.py` | 封面、摘要、目录域、正文、图表、公式、文献；可选Word更新及PDF/PNG |
| 整理已有报告 | DOCX + 样式映射 + `--normalize` | 已识别段落格式、引文链接与已映射正文表格 |
| 只检查，不修改 | `audit_docx.py`、`audit_citations.py` | 输出检查记录，输入文件不变 |
| 只修引文、文献表和表线 | `scholarly_format.py` | 合并上标引文、文献跳转、作者大小写、移除 DOI、已识别正文三线表 |

脚本消费准备好的内容，不会自行检索文献、阅读全文或根据题目生成科学结论。内容写作、原文证据判断和逐页验收由调用技能的助手或作者完成。

## 首轮选择处理方式

处理现有报告时，若没有明确说要检查还是修改，助手会先提示三种方式：

| 选择 | 实际操作 | 交付结果 |
|---|---|---|
| 1 仅检查格式 | 只读检查，不保存修复稿 | 带位置、依据和建议的问题清单、检查记录 |
| 2 直接修改副本 | 保留原件，按范本修复已授权格式问题并复核 | 新Word副本、可用PDF、修改说明及待核项 |
| 3 先检查、确认后修改 | 先检查并给出具体方案，等待确认后再改 | 先交问题/方案，获批后交修改副本和复核结果 |

可回复数字或直接限定范围，如“只检查，不修改”“直接改副本，只调字体和行距”“先给我问题清单，确认再改”。已明确选择时不重复询问；同一任务沿用该选择，用户可以随时改选。没有回复不代表允许修改。

方式2默认处理格式，不授权改研究数据、结论或公开上传。方式3需要对具体方案的确认后才进入修复。明确的新建报告任务直接按新建要求执行，不强迫选择旧稿检查方式。完整提示文案与工具边界见[首轮对话提示](first-turn-dialog.md)。这一提示由助手执行，CLI脚本本身不弹出选择框。

## 环境准备与安装

### 运行条件

- Python 3.10+。运行 `pip install -r requirements.txt` 安装 python-docx、lxml 和 PDF 审计所用的 pypdf。
- `--word`要求Windows及可通过COM正常启动的Microsoft Word。
- `--render`要求Poppler的`pdftoppm`位于PATH，且同时启用`--word`。
- 主要字体为宋体、黑体、Times New Roman和Arial；公式及字体替换仍需在最终Word/PDF核对。

从仓库根目录安装Python依赖：

```powershell
python -m pip install -r requirements.txt
python --version
python -c "import docx, lxml; print('Python依赖可用')"
Get-Command pdftoppm
```

脚本不自动安装Word、字体或Poppler。无Word可先起稿，目录、页码及最终分页须在后续文档流程中更新。

### 个人技能目录

将仓库内`skills/artifact-template-xzit-training-report`整个目录放入个人技能目录，保留`SKILL.md`、`scripts`、`references`、`assets`、`agents`、`examples`的相对结构。不要只复制SKILL.md，不改动`assets/reference.docx`。

本指南后续命令均在**技能根目录**运行；在仓库根目录时先执行：

```powershell
Set-Location skills/artifact-template-xzit-training-report
```

已安装版本可在其自身目录执行相同命令。输出位置需要写入权限。

## 自然语言调用示例

新报告：

> $artifact-template-xzit-training-report 根据附件任务书和论文全文编写报告。主题为……，本次至少实际引用……篇文献，身份信息如下……。不编造实验经历，生成Word、PDF和检查记录，核对全部页面。

身份手填：

> 使用这个技能生成报告。姓名、学号、班级、教师和专业都留空，保留填写栏，由我后续手填。其他内容依据附件。

整理旧稿：

> 使用这个技能整理附件DOCX，只调整格式和引文链接，保留结论、数据、图片和可编辑公式。说明文献与页面中还需要核实的内容。

提供报告类型、主题、任务书要求、已有资料与来源、封面信息和日期。字数、篇数、语种比例由本次任务决定，不从样例推定为全校要求。

## 从JSON生成新报告

### 运行演示

附带[可运行演示输入](../examples/report-input.json)，含显式空身份、示例表格和原生分式。其文献明确为演示条目，不是真实出版物，不能提交为科研报告。

仅起稿：

```powershell
python scripts/report_pipeline.py examples/report-input.json test-output/demo-draft --min-references 1
```

命令会生成文件，但目录尚未更新，格式检查预期报告“目录待更新”并退出1。这是草稿状态，不能当终稿。

本机Word完整流程：

```powershell
python scripts/report_pipeline.py examples/report-input.json test-output/demo-word --word --auto-continue --render --min-references 1
```

依次生成副本、更新目录与分页、测量表格行页码、必要时重建续表、导出PDF、渲染PNG并汇总检查。新输出目录必须不存在；重跑请换新目录名。

### 使用真实材料

把演示另存为`my-report.json`，替换题目、摘要、正文、文献和证据。完整字段见[输入规范](input-schema.md)。

```powershell
python scripts/report_pipeline.py my-report.json report-output-v1 --word --auto-continue --render --min-references 10
```

数字10仅为示例任务条件，应替换为实际要求；无最低数量时可省略参数。图片路径可以是绝对路径；统一入口中的相对图片路径相对JSON文件所在目录解析。

## 统一入口参数

### 检查、保全与联网选项

```powershell
python scripts/report_pipeline.py existing.docx audit-v1 --check-only --min-references 10
python scripts/report_pipeline.py existing.docx repaired-v1 --role-map roles.json --normalize --word
```

`--check-only` 只写 `inspection.md` 与 `inspection.json`，同时核对源文件哈希。不与 Word、渲染或修复选项混用。样式建议含预览、段落位置和依据，可能存在待核角色。确认后可将合适的建议写入 `roles.json`。必要时使用 `--course-rules course.json` 指定按角色覆盖的字号，例如 `{"Body":{"size_pt":12.5,"source":"课程任务书第2页"}}`。修整旧稿时须同时提供 `--normalize` 和显式角色映射。生成和审计使用相同的覆盖规则。

旧稿修复在格式调整后及 Word 另存后比较正文、表格、图片与公式。发现差异时输出 `run.json` 中的首处差异与阶段，修复候选保留为中间稿。PDF 核对结果见 `checks.json` 的 `pdf`，检查印刷页码的连续性与目录标题实际落页。重复标题或无法提取的页码标待核。

用户希望联网对照时，可加 `--crossref`。工具只向 Crossref 传文献的 DOI，核对题名、首位作者姓氏与年份，并记录缺项、未匹配及查询失败。文献能否支持报告论断依然需要读原文。

新稿数字间距修正及可选 LaTeX 片段用法见[输入规范](input-schema.md)。

| 参数 | 默认值 | 含义与限制 |
|---|---|---|
| `input` | 必填 | UTF-8 JSON内容文件或已有DOCX |
| `output_dir` | 必填 | 全新目录，已存在即拒绝 |
| `--min-references N` | 0 | 非负的最低参考条目数，仍检查引用对应 |
| `--word` | 关闭 | 本机Word更新字段、导出PDF并记录页数 |
| `--auto-continue` | 关闭 | 仅JSON输入；必须同时加`--word`，按实际分页调整续表 |
| `--max-passes N` | 5 | 最多分页轮数，允许1—8，未稳定则失败 |
| `--render` | 关闭 | PDF渲染为PNG，要求`--word`和pdftoppm |
| `--role-map PATH` | 无 | DOCX原段落样式名到角色名的JSON映射 |
| `--normalize` | 关闭 | 按角色整理DOCX，要求显式`--role-map` |
| `--check-only` | 关闭 | 只读检查 DOCX，输出问题与拟修复清单 |
| `--course-rules PATH` | 无 | JSON 课程字号覆盖，记录规则出处 |
| `--crossref` | 关闭 | 向 Crossref 查询原始 DOCX 或 JSON 输入中的 DOI |

`--normalize`与样式映射面向DOCX；JSON稿由生成器直接使用本技能样式，应省略这些旧稿选项。

Word子进程仅对本技能脚本使用进程级ExecutionPolicy Bypass，不修改用户或机器的持久执行策略。它创建隐藏Word实例，只关闭自己创建的文档和实例。

## 输入内容与编号

| 对象 | 字段及规则 |
|---|---|
| 封面 | `title_zh`、`title_en`、`student`、`student_id`、`class`、`supervisor`、`major`、`date` |
| 摘要 | `abstract_zh`、`abstract_en`，换行分段 |
| 关键词 | `keywords_zh`、`keywords_en`分别为数组，英文小写、不带标点 |
| 标题 | `type: heading`，`level: 1..4`，`text`内含要显示的编号 |
| 正文 | `type: paragraph`、`text`；引用写成`[1]`等 |
| 表格 | `type: table`、`caption`、矩形`rows`数组，首行为表头 |
| 图片 | `type: figure`、`path`、`caption`；`width_inches`默认4.5，允许大于0且不超过6.4 |
| 公式 | `type: equation`、`number`；`math`与`omml`二选一 |
| 结论/致谢 | `conclusion`、`acknowledgements`可省略、null或空字符串，为空不生成 |
| 文献 | 非空字符串数组`references`，生成器添加从1起的[n]编号 |
| 附录 | `appendices`数组，每项含`title`和`text`，允许空数组 |
| 证据 | 可选`evidence`数组，编号对应已有文献条目 |

普通模式五项身份必须非空。明确选择手填时，设置`identity_mode: blank`并将五项全部显式传空字符串，不省略、不混填。不能虚构身份或经历。

标题、图表题、公式编号由调用者按章填写。目录显示至三级，章及结尾部分另起新页。封面不显示页码，第二页（中文摘要）从大写罗马数字 I 开始，正文从阿拉伯数字 1 开始。英文摘要正文使用两端对齐，Abstract 标题居中。

## 可编辑公式

`"math":"E=mc²"`生成线性可编辑对象；字符串只显示文本，不解释LaTeX、不计算。真正的分式或下标使用结构节点：

```json
{"type":"equation","number":"2-1","math":["η=",{"kind":"fraction","num":{"kind":"subscript","e":"P","sub":"out"},"den":{"kind":"subscript","e":"P","sub":"in"}}],"definitions":[{"symbol":"η","text":"变量含义、单位和适用条件"}]}
```

| kind | 必需字段 |
|---|---|
| `fraction` | `num`、`den` |
| `subscript` | `e`、`sub` |
| `superscript` | `e`、`sup` |
| `root` | `deg`、`e`，平方根deg可写字符串`2` |

支持字符串列表及节点嵌套，最大深度32。编号输入`2-1`，显示`(2-1)`；制表位使公式居中、编号靠右。释义用制表位和悬挂缩进对齐，复杂符号转行须看实际页面。

矩阵、积分等可用单个带命名空间的 `m:oMath` XML 字符串传给 `omml`。如果 PATH 已配置 Pandoc，也可通过 `latex` 字段提供单行数学片段。见[输入规范](input-schema.md#可编辑公式)。

默认DOCX兼容级别15。若Word保存时提示公式会转换成图片，应取消并检查兼容模式，不盲目继续。

## 表格和续表

三线表上下边线各 1.5 磅，表头下线 0.5 磅。数据行之间和表格左右两侧留白。封面身份栏沿用其专用格式。表题在上，图题在下。

`widths_cm`列数与表格列数相同，各列为正，总宽不超过16.5cm；未传则自动调整。短表尽量同页，实际高度必须核对。

### 自动续表

使用`--word --auto-continue`。工具按实际行页码重建跨页边界，各段重复首行表头，续段新页右上方生成“续表1-1”。表头挤走末行时会调整相邻边界，避免留下多余单行续表页；不保证消除所有内容本身造成的短页。

单行跨页、孤立表头或超轮数仍未稳定会记录失败，不删除数据或缩小字号强行通过。只能对JSON生成稿自动重建；现有DOCX复杂合并表格与任意版式变动需在副本中处理。

### 指定续表边界

```json
{"type":"table","caption":"表1-1 对比表","rows":[["项目","说明"],["A","说明"],["B","说明"],["C","说明"]],"break_before_rows":[3]}
```

rows索引0为表头，索引1为首行数据，这里在C行之前拆表。边界为递增不重复整数，范围2至`len(rows)-1`。非“表章号-序号”开头的题注需提供`continuation_label`。每个续页都应复查表头、标签、线条与数据；改内容或更换排版环境后重新核验。

## 整理已有DOCX

### 样式映射

`roles.json`示例：

```json
{"Heading 1":"H1","正文":"Body","文献条目":"Reference","表内文字":"Table","图表题":"Caption"}
```

键是文档实际段落样式名，值是技能角色。完整角色见[scripts/report_format.py](../scripts/report_format.py)的ROLES。不能把所有未知段落都当正文，也不能把题注或摘要误映射为Body。

内置Heading/标题1—4与XZIT样式直接识别。正文从第一个编号章标题开始，可为`1 绪论`、`2 …`或`第一章 …`；自定义章样式映射到H1。参考文献区按独立“参考文献/References/Bibliography”标题和[n]条目识别。正文和文献共用Normal时，整理器依据文献区域区分条目。

### 整理副本

```powershell
python scripts/report_pipeline.py existing.docx report-repaired-v1 --role-map roles.json --normalize --word --render
```

处理已识别段落的字体、字号、行距、对齐及缩进冲突；保留加粗、斜体、上下标、图片、原生公式、域及链接对象，保留公式释义的原悬挂缩进。仅处理已识别的正文表线。

工具不重建旧稿封面、不自动修正全部节的页边距、不猜测未知段落。封面、分节及页码偏差仍须根据检查另行修复。未加`--normalize`时主要处理引文链接、上标和已识别表线，再做检查，不全篇重排。

只检查时使用统一入口的 `--check-only`，或后文独立审计命令。

## 参考文献与证据记录

正文引用生成黑色、无下划线的上标跳转。文后编号保持基线。文献数组按首次引用顺序排列，同一条多次引用只算一篇。

将正文引文放在所支持句子的末尾、句末标点之前，例如“不同脱层工艺需在相同条件下比较[2,4]。”。同一句的不同分句各有独立来源时，先按论断拆句，再各自标注。旧稿的引文位置需结合原文判断，格式化脚本按文本中已有的引用位置处理上标和跳转。

相邻引文 `[1][3]` 显示为 `[1,3]`，连续三条及以上压缩为区间，如 `[1,2,3]` 显示为 `[1-3]`。显示的编号保留跳转，区间内编号也纳入缺引检查。参考文献列表省略 DOI 和 DOI 跳转。英文作者用常规大小写，例如 `MARTÍNEZ M` 写成 `Martínez M`。核对用的 DOI 可来自原始文献或输入数据。

证据线索：

```json
{"reference":1,"claim":"具体论断","source":"出版页URL或全文位置","locator":"页码/表号，或明确仅摘要可见","scope":"对象、工况、指标口径与局限"}
```

`record_status: recorded`只说明线索字段登记完整，`source_authenticity`和`claim_support`仍是pending。没有全文不编页码。脚本不自动联网验证或上传全文，演示条目不能作为真实来源。

## 输出与状态说明

| 文件 | 用途 |
|---|---|
| `report.docx` | 当前流程的最终工作副本，是否可提交仍看待核事项 |
| `report.pdf` | Word导出的PDF |
| `checks.json` | 格式、文献对应、链接、分页及JSON稿公式数量检查 |
| `evidence.json` | 文献、部分引用上下文与来源线索 |
| `run.json` | 输入/范本/规则哈希、轮次、页数、错误和输出校验值 |
| `resolved-input.json` | JSON模式的解析输入，含图片路径和续表边界 |
| `pass-N*` | 每轮起稿、Word更新稿、PDF、分页测量 |
| `pages/page-*.png` | `--render`生成的逐页图片 |

输出可能包含身份、绝对路径和来源位置，分享前按需要检查。失败时尽量保存run.json与已生成文件；目录建立前的参数错误可能只有命令行提示。修复后换新目录重跑，不能把旧目录文件当本次成功输出。

| 状态 | 含义 |
|---|---|
| `machine_status: pass` | 当前机器检查未发现失败，不能证明科学内容正确或页面已查看 |
| `machine_status: fail` | 检查失败或运行错误，查checks.json与run.json.errors |
| `pending` | 未执行或待人工判断，例如未启用自动续表时的分页状态 |
| `ready_for_submission: false` | 脚本始终保守保留，不能改代码让所有报告自动可提交 |

退出0表示机器未发现失败，退出1表示失败。目录未更新的JSON起稿预期失败。PNG存在不等于已逐页检查；字段更新、来源核实、论断支持和视觉验收分别记录。

## 单独运行工具

```powershell
# 起稿，目录缓存未更新
python scripts/build_report.py my-report.json draft.docx

# 只读格式检查，自定义样式可追加 --role-map roles.json
python scripts/audit_docx.py draft.docx --output format-audit.json

# 只读引用对应检查
python scripts/audit_citations.py draft.docx --min-references 10 --output citations.json

# 引文和已映射表线修复，另存
python scripts/scholarly_format.py existing.docx links-fixed.docx --role-map roles.json

# 按角色整理格式，另存
python scripts/normalize_report.py existing.docx normalized.docx --role-map roles.json

# 本机Word更新及PDF，输出父目录须已存在
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/refresh_fields.ps1 -InputDocx draft.docx -OutputDocx updated.docx -OutputPdf updated.pdf

# 工具测试
python -m unittest discover -s tests -v
```

已有输出不覆盖。直接调用build_report.py时，相对图片路径相对当前工作目录；统一入口相对JSON所在目录。

## 常见问题

| 现象 | 处理 |
|---|---|
| `No module named docx/lxml` | 给实际运行脚本的同一个Python安装requirements，检查解释器是否选错 |
| 拒绝覆盖 | 换新目录/文件名，保留旧结果供比较 |
| `目录待更新` | 运行Word字段更新或文档工具的等效流程，再复查 |
| COM启动失败或Access denied | 检查Word安装、许可和运行环境权限，保留草稿与错误，不声称已导出PDF |
| 找不到pdftoppm | 配置Poppler或使用受控替代渲染；已有PDF不代表PNG检查完成 |
| blank模式报错 | 五项身份全部显式传空字符串；普通模式填写真实身份 |
| 正文起点无法识别 | 检查首个编号章的样式和H1映射，不能把所有正文映射成H1 |
| 文献区无法识别 | 使用独立参考文献标题与[n]条目，检查区域位置和样式 |
| `Uncited bibliography` | 根据内容补真实引用或删无关条目，然后重新编号 |
| 引文缺对应条目 | 核对正文与文后编号，不通过虚构条目消除错误 |
| 自动续表未稳定或单行过高 | 合理重组超高行、处理版式，不删数据或缩字号冒充通过 |
| LaTeX显示为普通字符串 | 公式 `math` 字段按原文显示。用 `latex` 字段和 Pandoc 转换，或提供结构节点或正确 OMML |
| 标题前小黑方点 | 先查看是否非打印分页标记及PDF，不删章前分页/与下段同页设置 |
| 机器通过仍待核 | 阅读来源、查看全部页面并独立记录真实复核结果 |

## 交付前核验

1. 核对真实封面信息或用户明确选择的空身份；移除演示研究内容。
2. 阅读来源，核对题名、作者、DOI、页码、工况与论断，区分文献、实测和计算。
3. 处理格式、引文、链接失败项及未知样式，修复目录错误和占位。
4. 在Word更新目录、页码及相关交叉引用，确认正文从1开始并连续。
5. 查看最终全部页面：封面溢出、目录、图表题、每个续页、公式、页眉页脚。
6. 改内容后重新核对受影响页与连锁分页，交付DOCX、可用PDF、证据记录和待核说明。

补充阅读：[完整输入规范](input-schema.md)、[本地Word工作流](word-local-workflow.md)、[格式规则与依据](format-requirements.md)、[统一入口速查](pipeline.md)。
