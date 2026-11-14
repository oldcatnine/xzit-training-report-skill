# 统一生成、修复与检查入口

需要完整操作步骤、演示输入、全部参数与排错时，阅读[详细使用文档](usage-guide.md)。本页保留入口速查。

依赖Python 3.10+、python-docx和lxml。Word分页/导出要求Windows及已安装Microsoft Word；渲染PNG还需PATH中的Poppler `pdftoppm`。不自动安装依赖或上传内容。

## 生成

```sh
python scripts/report_pipeline.py input.json new-output --word --auto-continue --render --min-references 10
```

最低篇数仅是本次命令参数，按任务替换。未加`--word`时也可生成与检查，但目录尚未更新，因此检查会报告“目录待更新”，不能视为终稿。Word处理通过单个子进程的ExecutionPolicy Bypass运行本技能脚本，不修改机器或用户的持久执行策略。Word COM不可用或访问被拒绝时记录错误，不能冒称已完成导出。

输出目录必须全新。保留`report.docx`、有Word时的`report.pdf`、`checks.json`、`evidence.json`、`run.json`及各轮中间文件；`--render`额外生成pages目录。`resolved-input.json`保留解析后的图片路径与续表边界。运行记录包含输入、范本和规则哈希、执行轮数、字段更新时间状态、文件校验值和未核事项。失败时尽量保存运行记录；修复后使用新输出目录重试。

退出0表示机器未发现失败，退出1表示失败。始终保留`ready_for_submission=false`与来源、论断、逐页视觉待核状态。PNG生成不等于模型或人工已经看过；另记录实际复核结果。文献真实性、引用论断及公式变量物理意义由阅读原始来源确认。

## 整理现有DOCX

```sh
python scripts/report_pipeline.py existing.docx new-output --role-map roles.json --normalize --word --render
```

`roles.json`是原样式名到角色名的映射，例如：

```json
{"标题 1":"H1","正文":"Body","文献条目":"Reference","表内文字":"Table"}
```

支持角色见`report_format.py`的ROLES。XZIT样式和Word内置Heading/标题1—4可直接识别；参考文献区域按独立的“参考文献/References/Bibliography”标题及[n]编号条目识别；正文从第一个编号章标题开始，不限定必须为“1 空格”。自定义章样式应显式映射到H1。无法识别正文或文献区直接失败，不把零条检查当作通过。

`--normalize`只对已识别或已映射段落应用范本样式，清理冲突字体、字号、行距和缩进，保留加粗、斜体、上下标、公式、图片、域和链接对象；未知段落列出待核。公式定义保留原悬挂缩进。它不会重建旧稿封面、改页边距或猜测正文角色；这些仍须按检查结果在副本中处理。不加`--normalize`只修正文献链接、上标与已映射正文表格，不全篇重排。

也可分别使用`normalize_report.py`、`scholarly_format.py`、`audit_citations.py`的`--role-map`。原件和已有输出均不覆盖。未知样式、未映射的封面及目录必须人工复核。
