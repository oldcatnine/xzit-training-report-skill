# 起稿脚本输入

JSON编码UTF-8。所有封面信息和摘要等均必须有非空值；未知信息显式写“待补：…”，生成后只可称草稿。这里的JSON是结构说明，不是可提交的真实报告。

```json
{
  "title_zh":"实际题目", "title_en":"Actual English Title",
  "student":"实际姓名", "student_id":"实际学号", "class":"实际班级",
  "supervisor":"实际教师", "major":"实际专业", "date":"2026年11月13日",
  "abstract_zh":"与正文一致的中文摘要", "abstract_en":"English abstract.",
  "keywords_zh":["光伏组件","回收"], "keywords_en":["photovoltaic modules","recycling"],
  "blocks":[
    {"type":"heading","level":1,"text":"1 绪论"},
    {"type":"paragraph","text":"有真实来源的正文[1]。"},
    {"type":"heading","level":2,"text":"1.1 研究范围"},
    {"type":"table","caption":"表1-1 对比表","rows":[["方法","条件"],["方法名","实际条件"]]},
    {"type":"figure","path":"/absolute/path/figure.png","width_inches":4.5,"caption":"图1-1 实际图题"}
  ],
  "conclusion":"证据支持的结论", "acknowledgements":"真实帮助情况",
  "references":["已核实的作者. 原题[J]. 刊名, 年, 卷(期): 页码. DOI."],
  "appendices":[]
}
```

`figure.path`必须预先存在：直接调用`build_report.py`时相对当前工作目录；统一入口相对输入JSON所在目录。允许标题1—4级；`text`需包含要显示的编号。图表题由调用者按章连续编号，脚本不猜测科研内容。

`conclusion`和`acknowledgements`均可省略、传null或空字符串；省略时不生成该章节。附录仍为可选。不能把工具默认章节冒称为学校硬性要求。

## 可编辑公式

支持原生OMML，不把公式转成图片。简单公式用`math`字符串；结构化公式支持字符串列表与fraction、subscript、superscript、root节点。字符串是显示文本，不解释LaTeX，不计算表达式。例：

```json
{"type":"equation","number":"2-1","math":["η=",{"kind":"fraction","num":{"kind":"subscript","e":"P","sub":"out"},"den":{"kind":"subscript","e":"P","sub":"in"}}],"definitions":[{"symbol":"η","text":"效率，需说明单位与计算口径"}]}
```

`number`使用不带括号的章号-序号，脚本添加括号并设置居中/右端制表位。定义段落设置悬挂缩进；复杂或混合宽度符号的转行须查看渲染页并微调。结构化root节点需要`deg`和`e`。高级公式可用`omml`字段传单个`m:oMath` XML对象（含命名空间）；与`math`二选一。拒绝DTD、实体及非公式文档内容。不能同时提供线性`text`来假装已生成可编辑结构。复杂矩阵、积分等通过经核验OMML提供。

## 跨页表

表格至少包含一行表头和一行数据。可设置`break_before_rows:[4,8]`：rows数组索引0为表头，在原始第4、8索引行之前拆表；边界必须递增、不重复且在2至len(rows)-1之间。数据保留原顺序，每段重复表头，续段在新页右上方生成“续表1-1”。默认从`caption`开头的“表1-1”提取编号，非标准题注须显式提供`continuation_label`。每段均为完整1.5磅三线表。

统一入口`--word --auto-continue`可按Word实际行页码自动建立边界，最多5轮，允许`--max-passes 1..8`调整；超过次数、单行跨页或孤立表头会记录失败并交人工处理，不压缩字号或删除数据。该功能用于JSON生成报告，现有DOCX的长表保留人工处理。拆分后修改内容、字体或环境须重新生成、分页和核验。

## 文献证据记录

可选`evidence`数组：`{"reference":1,"claim":"正文论断","source":"出版页或全文位置","locator":"页码、表号或摘要范围","scope":"对象、工况与局限"}`。统一入口将其与正文引用汇总为独立`evidence.json`。这只是证据线索登记，不自动证明文献真实或支持该论断；记录完整和事实核实分开报告，不能伪造页码或全文阅读状态。

输出保留未更新的目录字段，因此必须在Word/LibreOffice更新。错误输入/已有输出失败退出，不能覆盖原件。

用户明确选择自己填写个人信息时，可添加`"identity_mode":"blank"`并将student/student_id/class/supervisor/major五项均显式设为`""`。不省略字段，保留原封面横线。普通模式仍要求非空。表格可选`"widths_cm":[3,8.5,5]`，列数须匹配且总宽≤16.5cm；短表保持同页，长表需继续处理续表。

默认以现代DOCX兼容级别15输出，避免源DOC的级别11导致可编辑公式在用户保存时被转换为图片。只有用户明确要求特定旧版时才设置`compatibility_mode`，并另行验证公式、链接及格式。
