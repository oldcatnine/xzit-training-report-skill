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

`figure.path`相对当前工作目录或绝对路径；必须预先存在。允许标题1—4级；`text`需包含要显示的编号。图表题由调用者按章连续编号，脚本不猜测科研内容。公式块明确拒绝而不静默丢弃：用文档工具加入可编辑公式，按F20/F21处理。续表标识和真实分页须渲染后补齐。

输出保留未更新的目录字段，因此必须在Word/LibreOffice更新。错误输入/已有输出失败退出，不能覆盖原件。

用户明确选择自己填写个人信息时，可添加`"identity_mode":"blank"`并将student/student_id/class/supervisor/major五项均显式设为`""`。不省略字段，保留原封面横线。普通模式仍要求非空。表格可选`"widths_cm":[3,8.5,5]`，列数须匹配且总宽≤16.5cm；短表保持同页，长表需继续处理续表。

默认以现代DOCX兼容级别15输出，避免源DOC的级别11导致可编辑公式在用户保存时被转换为图片。只有用户明确要求特定旧版时才设置`compatibility_mode`，并另行验证公式、链接及格式。
