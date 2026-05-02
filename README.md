# DHU Template Injector

基于“清洗母版 + 锚点定位 + 内容直填”的东华大学本科论文填充方案实现。

这个目录不再沿用旧的“Markdown 先转普通 Word、再硬修格式”的路线，而是拆成两步：

1. 先从原始东华模板生成一份可复用的干净母版。
2. 再把 Markdown 内容和封面元数据写回母版。

## 目录说明

| 文件 | 作用 |
| --- | --- |
| `prepare_clean_template.py` | 从 `本科毕业论文参考模版（2025版）.docx` 生成干净母版 |
| `optimize_clean_template.py` | 按旧 formatter 规则校正现有干净模板的占位格式 |
| `fill_template.py` | 将 Markdown 与元数据直填进母版 |
| `run_dhu_pipeline.py` | 一键执行“优化模板 + 填充文档” |
| `template_locator.py` | 锚点定位工具 |
| `cover_field_locator.py` | 封面表格与标题区定位工具 |
| `markdown_parser.py` | Markdown 结构解析器 |
| `inspect_template.py` | 输出模板锚点、节数、段落数等摘要 |
| `dhu_template_schema.json` | 锚点、封面字段和占位内容配置 |
| `legacy_format_rules.py` | 从旧 formatter 收敛出的东华段落格式规则 |
| `docs/` | 东华模板直填方案与复用评估文档 |
| `outputs/` | 干净母版和生成结果输出目录 |

## 依赖

```bash
pip install -r requirements.txt
```

## 当前推荐流程

如果你已经有一份手工清理过提示气泡的干净模板，推荐只走这两步：

```bash
cd DHU-Template-Injector
python optimize_clean_template.py
python fill_template.py ../06_论文Markdown草稿/毕业论文初稿.md
```

说明：

- `optimize_clean_template.py` 只会修正当前干净模板里的占位段落格式，不会回写原始母体模板。
- 标题、正文、关键词、参考文献等格式规则已经收敛到 `legacy_format_rules.py`。

如果你想直接一条命令跑完整流程，也可以：

```bash
cd DHU-Template-Injector
python run_dhu_pipeline.py ../06_论文Markdown草稿/毕业论文初稿.md
```

## 1. 首次生成干净母版

```bash
cd DHU-Template-Injector
python prepare_clean_template.py
```

只有在你还没有干净模板时，才需要这一步。默认会读取上一级目录的：

- `本科毕业论文参考模版（2025版）.docx`

并输出到：

- `outputs/dhu_undergrad_clean_template.docx`

这个“干净母版”保留了：

- 封面结构
- 诚信声明 / 授权书
- 摘要、目录、正文、参考文献、附录、致谢等大区块
- section、页边距、页眉页脚等模板结构

同时会清空或替换掉：

- 示例摘要
- 示例正文
- 示例参考文献
- 封面示例标题

说明：

- 为了后续脚本继续稳定定位，正文区会保留少量“样式占位段落”，例如一级标题、二级标题、三级标题和正文段落占位。这些占位内容会在正式填充时被覆盖。
- 当前这一步是“半自动清洗”。如果后面你手工进一步清理模板里的特殊对象、批注、域或图片，建议继续基于 `outputs/dhu_undergrad_clean_template.docx` 演进。

## 2. 用 Markdown 直填母版

```bash
cd DHU-Template-Injector
python fill_template.py ../06_论文Markdown草稿/毕业论文初稿.md
```

默认使用：

- `outputs/dhu_undergrad_clean_template.docx`

默认输出到：

- `outputs/毕业论文初稿_东华直填.docx`

### 可选：补充封面和摘要元数据

如果 Markdown 里没有英文题目、摘要、关键词、封面字段，可以再给一个 JSON：

```json
{
  "title": "可降解类玻璃高分子材料结构设计与性能研究",
  "english_title": "Structure Design and Properties of Degradable Vitrimer Materials",
  "cover": {
    "college": "材料科学与工程学院",
    "major": "高分子材料与工程",
    "author": "冯泽宇",
    "student_id": "221040116",
    "supervisor": "XXX",
    "submit_date": "2026年4月"
  },
  "cn_abstract": [
    "这里填写中文摘要。"
  ],
  "cn_keywords": [
    "类玻璃高分子",
    "PDK",
    "PCL"
  ],
  "en_abstract": [
    "English abstract goes here."
  ],
  "en_keywords": [
    "vitrimer",
    "PDK",
    "PCL"
  ]
}
```

然后执行：

```bash
python fill_template.py ../06_论文Markdown草稿/毕业论文初稿.md -m metadata.json
```

## 当前已实现范围

- 封面表格字段写回
- 封面中文/英文题目写回
- 中文摘要 / 关键词
- 英文摘要 / 关键词
- 正文一级、二级、三级标题编号重建
- 正文段落写回
- 正文引用 `[1]` / `[1,4]` 自动转上标
- Markdown 图片写回
- Markdown 表格写回
- 参考文献段落写回
- 致谢 / 附录 / 外文原文及译文基础写回

## 当前限制

- 目录仍建议在 Word / WPS 打开后手动更新
- 公式对象仍未自动重建
- 默认不自动生成“续表”；精确续表依赖 Word / WPS 完成真实分页后的结果
- 当前锚点配置基于现有这份东华模板整理，如果后面你换了模板版本，优先改 `dhu_template_schema.json`

## Markdown 输入规则

完整规则见：

- `docs/Markdown书写规则.md`

最常用的写法如下：

```md
# 论文总题目

# 一、绪论

## （一）研究背景

正文引用示例[1]。

表题：测试组分数据
| 组分 | Hf(kcal/mol) | Sf(kcal/mol) | Cp(kcal/mol) |
| --- | --- | --- | --- |
| A1 | 100 | 100 | 100 |

图题：示例图片展示效果
![](./figures/demo.png)

# 参考文献

[1] ...
```

## 推荐后续迭代

1. 在这套目录里继续新增 `template_style_inspector.py`，把样式画像单独导出。
2. 针对图片、表格、公式补一层块级写回。
3. 生成后接一个轻量质检脚本，检查锚点、section 数和页眉是否异常。
