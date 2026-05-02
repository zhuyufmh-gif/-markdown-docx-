# Markdown DOCX Converter

把 Markdown 论文稿写入东华大学本科毕业论文 Word 模板，生成带封面、摘要、正文、图表、参考文献等结构的 `.docx` 文件。

## 功能效果

- 使用 `outputs/dhu_undergrad_clean_template.docx` 作为干净 Word 母版
- 自动写入封面中文题目、英文题目和学院、专业、作者、学号、导师、日期等字段
- 支持中文摘要、英文摘要、关键词
- 支持正文一级、二级、三级标题，并按论文格式重建标题编号
- 支持普通正文段落和参考文献段落
- 支持正文引用标记上标，例如 `[1]`、`[1,4]`
- 支持 Markdown 表格写入 Word 表格
- 支持 Markdown 图片写入 Word，并保留图题
- 支持致谢、附录、外文原文及译文等常见毕业论文区块

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `run_dhu_pipeline.py` | 一键生成 Word 文档 |
| `fill_template.py` | 将 Markdown 内容填入 Word 母版 |
| `optimize_clean_template.py` | 优化干净模板中的占位段落格式 |
| `prepare_clean_template.py` | 从原始学校模板生成干净母版 |
| `markdown_parser.py` | 解析 Markdown 标题、段落、图表和参考文献 |
| `docx_utils.py` | Word 文档写入工具函数 |
| `legacy_format_rules.py` | 论文段落、标题、图表等格式规则 |
| `dhu_template_schema.json` | 模板锚点和封面字段配置 |
| `outputs/dhu_undergrad_clean_template.docx` | 默认使用的干净 Word 母版 |
| `docs/核心机制.md` | 项目核心机制说明 |
| `docs/Markdown书写规则.md` | Markdown 输入规范 |

## 安装依赖

建议使用 Python 3.9 或更新版本。

```bash
pip install -r requirements.txt
```

## 快速使用

进入项目目录：

```bash
cd markdown-docx
```

使用默认干净模板生成 Word：

```bash
python run_dhu_pipeline.py path/to/论文.md
```

也可以只执行填充步骤：

```bash
python fill_template.py path/to/论文.md
```

生成结果会输出到 `outputs/` 目录，文件名通常为：

```text
原Markdown文件名_东华直填.docx
```

## 补充封面和摘要信息

如果 Markdown 里没有完整的封面、摘要或关键词信息，可以新建一个 JSON 文件，例如 `metadata.json`：

```json
{
  "title": "可降解类玻璃高分子材料结构设计与性能研究",
  "english_title": "Structure Design and Properties of Degradable Vitrimer Materials",
  "cover": {
    "college": "材料科学与工程学院",
    "major": "高分子材料与工程",
    "author": "姓名",
    "student_id": "学号",
    "supervisor": "导师姓名",
    "submit_date": "2026年4月"
  },
  "cn_abstract": [
    "这里填写中文摘要第一段。",
    "这里填写中文摘要第二段。"
  ],
  "cn_keywords": [
    "关键词一",
    "关键词二",
    "关键词三"
  ],
  "en_abstract": [
    "English abstract paragraph."
  ],
  "en_keywords": [
    "keyword one",
    "keyword two",
    "keyword three"
  ]
}
```

运行：

```bash
python fill_template.py path/to/论文.md -m metadata.json
```

## Markdown 写法示例

```md
# 可降解类玻璃高分子材料结构设计与性能研究

# 一、绪论

## （一）研究背景

这里是正文内容。文献引用可以写成[1]，多个引用可以写成[1,4]。

表题：样品配方
| 样品 | PCL-COOH | PCL-TK | Chitosan |
| --- | ---: | ---: | ---: |
| S1 | 1.00 | 0.50 | 0.20 |
| S2 | 1.00 | 0.75 | 0.20 |

图题：样品显微形貌
![](./figures/micrograph.png)

# 参考文献

[1] 作者. 文献题名[J]. 期刊名, 年份, 卷(期): 页码.

# 致谢

这里填写致谢内容。
```

完整写法见 `docs/Markdown书写规则.md`。

## 常用命令

生成前先优化默认干净模板：

```bash
python optimize_clean_template.py
```

从学校原始模板重新生成干净母版：

```bash
python prepare_clean_template.py
```

查看模板锚点和结构摘要：

```bash
python inspect_template.py
```

## 注意事项

- 默认模板路径为 `outputs/dhu_undergrad_clean_template.docx`
- 目录建议在 Word 或 WPS 中打开后手动更新
- 公式对象暂不自动转换为 Word 公式
- 精确续表依赖 Word 或 WPS 完成分页后的真实版面
- 如果更换学校模板版本，优先检查并调整 `dhu_template_schema.json`
