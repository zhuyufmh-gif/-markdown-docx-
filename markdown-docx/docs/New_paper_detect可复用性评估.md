# `New_paper_detect-main` 可复用性评估

## 结论

`New_paper_detect-main` **有不少值得借鉴的部分**，但更适合借用其中的：

- Word 模板结构解析能力
- 锚点定位能力
- 字体/段落/分页信息读取能力
- 规则配置的组织方式

**不适合整套直接搬进我们现在的东华本科论文生成系统。**

原因很简单：

- 它本质上是“格式检测与标注系统”
- 我们现在要做的是“基于干净东华模板的内容填充系统”

所以正确做法不是“把检测器直接改成生成器”，而是：

**提取其中的底层 Word 解析逻辑，服务于新的模板直填方案。**

---

## 1. 这个系统是怎么实现的

它基本可以分成 3 层。

### 1.1 规则层

核心是 `templates/*.json`。

例如：

- `templates/header.json`
- `templates/Content.json`

这里面定义了：

- 检测开关
- 标题/页眉/目录/页边距等规则
- 正则锚点
- 报错文案

也就是说，这个项目不是把规则硬编码死在每个 Python 文件里，而是采用了“**代码 + JSON 规则**”的组织方式。

---

### 1.2 Word 结构解析层

这是最值得借鉴的一层。

典型文件有：

- `paper_detect/other_format_common.py`
- `paper_detect/back_matter_common.py`
- `backend/app/utils/document_parser.py`
- `scripts/extract_word_format.ps1`

它们负责做这些事情：

- 找到某类标题或锚点所在段落
- 推断正文起止范围
- 识别 section 起点
- 读取页码、页眉页脚、页边距
- 识别段落对齐、段前段后、字体、字号、行距
- 从封面表格或段落中提取“学位申请人”“所在学院”等字段

这部分逻辑和我们现在的“模板直填”是高度相关的。

---

### 1.3 执行与报告层

核心入口是：

- `run_all_detections.py`

它做的是：

1. 按 `TEMPLATE_MAPPING` 组织各检测模块
2. 串行/并行调度各模块
3. 汇总错误
4. 输出文本报告和标注文档

配套还有：

- `paper_detect/report_data_converter.py`
- `paper_detect/error_formatter.py`
- `paper_detect/word_report_generator.py`

这一层更偏“检测结果表达”，对我们当前生成系统帮助有限。

---

## 2. 最值得借鉴的部分

## A. `scripts/extract_word_format.ps1`

这是我认为最有价值的一个现成工具。

它会通过 Word COM 打开模板，然后导出：

- 段落样式频次
- 样式详情
- 每节页面设置
- 页眉页脚文本
- 表格信息
- 段落样例及字体/字号/缩进/行距

它本质上是在做“**模板画像**”。

对我们现在的东华本科项目来说，它最适合用于：

1. 对清洗后的东华本科母版做一次格式采样
2. 形成一份机器可读或人工可读的模板说明
3. 帮我们确认哪些段落样式、节设置、页眉页脚是必须保留的

这和我们后面做“模板复制 + 填内容”非常匹配。

---

## B. `paper_detect/other_format_common.py`

这个文件里有一批非常适合抽出来复用的基础函数。

特别值得关注的函数有：

- `load_template`
- `resolve_template_path`
- `find_paragraph_index_by_pattern`
- `detect_body_range`
- `get_section_start_paragraph_indices`
- `get_section_start_pages`
- `find_anchor_pages`
- `find_anchor_pages_bulk`

这些函数的意义是：

- 根据正则或锚点定位内容区域
- 把 section、段落、页码关系建立起来
- 找到“正文开始”“参考文献开始”“声明页结束”等关键边界

对我们的新系统来说，它们可以直接服务于：

- 模板锚点定位
- 内容替换边界识别
- 后续页眉题目或 section 关系校验

---

## C. `paper_detect/back_matter_common.py`

这个文件适合拿来做“**模板样式侦测工具箱**”。

特别有用的函数有：

- `get_effective_rfonts`
- `detect_paragraph_alignment`
- `detect_space_before_after_pt`
- `detect_font_for_run`
- `check_standard_section_title`

这些逻辑的价值不在于“报错”，而在于：

- 可以帮助我们确认模板某个标题段到底继承了什么字体
- 可以判断一个目标段落的实际对齐/间距/字号
- 可以在内容填充前后做轻量校验

以后如果我们做：

- “把一级标题插入到模板后，确认它是否继承了正确样式”
- “识别模板中摘要标题的真实格式”

这些函数都会很有帮助。

---

## D. `backend/app/utils/document_parser.py`

这个文件的可借鉴点不是“提取结果本身”，而是“**按标签找字段**”的思路。

它已经实现了：

- 从封面表格中查找“学位申请人”
- 从封面表格中查找“所在学院”
- 表格找不到时，再从普通段落中找

这个思路特别适合我们后面做：

- 封面字段填写
- 模板字段替换

也就是说，我们可以把它从“抽取信息”改成“定位字段位置并写回信息”。

---

## E. `templates/*.json` 的组织方式

这些 JSON 规则本身不一定能直接复用，因为它们很多针对的是东华研究生格式，不是本科模板。

但是它们的组织方式很值得借鉴。

我们完全可以在自己的新系统里做一个更小的配置层，例如定义：

- 母版路径
- 摘要开始/结束锚点
- 英文摘要开始/结束锚点
- 正文开始/结束锚点
- 参考文献开始锚点
- 致谢/附录锚点
- 封面字段标签
- 需要保留或跳过的 section

这样后续改模板时，不必大量改 Python 逻辑。

---

## 3. 不建议直接复用的部分

以下部分我不建议直接搬到当前东华本科生成系统里：

- `run_all_detections.py`
- `paper_detect/Content_detect.py`
- `paper_detect/Abstract_detect.py`
- `paper_detect/Reference_detect.py`
- `paper_detect/word_report_generator.py`
- `paper_detect/report_data_converter.py`
- `paper_detect/error_formatter.py`
- `backend/` 下的大部分 Web 服务与任务调度逻辑

原因：

- 它们的目标是“找问题”
- 我们的目标是“生成正确文档”
- 里面很多逻辑是围绕报错、容错、兼容不同检测模块而写的
- `Content_detect.py` 这类文件已经明显带有特定论文结构假设，不适合作为生成核心

可以把它们保留为后期“生成后质检”参考，但不应该作为主流程核心。

---

## 4. 对当前项目最实际的借法

我们当前项目里，旧路线主要还是：

- `md_to_dhu_thesis.py`
- `dhu_thesis_formatter.py`

这条路线本质上是：

1. Markdown 先转成普通 Word
2. 再靠 formatter 把它硬修成东华格式

而现在我们已经明确准备转向：

**复制东华本科干净母版，再把内容填进去。**

在这个新方向下，`New_paper_detect-main` 最适合提供 4 类帮助。

### 4.1 做模板分析

借 `extract_word_format.ps1` 的能力，把东华本科母版的：

- 样式
- section
- 页眉页脚
- 页面设置
- 标题格式

先摸清楚。

---

### 4.2 做锚点定位

借 `other_format_common.py` 的锚点/范围函数，帮我们稳定定位：

- 摘要区
- 英文摘要区
- 正文区
- 参考文献区
- 致谢区
- 附录区

---

### 4.3 做字段定位

借 `document_parser.py` 的标签识别思路，定位封面中的：

- 论文题目
- 学生姓名
- 学号
- 学院
- 专业
- 指导教师

---

### 4.4 做生成后轻量质检

等新脚本把内容填进模板之后，可以保留极少量检测逻辑做 sanity check，例如：

- 页眉是否仍存在
- section 数量是否异常变化
- 正文起始锚点是否还在
- 一级标题是否继承了模板样式

这个用途远比整套格式检测更适合我们现在的目标。

---

## 5. 建议的迁移方式

推荐这样借，而不是整仓照搬。

### 第一步

从 `New_paper_detect-main` 中挑出一个新的“公共工具子集”，优先候选：

- `other_format_common.py` 中的锚点/section/page 工具
- `back_matter_common.py` 中的字体/段落格式检测工具
- `document_parser.py` 中的封面字段定位逻辑

---

### 第二步

把这些逻辑改造成面向“模板填充”的辅助模块，例如：

- `template_locator.py`
- `template_style_inspector.py`
- `cover_field_locator.py`

不要继续沿用“detect”命名，否则后面职责会越来越混。

---

### 第三步

在我们自己的项目里新增一个轻量配置文件，比如：

- `dhu_template_schema.json`

里面只描述：

- 锚点
- 封面标签
- 替换区边界
- 必须保留的 section 约束

---

### 第四步

把当前主流程改成：

1. 复制干净东华母版
2. 定位封面字段和正文锚点
3. 清空模板示例内容
4. 把 Markdown 解析结果写入模板
5. 做一次轻量校验
6. 输出最终 `.docx`

---

## 6. 最终判断

最终判断很明确：

- **能借鉴，而且值得借鉴**
- **但借的是底层 Word 解析能力，不是整套检测流程**

如果后续按“模板直填”路线推进，那么 `New_paper_detect-main` 最有价值的身份应该是：

**东华模板解析与定位能力的参考库。**

而不是新的主流程框架。
