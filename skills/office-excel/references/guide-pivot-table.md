# 数据透视表增强指引

本文档指导如何为 Excel 任务添加数据透视表/透视图，使用 pandas + openpyxl 实现。透视表是一种跨场景的增强能力——在完成主任务后叠加，而非替代原有分析。

## 目录

- [激活条件](#激活条件)
- [字段规划](#字段规划)
- [实现工作流](#实现工作流)
- [格式规范](#格式规范)
- [图表选择与生成](#图表选择与生成)
- [强制质量验证](#强制质量验证)
- [Canvas 预览交付](#canvas-预览交付)
- [常见问题](#常见问题)

---

## 激活条件

满足以下**任一**条件即激活透视表功能：

### 显式触发

用户明确提到以下关键词：
- "透视表"、"数据透视表"、"pivot table"
- "用图表展示汇总"、"交叉分析"、"多维汇总"
- "按XX分类统计并做图表"、"分组对比图"

### 隐式触发

模型在数据预检（2.2）后判断数据**同时**满足以下条件时，主动向用户建议使用透视表展示：

1. **数据规模**：行数 >= 50
2. **字段结构**：至少 1 个分类字段（文本/类别型）+ 1 个数值字段
3. **任务性质**：涉及分组汇总、对比分析、占比分析、趋势对比

隐式触发时，先简要说明"当前数据适合用透视表展示，可以从 XX 维度汇总 YY 指标"，征得用户同意或用户未反对后再执行。

---

## 字段规划

在生成透视表前，先规划四类字段。这一步决定了透视表的结构和分析维度。

| 字段角色 | 说明 | 选择原则 | 示例 |
|---------|------|---------|------|
| **行字段** (rows) | 主分类维度 | 唯一值数量适中（5~50），是分析的主要分组依据 | 产品名称、部门、地区 |
| **列字段** (columns) | 次分类维度 | 唯一值**较少**（<= 10），用于交叉对比 | 季度、年份、评级等级 |
| **值字段** (values) | 被聚合的数值 | 必须是数值类型，搭配聚合函数 | 销售额:sum、订单数:count |
| **筛选字段** (filters) | 可选过滤条件 | 用户可能想按某维度切片查看 | 年份、产品线 |

### 聚合函数

| 函数 | 用途 | 适用场景 |
|------|------|---------|
| `sum` | 求和 | 金额、数量等累计指标 |
| `count` | 计数 | 记录条数、出现次数 |
| `mean` | 平均值 | 均价、平均评分 |
| `max` / `min` | 最大/最小值 | 极值分析 |
| `median` | 中位数 | 薪资、价格等偏态分布 |

### 字段规划示例

```
数据：销售明细表（5000 行）
字段：日期、地区、产品类别、销售员、销售额、数量

透视规划：
  行字段：产品类别
  列字段：地区
  值字段：销售额:sum, 数量:sum
  筛选字段：（无，或按年份筛选）
```

---

## 实现工作流

严格按以下步骤执行，每步标注使用的超能工具。

### 步骤 1：预检确认数据适合度

```
[工具: Read / Bash]
运行 inspect_workbook.py 或 Read 预览数据，确认：
- 存在分类字段和数值字段
- 数据无严重质量问题（大量空值、格式混乱）
- 无合并单元格（或已做 ffill 处理）
```

### 步骤 2：规划字段映射

根据用户需求和数据结构，确定行/列/值/筛选字段（参见上方"字段规划"）。

### 步骤 3：用 pandas 计算透视数据

```python
import pandas as pd

df = pd.read_excel('input.xlsx', sheet_name='Sheet1')

# 排除汇总行（如有）
df = df[~df['字段名'].str.contains('合计|总计|小计', na=False)]

pivot = pd.pivot_table(
    df,
    index=['产品类别'],        # 行字段
    columns=['地区'],          # 列字段（可选，无则省略）
    values=['销售额', '数量'], # 值字段
    aggfunc={'销售额': 'sum', '数量': 'sum'},
    margins=True,              # 添加汇总行/列
    margins_name='合计'
)

# 多层列索引展平（便于写入 Excel）
pivot.columns = [f'{v}_{c}' if c != '' else v for v, c in pivot.columns]
pivot = pivot.reset_index()
```

### 步骤 4：用 openpyxl 写入并格式化

```python
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = load_workbook('input.xlsx')
ws = wb.create_sheet('透视表')

# 写入数据
for r_idx, row in enumerate(dataframe_to_rows(pivot, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws.cell(row=r_idx, column=c_idx, value=value)

# 应用格式（参见下方"格式规范"）
apply_pivot_style(ws, header_row=1, data_rows=len(pivot), data_cols=len(pivot.columns))

wb.save('output.xlsx')
```

### 步骤 5：生成配套图表

```python
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

chart = BarChart()
chart.title = "各产品类别销售额对比"
chart.y_axis.title = "销售额"
chart.x_axis.title = "产品类别"

data = Reference(ws, min_col=2, min_row=1, max_col=5, max_row=len(pivot)+1)
cats = Reference(ws, min_col=1, min_row=2, max_row=len(pivot)+1)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
ws.add_chart(chart, "A" + str(len(pivot) + 4))

wb.save('output.xlsx')
```

### 步骤 6：执行强制质量验证

参见下方[强制质量验证](#强制质量验证)章节，五项检查全部通过后才可交付。

### 步骤 7：视觉校验

```
[工具: open_url_in_browser + take_screenshot]
用 LibreOffice 打开生成的 Excel 文件，截图确认：
- 透视表数据完整、格式清晰
- 图表非空壳、标题/坐标轴/图例齐全
- 无重叠、溢出、乱码
```

### 步骤 8：Canvas 预览交付

```
[工具: CanvasCreateFile (type=webview)]
创建 HTML 预览页面，包含：
- 透视表数据的 HTML 表格
- 图表截图（通过 FileBatchUpload 上传后嵌入）
- 数据摘要信息
```

### 步骤 9：文件交付

```
[工具: NotifyHuman]
将最终 Excel 文件推送给用户。
```

---

## 格式规范

### Monochrome 样式（默认，通用分析场景）

| 元素 | 格式 |
|------|------|
| 表头行 | 背景 `#333333` + 白色粗体字，行高 30 |
| 行标签列 | 背景 `#F5F5F5` + 黑色粗体字 |
| 数据区 | 白底，隔行浅灰 `#FAFAFA` |
| 汇总行/列 | 背景 `#E8E8E8` + 粗体 |
| 数值格式 | 右对齐，千分位分隔符，保留 2 位小数 |
| 百分比 | `0.0%` 格式 |

### Finance 样式（财务/营收场景）

| 元素 | 格式 |
|------|------|
| 表头行 | 背景 `#1F4E79` + 白色粗体字 |
| 行标签列 | 背景 `#D6E3F0` |
| 数据区 | 白底，隔行浅蓝 `#EDF2F7` |
| 汇总行/列 | 背景 `#BDD7EE` + 粗体 |
| 货币格式 | `¥#,##0`（或 `$#,##0`），表头注明单位 |
| 负数 | 括号表示 `(123)` + 红色字体 |

### 通用格式规则

效果目标：**表格整体紧凑不拥挤，列间有呼吸空间，长短内容各得其所。**

- 隐藏网格线：`ws.sheet_view.showGridLines = False`
- 内容从 `B2` 开始（非 `A1`），留白提升可读性
- 冻结窗格：冻结表头行和行标签列

#### 列宽自适应规则

| 规则 | 说明 |
|------|------|
| 计算方式 | 扫描该列所有单元格，取最长内容的显示宽度 + padding |
| 中文字符 | 每个中文字符按 2 个字符宽度计算 |
| padding | 列宽 = 最长内容宽度 + 3（留出呼吸空间） |
| 下限 | 8 字符（避免过窄导致数值截断） |
| 上限 | 40 字符（防止单列过宽挤压其他列） |
| 超长文本 | 内容超过上限时启用 `wrap_text=True` 自动换行 |

#### 行高规则

| 行类型 | 行高 | 说明 |
|--------|------|------|
| 表头行 | 30 | 粗体标题需要更多垂直空间 |
| 普通数据行 | 20 | 紧凑但不贴字，数值阅读舒适 |
| 汇总行 | 26 | 略高于数据行，视觉区分 |
| 含换行数据行 | 36 | 当该行存在 wrap_text 单元格时自动增高 |

#### 内边距

通过列宽 padding（+3 字符）实现左右呼吸空间，数据列右对齐时自然留出右侧间距。行标签列首字符前通过 `Alignment(indent=1)` 增加 1 字符缩进。

```python
from openpyxl.utils import get_column_letter


def auto_fit_columns(ws, header_row, data_rows, data_cols, min_width=8, max_width=40, padding=3):
    """根据内容自动调整列宽，超长文本启用换行"""
    for col in range(1, data_cols + 1):
        max_len = 0
        for row in range(header_row, header_row + data_rows + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                cell_len = sum(2 if ord(c) > 127 else 1 for c in str(val))
                max_len = max(max_len, cell_len)
        width = min(max(max_len + padding, min_width), max_width)
        ws.column_dimensions[get_column_letter(col)].width = width
        # 超长内容列启用自动换行
        if max_len + padding > max_width:
            for row in range(header_row, header_row + data_rows + 1):
                cell = ws.cell(row=row, column=col)
                cell.alignment = Alignment(
                    horizontal=cell.alignment.horizontal or 'left',
                    vertical='center',
                    wrap_text=True,
                )


def set_row_heights(ws, header_row, data_rows, data_cols, has_summary_row=True):
    """按行类型设置行高，含换行单元格的行自动增高"""
    ws.row_dimensions[header_row].height = 30
    last_data_row = header_row + data_rows
    for row in range(header_row + 1, last_data_row + 1):
        is_summary = has_summary_row and (row == last_data_row)
        base_height = 26 if is_summary else 20
        # 检查该行是否有换行单元格
        has_wrap = any(
            ws.cell(row=row, column=c).alignment and ws.cell(row=row, column=c).alignment.wrap_text
            for c in range(1, data_cols + 1)
        )
        ws.row_dimensions[row].height = 36 if has_wrap else base_height


def apply_pivot_style(ws, header_row, data_rows, data_cols, style='monochrome'):
    """透视表统一格式化：颜色/字体/对齐 + 列宽自适应 + 行高分层"""
    header_fill = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    alt_fill = PatternFill(start_color='FAFAFA', end_color='FAFAFA', fill_type='solid')
    summary_fill = PatternFill(start_color='E8E8E8', end_color='E8E8E8', fill_type='solid')
    thin_border = Border(
        bottom=Side(style='thin', color='DDDDDD')
    )

    if style == 'finance':
        header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        alt_fill = PatternFill(start_color='EDF2F7', end_color='EDF2F7', fill_type='solid')
        summary_fill = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')

    # 表头
    for col in range(1, data_cols + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # 数据行（隔行变色 + 行标签缩进）
    last_data_row = header_row + data_rows
    for row in range(header_row + 1, last_data_row + 1):
        is_summary = (row == last_data_row)
        for col in range(1, data_cols + 1):
            cell = ws.cell(row=row, column=col)
            if is_summary:
                cell.fill = summary_fill
                cell.font = Font(bold=True)
            elif (row - header_row) % 2 == 0:
                cell.fill = alt_fill
            cell.border = thin_border
            if col == 1:
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif col > 1:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0.00'

    # 列宽自适应 + 行高分层
    auto_fit_columns(ws, header_row, data_rows, data_cols)
    set_row_heights(ws, header_row, data_rows, data_cols)

    ws.sheet_view.showGridLines = False
```

---

## 图表选择与生成

根据数据特征和分析目的选择图表类型：

| 分析目的 | 推荐图表 | openpyxl 类 | 适用条件 |
|---------|---------|------------|---------|
| 并排比较类别 | 簇状柱形图 | `BarChart` | 分类 <= 15，值字段 1~3 个 |
| 时间/序列趋势 | 折线图 | `LineChart` | 行字段为时间/有序序列 |
| 整体占比 | 饼图 | `PieChart` | 分类 <= 6，单值字段 |
| 堆叠对比 | 堆叠柱形图 | `BarChart(grouping='stacked')` | 需要看各部分占总量的比例 |

### 图表质量标准

- 必须有中文标题（`chart.title`）
- 必须有坐标轴标签（`chart.x_axis.title` / `chart.y_axis.title`）
- 图例位置清晰不遮挡数据
- 图表放在数据区域下方，间隔至少 2 行
- 图表尺寸合理：宽度 15~20 cm，高度 10~12 cm

---

## 强制质量验证

添加透视表后，**必须**通过以下五项验证，任一失败则需修复或回退。

### 验证 1：原始数据完整性

```
[工具: Bash]
用 Python 对比添加透视 Sheet 前后：
- 原数据 Sheet 行数不变
- 原数据 Sheet 列数不变
- 抽查 5 个关键单元格值一致
```

```python
from openpyxl import load_workbook
wb = load_workbook('output.xlsx', data_only=True)
ws = wb['原始数据Sheet名']
assert ws.max_row == expected_rows, f"行数变化: {ws.max_row} != {expected_rows}"
assert ws.max_column == expected_cols, f"列数变化: {ws.max_column} != {expected_cols}"
```

### 验证 2：透视数据准确性

```
[工具: Bash]
用 pandas 独立重新计算聚合结果，与透视 Sheet 中的值逐一比对：
```

```python
import pandas as pd
from openpyxl import load_workbook

df = pd.read_excel('output.xlsx', sheet_name='原始数据Sheet名')
expected = df.groupby('分类字段')['值字段'].sum()

wb = load_workbook('output.xlsx', data_only=True)
ws = wb['透视表']
for idx, (category, exp_val) in enumerate(expected.items()):
    actual = ws.cell(row=idx+2, column=2).value
    assert abs(actual - exp_val) < 0.01, f"{category}: {actual} != {exp_val}"
```

### 验证 3：图表有效性

```
[工具: Bash]
用 openpyxl 加载确认图表对象有真实数据引用：
```

```python
wb = load_workbook('output.xlsx')
ws = wb['透视表']
assert len(ws._charts) > 0, "透视表 Sheet 中无图表"
for chart in ws._charts:
    assert len(chart.series) > 0, "图表无数据系列"
    assert chart.title is not None, "图表缺少标题"
```

```
[工具: open_url_in_browser + take_screenshot]
用 LibreOffice 打开 Excel，截图视觉确认图表非空壳。
```

### 验证 4：公式重算

```
[工具: Bash]
python scripts/formula_verify.py output.xlsx
确认 total_errors == 0
```

### 验证 5：文件可用性

```
[工具: Bash]
用 openpyxl 重新加载文件，确认无异常：
```

```python
try:
    wb = load_workbook('output.xlsx')
    sheet_names = wb.sheetnames
    assert '透视表' in sheet_names, "透视表 Sheet 不存在"
    print(f"文件正常，包含 {len(sheet_names)} 个 Sheet: {sheet_names}")
except Exception as e:
    raise RuntimeError(f"文件无法正常打开: {e}")
```

### 验证失败的回退策略

如果任一验证未通过且无法快速修复：

1. 删除透视表 Sheet，恢复原始文件
2. 用 `NotifyHuman` 告知用户回退原因
3. 将透视数据以 Canvas (markdown) 形式展示给用户作为替代

---

## Canvas 预览交付

透视表生成后，用 `CanvasCreateFile` 创建 webview 预览，让用户无需下载即可确认结果：

```html
<!-- canvas 模板结构 -->
<html>
<head><style>
  table { border-collapse: collapse; width: 100%; }
  th { background: #333; color: #fff; padding: 8px 12px; }
  td { padding: 6px 12px; border-bottom: 1px solid #eee; }
  tr:nth-child(even) { background: #fafafa; }
  .summary { margin: 16px 0; padding: 12px; background: #f5f5f5; border-radius: 4px; }
</style></head>
<body>
  <h2>数据透视表预览</h2>
  <div class="summary">
    <p>数据源: XX.xlsx | 行数: N | 透视维度: 产品类别 × 地区 | 值: 销售额(sum)</p>
  </div>
  <table>
    <!-- 透视表数据渲染为 HTML 表格 -->
  </table>
  <h3>图表预览</h3>
  <img src="截图URL" alt="透视图" style="max-width:100%;">
</body>
</html>
```

**流程**：
1. `Bash` 将 pivot DataFrame 转为 HTML 表格字符串
2. `FileBatchUpload` 上传图表截图获得 URL
3. `CanvasCreateFile` 创建 webview canvas，嵌入表格 + 截图
4. Canvas 自动上屏展示给用户（无需 NotifyHuman）

---

## 常见问题

### 合并单元格处理

源数据含合并单元格时，合并区域只有左上角有值，其余为 None。读取后需要 ffill：

```python
df = pd.read_excel('input.xlsx')
df.fillna(method='ffill', inplace=True)
```

### 空值处理

- `pivot_table` 默认忽略 NaN（`dropna=True`）
- 若需要统计空值数量，使用 `aggfunc='count'`（count 不含 NaN）vs `aggfunc='size'`（size 含 NaN）
- 透视结果中的 NaN 在写入 Excel 前替换为 0 或空字符串：`pivot.fillna(0)`

### 大数据量处理

数据超过 10000 行时：
- 考虑先按关键维度预聚合，再做透视
- 图表数据点过多（>50 个类别）时，只展示 Top N + "其他"
- 使用 `Bash` 分批处理，避免内存溢出

### 透视表 vs 普通公式汇总

| 场景 | 用透视表 | 用公式 |
|------|---------|-------|
| 多维分组（行+列交叉） | 适合 | 不适合（需要大量 SUMIFS） |
| 数据量 50+ 行 | 适合 | 可以但维护困难 |
| 需要配套图表 | 适合（透视表+图表一体） | 需要单独建图 |
| 简单单列求和 | 过度 | 适合（一个 SUM 即可） |
| 用户要求"就地计算" | 不适合（需新建 Sheet） | 适合（在原 Sheet 写公式） |
