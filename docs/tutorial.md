# PPT Automation Demo Tutorial

Step-by-step guide to the `ppt-automation-demo` project: from sample data, to PPT template, to batch-generated reports.

This tutorial is aligned with the current code in:

- `generate_sample_data.py`
- `create_template.py`
- `batch_generate.py`

---

## Part 1: Environment & Installation

### 1.1 Install Dependencies

From the project root (`ppt-automation-demo/`):

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install python-pptx pandas openpyxl numpy
```

Python 3.9+ is recommended.

---

## Part 2: Generate Raw Sample Data

### 2.1 Run Data Generator

```bash
python generate_sample_data.py
```

This script:

- Creates `data/raw/survey_responses.xlsx`
- Generates data for 5 departments:
  - Sales
  - Engineering
  - Marketing
  - Customer Service
  - Operations
  - For each employee, it generates 4 quarters of survey responses
  - Columns include: `department`, `employee_id`, `quarter`, `age`,
    `responded`, `satisfaction`, `engagement`, `retained`

Example schema (simplified):

```text
department | employee_id | quarter | age | responded | satisfaction | engagement | retained
------------------------------------------------------------------------------------------
Sales      | Sales_001   | Q1      | 28  | 1         | 0.82         | 0.79       | 1
Sales      | Sales_001   | Q2      | 28  | 0         |              |            | 1
Sales      | Sales_002   | Q1      | 35  | 1         | 0.76         | 0.73       | 1
...
```

You can open the file in Excel to inspect or replace it with your own raw data.

---

## Part 3: Aggregate Raw Data → Department Metrics

### 3.1 Run Data Preparation Script

```bash
python prepare_data.py
```

This script:

- Reads `data/raw/survey_responses.xlsx`
- Aggregates to department-level metrics and writes:
  - `data/processed/department_metrics.xlsx`
- For each department, it computes:
  - Trend metrics: `satisfaction`, `engagement`, `response_rate`, `retention` (Q1–Q4)
  - Age distribution: 18–25, 26–35, 36–45, 46–55, 56+
  - Summary info: department size, latest satisfaction/engagement/response rate

Processed Excel schema (simplified):

```text
department       metric_type        Q1      Q2      Q3      Q4      ...
-------------------------------------------------------------------------------
Sales            satisfaction       0.75    0.78    0.81    0.84
Sales            engagement         0.70    0.72    0.74    0.76
Sales            response_rate      0.80    0.81    0.82    0.83
Sales            retention          0.90    0.91    0.92    0.93
Sales            age_distribution   ... age_18_25, age_26_35, ...
Sales            info               ... department_size, latest_*, ...
Engineering      satisfaction       ...
...
```

This mirrors the format used by `batch_generate.py` and keeps a clean separation:
raw data → processed metrics → PPT.

---

## Part 4: Create the PPT Template

### 4.1 Generate Template Programmatically

```bash
python create_template.py
```

This script creates `data/raw/template.pptx` with:

- Slide size: 10 in × 7.5 in
- Three slides with **unified styling** (consistent fonts, colors, positions):

  **Slide 1 – Summary Dashboard**
  - Title: `[Department Name] - Summary Dashboard` (24pt, dark blue)
  - Subtitle: key conclusions across satisfaction, engagement, response, retention and age structure (12pt, gray)
  - Six business-style conclusion cards (2×3 grid), each a text box:
    - Overall Satisfaction (current level + QoQ change)
    - Employee Engagement (current level + QoQ change)
    - Response & Retention (survey response rate + retention rate)
    - Age Structure (largest age group and percentage)
    - Satisfaction vs Engagement (gap between the two metrics)
    - Overall Summary (one-sentence summary)

  **Slide 2 – Overview**
  - Title: `[Department Name] - Quarterly Survey Report` (24pt, dark blue)
  - Subtitle: `Q4 2024 | Department Size: XXX | Response Rate: XX%` (12pt, gray)
  - Four charts (2×2 grid, starting at 1.4" from top):
    - Satisfaction (line chart, %)
    - Engagement (line chart, %)
    - Response Rate (line chart, %)
    - Age Distribution (column chart, count)
  - Key findings text at the bottom (11pt)

  **Slide 3 – Detailed Metrics**
  - Title: `[Department Name] - Detailed Metrics` (24pt, dark blue - **same as Slide 2**)
  - Subtitle: `QoQ changes and demographic breakdown` (12pt, gray - **same as Slide 2**)
  - Key metrics table: Satisfaction / Engagement / Response Rate / Retention (Q3, Q4, QoQ Δ)
    - Starting at 1.4" from top - **aligned with Slide 1 charts**
  - Age distribution table: age groups vs share of respondents (%)
  - Notes section for auto-generated comments (11pt)

**Style Constants**: All colors, font sizes, and layout positions are defined as constants at the top of `create_template.py` (lines 20-50). This ensures:
- Visual consistency across both slides
- Easy customization - change once, apply everywhere
- No risk of style drift when adding new slides
- **Chart elements** optimized for professional appearance:
  - Chart titles: 14pt
  - Axis labels: 9pt (smaller, less distracting)
  - Legend text: 9pt
  - Data labels: 8pt (automatically shown on all data points)
- **Table elements** balanced for readability:
  - Headers: 11pt (bold)
  - Body cells: 9pt

The data in these charts is dummy data; it only defines the visual style and structure. The batch script later replaces the data but preserves the formatting.

If you want to design your own template in PowerPoint instead of using this script, you can:

1. Design a slide with the layout you want.
2. Save it as `data/raw/template.pptx`.
3. Use `tools/analyze_template.py` to inspect shape names and chart titles (see Part 6).

---

## Part 5: Batch Generate Department Reports

### 5.1 Run Batch Generation

```bash
python batch_generate.py
```

What it does:

1. Load processed Excel data from `data/processed/department_metrics.xlsx`.
2. Discover all departments in the data.
3. For each department:
   - Load a fresh copy of `data/raw/template.pptx`.
   - Extract that department’s data.
   - Update text and charts on slide 1 (overview).
   - Update tables and notes on slide 2 (detailed metrics).
   - Save as `data/output/reports/<Department>_Report.pptx`.

Example console output (simplified):

```text
================================================================================
📊 Batch PPT Generation - Starting
================================================================================

📄 Loading data...
  ✓ Found 5 departments
  ✓ Template loaded

================================================================================
📌 Generating Reports
================================================================================

[1/5] Sales...
  ✓ Updated 4 charts
  ✓ Saved: data/output/reports/Sales_Report.pptx

...

📊 Generation Complete
✅ Success: 5 reports
💡 All reports saved to: data/output/reports
```

Open any of the generated PPTX files (e.g. `Sales_Report.pptx`) to verify the result.

---

## Part 6: How the Automation Works

### 6.1 High-Level Flow

```text
Raw Excel (employee×quarter)
  → pandas DataFrame
  → prepare_data.py (aggregation)
  → data/processed/department_metrics.xlsx
  → extract_department_data()
  → update_slide() / update_detail_slide()
  → python-pptx → PPT report
```

### 6.2 Data Extraction (per department)

In `batch_generate.py`:

```python
def extract_department_data(df, department):
    dept_data = df[df['department'] == department]
    # Build a dict with trend metrics, age distribution, and info
    ...
    return data
```

The returned `data` dict contains keys like:

- `satisfaction`, `engagement`, `response_rate`, `retention`:

  ```python
  {
      "categories": ["Q1", "Q2", "Q3", "Q4"],
      "values": [0.75, 0.78, 0.81, 0.84],
  }
  ```

- `age_distribution`:

  ```python
  {
      "categories": ["18-25", "26-35", "36-45", "46-55", "56+"],
      "values": [20.5, 30.1, 25.3, 15.2, 8.9],
  }
  ```

- `info`:

  ```python
  {
      "department_size": 180,
      "latest_satisfaction": 0.84,
      "latest_engagement": 0.78,
      "latest_response_rate": 0.82,
  }
  ```

### 6.3 Text Updates (preserve formatting)

```python
def update_text_preserve_format(shape, new_text):
    """Update text while preserving formatting"""
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    if text_frame.paragraphs and text_frame.paragraphs[0].runs:
        # ✅ Correct – keeps font, size, color, etc.
        text_frame.paragraphs[0].runs[0].text = new_text
    else:
        # Fallback – may reset formatting
        text_frame.text = new_text

def update_table_cell_preserve_format(cell, new_text):
    """Update table cell text while preserving formatting"""
    text_frame = cell.text_frame
    if text_frame.paragraphs and text_frame.paragraphs[0].runs:
        # ✅ Correct – keeps font size, color, etc.
        text_frame.paragraphs[0].runs[0].text = new_text
    else:
        # Fallback – may reset formatting
        cell.text = new_text
```

**CRITICAL**: These functions are used to:

- Update text boxes: titles, subtitles, key findings
- Update table cells: all cells in metrics and age tables
- **Why**: Direct assignment (`shape.text`, `cell.text`) **destroys all formatting** including font size, color, bold, etc.

### 6.4 Chart Updates (preserve styling + auto-scale)

Core helper in `batch_generate.py`:

```python
from pptx.chart.data import CategoryChartData
from pptx.util import Pt

def update_chart_data(chart, categories, values, number_format=None,
                      auto_scale=True, scale_factor=1.2):
    """Update chart data while preserving template styling."""
    chart_data = CategoryChartData()
    chart_data.categories = categories

    # Reuse existing series name if present
    if chart.series:
        series_name = chart.series[0].name
    else:
        series_name = "Series 1"

    chart_data.add_series(series_name, values)
    chart.replace_data(chart_data)  # Preserves colors, line styles, etc.

    # CRITICAL: Re-apply font sizes and number format after replace_data()
    # chart.replace_data() may reset some formatting properties
    chart.category_axis.tick_labels.font.size = Pt(9)

    # IMPORTANT: Number format must be set via tick_labels, not axis directly
    value_labels = chart.value_axis.tick_labels
    value_labels.font.size = Pt(9)
    if number_format:
        value_labels.number_format = number_format
        value_labels.number_format_is_linked = False  # Critical!

    chart.legend.font.size = Pt(9)

    # Re-enable and format data labels
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.size = Pt(8)
    if number_format:
        data_labels.number_format = number_format
        data_labels.number_format_is_linked = False  # Critical!

    # Optional: auto-scale Y axis based on values
    if auto_scale and values:
        ...
```

**Key points**:

- `chart.replace_data()` preserves **visual styles** (colors, line types) but may **reset font sizes and data labels**
- **Must re-apply** font sizes and data label settings after `replace_data()`
- **Number format critical**: Must use `tick_labels.number_format` and set `number_format_is_linked = False`
- Automatically adjusts the Y-axis range based on data size (with padding)
- This pattern learned from production projects ensures 100% format consistency

**Number Format Strings**:
- `'0%'` - Standard percentage: 0.82 → 82% (auto multiply by 100)
- `'0.0%'` - With decimals: 0.825 → 82.5%
- `'0"%"'` - Custom percentage: 20.5 → 20% (no multiplication)
- `'0.0"%"'` - Custom with decimals: 20.5 → 20.5%

### 6.5 Intelligent Shape & Chart Mapping + Mapping Spec

`update_slide()` and `update_detail_slide()` use an **external mapping spec** to
find shapes and tie them to data:

1. `docs/data_mapping.md` defines, in a markdown table, for each element:
   - slide index, element type (`text` / `chart` / `table`)
   - logical name (e.g. `chart_satisfaction`)
   - shape name (e.g. `Chart_Satisfaction`)
   - data key (e.g. `satisfaction`, `info.subtitle`)
   - number format and fallback chart title keyword
2. `mapping.py` parses this markdown file at runtime and exposes:
   - `shape_names` mapping (logical name → shape name)
   - chart mappings used by `update_slide()`
3. `update_slide()` tries to find shapes in two ways:
   - **By name** (recommended): via `SHAPE_NAMES` loaded from the mapping spec
   - **By content**: fallback search using text or chart title keywords like
     `"Quarterly Survey Report"`, `"Satisfaction"`, `"Age"`, etc.

This design makes the markdown document both human documentation and
machine-readable configuration (**“documentation-as-code”**) and keeps
the code clean when templates evolve.

---

## Part 7: Adapting to Your Own Data & Template

### 7.1 Customize Data Structure

If your real data schema is different:

1. Modify `generate_sample_data.py` or replace `data/raw/survey_responses.xlsx`
   with your own raw data.
2. Update `prepare_data.py` to aggregate your raw schema into
   `data/processed/department_metrics.xlsx` in whatever format you need.
3. Update `docs/data_mapping.md` to map new metrics to PPT shapes.
4. Update `extract_department_data()` in `batch_generate.py` if you change the
   processed file structure:
   - Change how you filter rows (e.g. by region, product, BU).
   - Change what metrics you compute and return.

Example (conceptual):

```python
dept_data = df[df["region"] == department]
data["revenue"] = {
    "categories": ["Jan", "Feb", "Mar"],
    "values": [...],
}
```

### 7.2 Customize Font Sizes and Styling

All font sizes are defined as constants at the top of `create_template.py`. To adjust:

```python
# In create_template.py, modify these constants (lines 28-37):
FONT_SIZE_CHART_TITLE = 14      # Make chart titles bigger/smaller
FONT_SIZE_CHART_AXIS = 9        # Adjust axis label size
FONT_SIZE_CHART_LEGEND = 9      # Adjust legend size
FONT_SIZE_CHART_DATA_LABEL = 8  # Adjust data point labels
FONT_SIZE_TABLE_BODY = 9        # Adjust table cell text
```

After changing constants, regenerate the template:

```bash
python create_template.py
python batch_generate.py
```

All reports will use the new font sizes!

### 7.3 Adapt to a Different PPT Template

If you have an existing PPT template:

1. Save it as `data/raw/template.pptx`.
2. Inspect structure:

   ```bash
   python tools/analyze_template.py data/raw/template.pptx
   ```

   This prints:

   - Slide indices
   - Shape indices and names
   - For charts: title, series names

3. Rename shapes in PowerPoint (Selection Pane) to semantic names, e.g.:
   - `TextBox_Title`
   - `TextBox_Subtitle`
   - `Chart_Satisfaction`
4. Align `update_slide()` to your template:
   - Update the shape name mapping.
   - Adjust text contents and chart mappings.

---

## Part 8: Troubleshooting

### Charts not updating

- Check that chart titles in the template contain the expected keywords:
  - `"Satisfaction"`, `"Engagement"`, `"Response Rate"`, `"Age"`, etc.
- Verify that `extract_department_data()` returns the expected structure:

  ```python
  print(dept_data)
  print(dept_data["satisfaction"])
  ```

### Format lost

- Confirm that text updates go through `update_text_preserve_format()`.
- Avoid `shape.text = ...` unless you are OK with losing formatting.

### Number format not showing (e.g., 0.82 instead of 82%)

This is a **critical** issue! The number format must be set correctly:

```python
# ❌ WRONG - does NOT work
chart.value_axis.number_format = '0%'

# ✅ CORRECT - must use tick_labels
value_labels = chart.value_axis.tick_labels
value_labels.number_format = '0%'
value_labels.number_format_is_linked = False  # Must set to False!

# Same for data labels
data_labels.number_format = '0%'
data_labels.number_format_is_linked = False  # Must set to False!
```

**Why**:
- PowerPoint's number format API requires using `tick_labels` object
- The `number_format_is_linked` flag must be set to `False` to use custom format
- Without `number_format_is_linked = False`, PowerPoint ignores your format string

### Wrong department data

- Confirm department names in Excel match exactly the names you expect.
- Print out `departments = df["department"].unique()` to see what the script discovered.

### Template issues

- Run:

  ```bash
  python tools/analyze_template.py data/raw/template.pptx
  ```

- Check that:
  - The slide actually has charts.
  - The shape names / titles match what `update_slide()` is searching for.

---

## Part 9: Best Practices & Next Steps

- Design your PPT template first; treat it as the source of truth for formatting.
- Keep data extraction (`pandas`) and PPT update (`python-pptx`) logic clearly separated.
- Start simple:
  - One department, one slide → then expand to multiple departments and more content.
- Read more:
  - [python-pptx documentation](https://python-pptx.readthedocs.io/)
  - Project README: `README.md`
  - Methodology guide: `../.claude/skills/ppt-automation/SKILL.md`

You can now safely customize both the data and the template while reusing the same automation pipeline.
