# PPT Automation Demo

A demonstration project showing how to automate PowerPoint report generation using python-pptx.

**Key Scenario**: Batch generate customized reports for multiple departments from a single template - transforming hours of manual work into seconds of automated processing.

## Features

- ✅ **Batch Processing**: Generate reports for 5 departments in seconds, not hours
- ✅ **Format Preservation**: Update text and charts while keeping all formatting
- ✅ **Comprehensive Reports**: Each report has 2 slides – an overview (title, 4 charts, key findings) and a detailed metrics page (KPI tables + demographics)
- ✅ **Data-Driven**: Pull data from Excel and populate PPT automatically
- ✅ **Production-Ready**: Demonstrates real-world automation patterns

## Why Use Automation?

**Manual Process** (❌ Inefficient):
- Create 5 reports × 15 minutes each = 75 minutes
- Copy template, update title, update 4 charts, update metrics, update conclusions
- Error-prone: easy to miss updates or use wrong data
- Tedious and repetitive

**Automated Process** (✅ Efficient):
- Generate 5 reports in ~5 seconds
- Consistent formatting and data accuracy
- Easy to scale to 50 or 500 reports
- Focus on insights, not copy-paste

## Project Structure

```
ppt-automation-demo/
├── README.md                 # This file
├── requirements.txt          # Python dependencies
├── LICENSE                   # MIT License
├── generate_sample_data.py   # Generate raw, row-level survey data
├── prepare_data.py           # Aggregate raw data → department metrics
├── create_template.py        # Create 2-slide PPT template (overview + detail tables)
├── batch_generate.py         # Main script: batch generate reports from processed data
├── mapping.py                # Load data↔PPT mapping from markdown spec
├── tools/
│   └── analyze_template.py   # Analyze PPT template structure
├── data/
│   ├── raw/
│   │   ├── template.pptx     # PPT template (generated)
│   │   └── survey_responses.xlsx  # Raw survey data (generated)
│   ├── processed/
│   │   └── department_metrics.xlsx  # Aggregated metrics used by PPT
│   └── output/
│       └── reports/          # Generated reports (5 files)
└── docs/
    ├── tutorial.md           # Step-by-step tutorial
    └── data_mapping.md       # Data↔PPT mapping spec
```

## Quick Start

### 1. Install Dependencies

```bash
pip install python-pptx pandas openpyxl numpy
```

Or using `uv` (recommended):
```bash
uv sync
```

### 2. Generate Raw Sample Data

```bash
python generate_sample_data.py
```

Creates `data/raw/survey_responses.xlsx` with row-level data for 5 departments:
- Sales
- Engineering
- Marketing
- Customer Service
- Operations

Each row is an employee-quarter observation with: age, satisfaction, engagement, response flag, retention flag, etc.

### 3. Aggregate Raw Data → Department Metrics

```bash
python prepare_data.py
```

Creates `data/processed/department_metrics.xlsx` with per-department, per-quarter metrics:
- satisfaction, engagement, response rate, retention (Q1–Q4)
- age distribution (5 age groups)
- summary info (department size, latest metrics)

### 4. Create Template

```bash
python create_template.py
```

Creates `data/raw/template.pptx` with three slides:

**Slide 1 – Summary Dashboard**
- 6 business-style conclusion cards summarizing:
  - Overall satisfaction (level + QoQ change)
  - Employee engagement (level + QoQ change)
  - Response rate & retention
  - Age structure (largest age group)
  - Satisfaction vs engagement gap
  - Overall summary sentence

**Slide 2 – Overview**
- Dynamic title and subtitle
- 4 charts in 2×2 grid (Satisfaction, Engagement, Response Rate, Age Distribution)
- Key findings section

**Slide 3 – Detailed Metrics**
- Title and subtitle
- Key metrics QoQ table (Satisfaction, Engagement, Response Rate, Retention)
- Age distribution table
- Notes / actions area

### 5. Batch Generate Reports

```bash
python batch_generate.py
```

Generates 5 customized reports in `data/output/reports/`:
- `Sales_Report.pptx`
- `Engineering_Report.pptx`
- `Marketing_Report.pptx`
- `Customer Service_Report.pptx`
- `Operations_Report.pptx`

**Result**: 5 complete reports generated in seconds! 🎉

### (Optional) One-Click Pipeline

To run the full pipeline (raw → processed → template → PPT) in one command:

```bash
python run_pipeline.py
```

This runs the complete end-to-end pipeline demonstrating an engineering-grade automation workflow.

## How It Works

### Template Design

Each report has two slides with **consistent styling** (unified fonts, colors, and layout):

**Slide 1: Overview**
1. **Title**: `[Department] - Quarterly Survey Report` (24pt, dark blue)
2. **Subtitle**: `Q4 2024 | Department Size: XXX | Response Rate: XX%` (12pt, gray)
3. **Four Charts** (2×2 grid, starting at 1.4"):
   - Satisfaction Trend (line chart, percentage) with data labels
   - Engagement Trend (line chart, percentage) with data labels
   - Response Rate Trend (line chart, percentage) with data labels
   - Age Distribution (bar chart, count) with data labels
4. **Key Findings**: Auto-generated summary with QoQ change analysis (11pt)

**Slide 2: Detailed Metrics**
1. **Title & Subtitle**: Same styling as Slide 1 for consistency
2. **Key Metrics Table** (Satisfaction, Engagement, Response Rate, Retention):
   - Columns: Metric, Q3, Q4, QoQ Δ (percentage points)
   - Starting position aligned with Slide 1 charts (1.4")
3. **Age Distribution Table**:
   - One row per age group, with share of respondents (%)
4. **Notes / Actions**:
   - Auto-generated bullet points highlighting key changes and largest segment

**Style Constants** (defined in `create_template.py`):
- All colors, font sizes, and positions are managed via constants
- **Chart elements** optimized for clarity:
  - Title: 14pt (prominent but not overwhelming)
  - Axis labels: 9pt (readable, not competing with data)
  - Legend: 9pt (subtle, consistent with axes)
  - Data labels: 8pt (visible on data points)
- **Table elements** sized for readability:
  - Headers: 11pt (bold, clear)
  - Body text: 9pt (compact, professional)
- Ensures visual consistency across both slides
- Easy to customize by changing constants once

### Data Pipeline & Structures

This demo now models a simple **three-step pipeline**:

1. **Raw data** – employee×quarter level (row-level survey records)  
   `data/raw/survey_responses.xlsx`

   ```
   department | employee_id | quarter | age | responded | satisfaction | engagement | retained
   Sales      | Sales_001   | Q1      | 28  | 1         | 0.82         | 0.79       | 1
   Sales      | Sales_001   | Q2      | 28  | 0         |              |            | 1
   ...
   ```

2. **Processed metrics** – department-level aggregates used by PPT  
   `data/processed/department_metrics.xlsx`

   ```
   department | metric_type      | Q1   | Q2   | Q3   | Q4   | ...
   Sales      | satisfaction     | 0.75 | 0.78 | 0.81 | 0.84 |
   Sales      | engagement       | 0.70 | 0.72 | 0.74 | 0.76 |
   Sales      | response_rate    | 0.82 | 0.83 | 0.84 | 0.85 |
   Sales      | retention        | 0.90 | 0.91 | 0.92 | 0.93 |
   Sales      | age_distribution | ...  | ...  | ...  | ...  |
   Sales      | info             | ... department_size, latest_* ...
   ```

3. **PPT mapping** – configuration describing how processed metrics map to PPT elements  
   `docs/data_mapping.md` (parsed at runtime via `mapping.py`)

   The mapping table defines, for each slide element:
   - which shape name it uses (`shape_name`)
   - which key in the processed data it reads (`data_key`)
   - which number format to apply (`number_format`)
   - fallback chart title keyword (`title_keyword`)

### Core Update Function

```python
from pptx.chart.data import CategoryChartData

def update_chart_data(chart, categories, values, number_format=None,
                      auto_scale=True, scale_factor=1.2):
    """Update chart data while preserving template styling"""
    chart_data = CategoryChartData()
    chart_data.categories = categories

    # Reuse existing series name from the template if available
    if chart.series:
        series_name = chart.series[0].name
    else:
        series_name = "Series 1"

    chart_data.add_series(series_name, values)

    # This preserves most formatting: colors, line styles, etc.
    chart.replace_data(chart_data)

    # CRITICAL: Re-apply font sizes and number format after replace_data()
    # chart.replace_data() may reset some properties like font sizes
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

    # Optional: auto-scale Y axis based on the data range
    if auto_scale:
        ...
```

### Number Format (Critical for Charts)

```python
# ❌ Wrong - does NOT work
chart.value_axis.number_format = '0%'

# ✅ Correct - must set via tick_labels
value_labels = chart.value_axis.tick_labels
value_labels.number_format = '0%'
value_labels.number_format_is_linked = False  # Critical! Must be False

# ✅ Correct - data labels also need number_format_is_linked
data_labels = plot.data_labels
data_labels.number_format = '0%'
data_labels.number_format_is_linked = False  # Critical!
```

**Number Format Strings:**
- `'0%'` - Standard percentage (0.82 → 82%, auto multiply by 100)
- `'0.0%'` - One decimal (0.825 → 82.5%)
- `'0"%"'` - Custom percentage (20.5 → 20%, no multiplication)
- `'0.0"%"'` - Custom with decimal (20.5 → 20.5%)

### Text Update (Format Preservation)

```python
# ✅ Correct - preserves formatting for text boxes
if paragraph.runs:
    paragraph.runs[0].text = new_text

# ❌ Wrong - destroys formatting
shape.text = new_text

# ✅ Correct - preserves formatting for table cells
def update_table_cell_preserve_format(cell, new_text):
    text_frame = cell.text_frame
    if text_frame.paragraphs and text_frame.paragraphs[0].runs:
        text_frame.paragraphs[0].runs[0].text = new_text
    else:
        cell.text = new_text  # Fallback

# ❌ Wrong - destroys table cell formatting
cell.text = new_text
```

## Use Cases

This demo pattern applies to:

- **Survey Reports**: Department/region-specific survey results
- **Performance Dashboards**: Monthly KPI reports for multiple teams
- **Sales Reports**: Regional sales performance with charts
- **Financial Reports**: Department budget vs. actual analysis
- **HR Reports**: Headcount and attrition reports by business unit

**Key Requirement**: When you need to create the same report structure for multiple entities with different data.

## Customization

### Adapt to Your Data

1. **Modify Raw Data Structure**: Edit `generate_sample_data.py` (or replace `survey_responses.xlsx`) to match your real raw data schema.
2. **Update Aggregation Logic**: Adjust `prepare_data.py` to compute the department-level metrics you need.
3. **Update Mapping Spec**: Edit `docs/data_mapping.md` to map new metrics to PPT shapes (the runtime config is parsed from this table).
4. **Update Extraction Logic**: If you change the processed file format, update `extract_department_data()` in `batch_generate.py`.
5. **Change Chart Types**: Modify `create_template.py` to use different chart types or layouts.
6. **Add More Charts**: Extend the template and mapping spec to include additional visualizations.

### Adapt to Your Template

1. **Analyze Your Template**:
   ```bash
   python tools/analyze_template.py your_template.pptx
   ```
2. **Name Shapes Semantically**: Use PowerPoint's Selection Pane
3. **Update Mapping Logic**: Modify `update_slide()` to match your shape names
4. **Test Incrementally**: Start with one chart, then expand

## Key Principles

1. **Template as Source of Truth**: PPT defines formatting, code only updates data
2. **Batch Processing**: Pre-load data once, generate multiple reports efficiently
3. **Format Preservation - Critical!**:
   - **Text boxes**: Use `paragraph.runs[0].text = new_text`, never `shape.text`
   - **Table cells**: Use `text_frame.paragraphs[0].runs[0].text`, never `cell.text`
   - **Charts**: Call `chart.replace_data()` then **re-apply font sizes and data labels**
   - **Number format**: Use `tick_labels.number_format` and set `number_format_is_linked = False`
   - **Why**: Direct text assignment (`shape.text`, `cell.text`) destroys all formatting
   - **Why**: `chart.replace_data()` may reset font sizes and data label settings
   - **Why**: Number format must be set via `tick_labels` API with linked flag disabled
4. **Configuration-Based**: Separate data extraction from PPT update logic
5. **Error Handling**: Validate data before processing, provide clear error messages

## Common Patterns

### Pattern 1: Multi-Entity Reports (This Demo)
One template → Multiple customized reports (one per entity)

### Pattern 2: Single Report with Multiple Sections
One template with multiple slides → Update different sections with different data sources

### Pattern 3: Time-Series Reports
One template → Generate reports for different time periods (monthly, quarterly)

## Learn More

- [docs/tutorial.md](docs/tutorial.md) - Detailed step-by-step guide
- [.claude/skills/ppt-automation/SKILL.md](../.claude/skills/ppt-automation/SKILL.md) - Complete methodology
- [python-pptx documentation](https://python-pptx.readthedocs.io/) - Library reference

## Troubleshooting

### Charts not updating
- Check chart titles match exactly (case-insensitive matching in demo)
- Verify data extraction returns correct format
- Use `analyze_template.py` to inspect structure

### Format lost
- Ensure you're using `paragraph.runs[0].text`, not `shape.text`
- Check if text frame has at least one paragraph and run

### Wrong data
- Print department data before updating: `print(dept_data)`
- Verify Excel data structure matches expected format
- Check department names match exactly

## License

MIT License - Feel free to use and adapt for your projects

---

**Built with**: python-pptx, pandas, numpy
**Purpose**: Demonstrate real-world PPT automation patterns
**Audience**: Data analysts, automation engineers, anyone doing repetitive PPT work
