# Data Mapping Specification

> This file is the **single source of truth** describing how processed data
> is mapped into the PPT template.  
> The `batch_generate.py` script parses this table at runtime, so updating
> this document **directly changes the behavior** of the automation.

| slide_index | element_type | logical_name        | shape_name           | data_key          | number_format | title_keyword          | description                                                   |
|------------:|-------------|---------------------|----------------------|-------------------|--------------|------------------------|---------------------------------------------------------------|
| 0           | text        | title               | TextBox_Title        | department_name   |              | Quarterly Survey Report | Slide 1 main title with department name                       |
| 0           | text        | subtitle            | TextBox_Subtitle     | info.subtitle     |              | Department Size        | Slide 1 subtitle with quarter, department size, response rate |
| 0           | text        | conclusion          | TextBox_Conclusion   | conclusion        |              | Key Findings           | Auto-generated key findings summary                           |
| 0           | chart       | chart_satisfaction  | Chart_Satisfaction   | satisfaction      | 0%           | Satisfaction           | Satisfaction trend (line chart, percentage)                   |
| 0           | chart       | chart_engagement    | Chart_Engagement     | engagement        | 0%           | Engagement             | Engagement trend (line chart, percentage)                     |
| 0           | chart       | chart_response      | Chart_Response       | response_rate     | 0%           | Response               | Response rate trend (line chart, percentage)                  |
| 0           | chart       | chart_age           | Chart_Age            | age_distribution  | 0.0"%"       | Age                    | Age distribution (bar chart, share of respondents)            |
| 1           | text        | detail_title        | TextBox_DetailTitle  | department_name   |              | Detailed Metrics       | Slide 2 title with department name                            |
| 1           | text        | detail_subtitle     | TextBox_DetailSubtitle | info.detail_subtitle |           | QoQ changes           | Slide 2 subtitle describing detailed metrics                  |
| 1           | table       | table_metrics       | Table_Metrics        | metrics           |              |                        | QoQ metrics table (satisfaction, engagement, response, retention) |
| 1           | table       | table_age           | Table_Age            | age_distribution  |              |                        | Age distribution table                                        |
| 1           | text        | notes               | TextBox_Notes        | notes             |              | Key Insights           | Auto-generated key insights                                   |

## Column Definitions

- `slide_index`  
  Zero-based slide index in the PPT template (0 = overview, 1 = detail).

- `element_type`  
  - `text`  – text box whose content is filled via automation  
  - `chart` – chart whose data is replaced via automation  
  - `table` – table whose cells are filled via automation

- `logical_name`  
  Stable identifier used in code to refer to this PPT element.

- `shape_name`  
  Name of the shape in PowerPoint's **Selection Pane**.  
  This should match exactly (case-sensitive) and is the primary way the code
  finds shapes on each slide.

- `data_key`  
  Key in the processed department data dictionary used by `batch_generate.py`.
  Nested keys (e.g. `info.subtitle`) indicate values derived from the `info`
  section rather than directly present in the raw data.

- `number_format`  
  Optional number format applied to chart axes / data labels, e.g.:
  - `0%`       – 0.82 → 82% (auto ×100)  
  - `0.0"%"`   – 20.5 → 20.5% (no ×100)  

- `title_keyword`  
  Fallback keyword used to locate a chart by its visible title when the shape
  name is missing or changed.

- `description`  
  Human-readable explanation of what the element represents.

---

To modify the mapping (e.g., add a new chart or rename a shape), update this
table and re-run `batch_generate.py`. No code changes are required as long as
you follow the existing column conventions.
