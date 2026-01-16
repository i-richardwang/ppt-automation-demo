#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Data preparation script for PPT Automation Demo.

Pipeline:
1) Raw data (employee × quarter) in `data/raw/survey_responses.xlsx`
2) Aggregated department metrics in `data/processed/department_metrics.xlsx`
3) PPT generation reads the processed file and updates charts/tables.
"""

import os
from typing import List

import numpy as np
import pandas as pd


QUARTERS: List[str] = ["Q1", "Q2", "Q3", "Q4"]


def prepare_department_metrics(
    raw_path: str = "data/raw/survey_responses.xlsx",
    output_path: str = "data/processed/department_metrics.xlsx",
) -> pd.DataFrame:
    """Aggregate raw survey data to department-level metrics."""

    if not os.path.exists(raw_path):
        raise FileNotFoundError(
            f"Raw data not found: {raw_path}. "
            "Run `python generate_sample_data.py` first."
        )

    df = pd.read_excel(raw_path)

    required_cols = {
        "department",
        "employee_id",
        "quarter",
        "age",
        "responded",
        "satisfaction",
        "engagement",
        "retained",
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in raw data: {sorted(missing)}")

    departments = df["department"].unique().tolist()

    output_rows = []

    for dept in departments:
        dept_df = df[df["department"] == dept].copy()
        dept_size = int(dept_df["employee_id"].nunique())

        # Guard against empty departments
        if dept_size == 0:
            continue

        metric_values = {
            "satisfaction": [],
            "engagement": [],
            "response_rate": [],
            "retention": [],
        }

        last_value = {key: None for key in metric_values.keys()}

        for quarter in QUARTERS:
            q_df = dept_df[dept_df["quarter"] == quarter]

            if q_df.empty:
                for key in metric_values.keys():
                    val = last_value[key] if last_value[key] is not None else 0.0
                    metric_values[key].append(val)
                    last_value[key] = val
                continue

            responded_df = q_df[q_df["responded"] == 1]

            # Satisfaction / Engagement: mean of respondents
            sat_val = (
                float(responded_df["satisfaction"].mean())
                if not responded_df.empty
                else None
            )
            eng_val = (
                float(responded_df["engagement"].mean())
                if not responded_df.empty
                else None
            )

            # Response rate: unique respondents / total employees
            if dept_size > 0:
                resp_val = float(
                    responded_df["employee_id"].nunique() / float(dept_size)
                )
            else:
                resp_val = None

            # Retention: average of retained flag
            ret_val = float(q_df["retained"].mean())

            quarter_vals = {
                "satisfaction": sat_val,
                "engagement": eng_val,
                "response_rate": resp_val,
                "retention": ret_val,
            }

            for key, val in quarter_vals.items():
                if val is None or np.isnan(val):
                    val = last_value[key] if last_value[key] is not None else 0.0
                metric_values[key].append(val)
                last_value[key] = val

        # Convert metric_values to rows (one row per metric_type)
        for metric_type, values in metric_values.items():
            row = {
                "department": dept,
                "metric_type": metric_type,
            }
            for quarter, val in zip(QUARTERS, values):
                row[quarter] = round(float(val), 4)
            output_rows.append(row)

        # Age distribution (based on unique employees)
        unique_emps = dept_df.drop_duplicates("employee_id")
        ages = unique_emps["age"].dropna()
        total_emp = len(ages)

        if total_emp > 0:
            def share(lower: int, upper: int) -> float:
                mask = (ages >= lower) & (ages <= upper)
                return float(mask.sum() / total_emp * 100.0)

            age_18_25 = share(18, 25)
            age_26_35 = share(26, 35)
            age_36_45 = share(36, 45)
            age_46_55 = share(46, 55)
            age_56_plus = share(56, 100)
        else:
            age_18_25 = age_26_35 = age_36_45 = age_46_55 = age_56_plus = 0.0

        output_rows.append(
            {
                "department": dept,
                "metric_type": "age_distribution",
                "age_18_25": round(age_18_25, 2),
                "age_26_35": round(age_26_35, 2),
                "age_36_45": round(age_36_45, 2),
                "age_46_55": round(age_46_55, 2),
                "age_56_plus": round(age_56_plus, 2),
            }
        )

        # Summary info row (latest metrics from Q4)
        latest_satisfaction = metric_values["satisfaction"][-1]
        latest_engagement = metric_values["engagement"][-1]
        latest_response_rate = metric_values["response_rate"][-1]

        output_rows.append(
            {
                "department": dept,
                "metric_type": "info",
                "department_size": dept_size,
                "latest_satisfaction": round(float(latest_satisfaction), 4),
                "latest_engagement": round(float(latest_engagement), 4),
                "latest_response_rate": round(float(latest_response_rate), 4),
            }
        )

    result_df = pd.DataFrame(output_rows)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    result_df.to_excel(output_path, index=False)

    print("\n" + "=" * 80)
    print("✅ Processed Department Metrics Generated Successfully")
    print("=" * 80)
    print(f"\nInput (raw): {raw_path}")
    print(f"Output (processed): {output_path}")
    print(f"Departments: {len(departments)}")
    print(f"Total rows: {len(result_df)}")
    print("=" * 80 + "\n")

    return result_df


def main():
    prepare_department_metrics()


if __name__ == "__main__":
    main()

