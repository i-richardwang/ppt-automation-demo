#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate Sample Data for PPT Automation Demo

Creates sample survey data for multiple departments with various metrics.
Demonstrates the value of batch automation.
"""

import pandas as pd
import numpy as np
import os


def generate_sample_data():
    """Generate raw, row-level survey data for multiple departments.

    This function produces realistic "raw" data at the employee×quarter
    granularity, which will be aggregated by the data preparation script
    to generate department-level metrics needed for PPT automation.
    """

    # Departments to generate reports for (batch processing demo)
    departments = [
        "Sales",
        "Engineering",
        "Marketing",
        "Customer Service",
        "Operations",
    ]

    # Time periods (quarters)
    quarters = ["Q1", "Q2", "Q3", "Q4"]

    all_rows = []

    for dept in departments:
        # Department-specific random generator (stable per department)
        dept_seed = abs(hash(dept)) % (2**32)
        rng = np.random.default_rng(dept_seed)

        # Department size
        dept_size = int(rng.integers(50, 500))

        # Age distribution for employees (used to derive age groups later)
        age_groups = ["18-25", "26-35", "36-45", "46-55", "56+"]
        age_probs = rng.dirichlet(np.ones(len(age_groups)))
        group_counts = rng.multinomial(dept_size, age_probs)

        employees = []
        emp_id = 1
        for group, count in zip(age_groups, group_counts):
            for _ in range(count):
                if group == "18-25":
                    age = int(rng.integers(18, 26))
                elif group == "26-35":
                    age = int(rng.integers(26, 36))
                elif group == "36-45":
                    age = int(rng.integers(36, 46))
                elif group == "46-55":
                    age = int(rng.integers(46, 56))
                else:
                    age = int(rng.integers(56, 66))

                employees.append(
                    {"employee_id": f"{dept}_{emp_id:03d}", "age": age}
                )
                emp_id += 1

        # Underlying department-level trends (true means)
        satisfaction_base = 0.70 + rng.uniform(0, 0.15)
        satisfaction_trend = rng.uniform(-0.02, 0.04)
        engagement_base = 0.65 + rng.uniform(0, 0.15)
        engagement_trend = rng.uniform(-0.01, 0.03)
        response_base = 0.75 + rng.uniform(0, 0.15)
        response_trend = rng.uniform(-0.02, 0.02)
        retention_base = 0.85 + rng.uniform(0, 0.10)
        retention_trend = rng.uniform(-0.01, 0.02)

        satisfaction_means = []
        engagement_means = []
        response_probs = []
        retention_probs = []

        for i in range(len(quarters)):
            satisfaction_means.append(
                float(
                    np.clip(
                        satisfaction_base
                        + i * satisfaction_trend
                        + rng.normal(0, 0.03),
                        0,
                        1,
                    )
                )
            )
            engagement_means.append(
                float(
                    np.clip(
                        engagement_base
                        + i * engagement_trend
                        + rng.normal(0, 0.03),
                        0,
                        1,
                    )
                )
            )
            response_probs.append(
                float(
                    np.clip(
                        response_base
                        + i * response_trend
                        + rng.normal(0, 0.03),
                        0.4,
                        0.95,
                    )
                )
            )
            retention_probs.append(
                float(
                    np.clip(
                        retention_base
                        + i * retention_trend
                        + rng.normal(0, 0.02),
                        0.7,
                        0.99,
                    )
                )
            )

        # Generate row-level data: employee × quarter
        for q_idx, quarter in enumerate(quarters):
            sat_mean = satisfaction_means[q_idx]
            eng_mean = engagement_means[q_idx]
            resp_p = response_probs[q_idx]
            retain_p = retention_probs[q_idx]

            for emp in employees:
                responded = rng.random() < resp_p
                if responded:
                    satisfaction = float(
                        np.clip(sat_mean + rng.normal(0, 0.08), 0, 1)
                    )
                    engagement = float(
                        np.clip(eng_mean + rng.normal(0, 0.08), 0, 1)
                    )
                else:
                    satisfaction = None
                    engagement = None

                retained = rng.random() < retain_p

                all_rows.append(
                    {
                        "department": dept,
                        "employee_id": emp["employee_id"],
                        "quarter": quarter,
                        "age": emp["age"],
                        "responded": int(responded),
                        "satisfaction": round(satisfaction, 4)
                        if satisfaction is not None
                        else None,
                        "engagement": round(engagement, 4)
                        if engagement is not None
                        else None,
                        "retained": int(retained),
                    }
                )

    # Create DataFrame
    df = pd.DataFrame(all_rows)

    # Save to Excel (raw data)
    output_path = "data/raw/survey_responses.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    df.to_excel(output_path, index=False)

    print("\n" + "="*80)
    print("✅ Raw Survey Data Generated Successfully")
    print("="*80)
    print(f"\nOutput (raw data): {output_path}")
    print(f"Departments: {len(departments)}")
    print(f"Total employees: {df['employee_id'].nunique()}")
    print(f"Total rows (employee × quarter): {len(df)}")
    print(f"\nDepartments list:")
    for dept in departments:
        print(f"  - {dept}")
    print("\n" + "="*80)

    return df


def print_instructions():
    """Print next steps for user"""
    print("\n📝 Next Steps:")
    print("="*80)
    print("\n1. Prepare processed department metrics from raw data:")
    print("   python prepare_data.py")
    print("\n2. Create PPT template:")
    print("   python create_template.py")
    print("\n3. Run batch generation:")
    print("   python batch_generate.py")
    print("\n4. Check outputs:")
    print("   data/output/reports/")
    print("   - Sales_Report.pptx")
    print("   - Engineering_Report.pptx")
    print("   - Marketing_Report.pptx")
    print("   - ...")
    print("\n💡 This demonstrates the power of automation:")
    print("   Instead of manually creating 5 reports,")
    print("   the script generates all of them in seconds!")
    print("\n" + "="*80)


def main():
    """Entry point"""
    df = generate_sample_data()
    print_instructions()


if __name__ == '__main__':
    main()
