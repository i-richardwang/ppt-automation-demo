#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Run the full PPT automation pipeline in one command.

Steps:
1) Generate raw, row-level survey data (employee × quarter)
2) Aggregate raw data into department-level metrics
3) Create PPT template (2 slides)
4) Batch-generate PPT reports for all departments

Usage:
    python scripts/run_pipeline.py
"""

import sys
import os

# Add project root to path so we can import from scripts/
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Import using importlib to handle numeric prefixes
import importlib.util

def import_module_from_file(filepath, module_name):
    """Import a module from a file path"""
    spec = importlib.util.spec_from_file_location(module_name, filepath)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

# Get script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Import modules
gen_data = import_module_from_file(
    os.path.join(script_dir, '1_generate_sample_data.py'),
    'generate_sample_data'
)
prep_data = import_module_from_file(
    os.path.join(script_dir, '2_prepare_data.py'),
    'prepare_data'
)
create_tpl = import_module_from_file(
    os.path.join(script_dir, '3_create_template.py'),
    'create_template'
)
batch_gen = import_module_from_file(
    os.path.join(script_dir, '4_batch_generate.py'),
    'batch_generate'
)

generate_sample_data = gen_data.generate_sample_data
prepare_department_metrics = prep_data.prepare_department_metrics
create_template = create_tpl.create_template
batch_generate = batch_gen.batch_generate


def main():
    print("\n" + "=" * 80)
    print("🚀 Running Full PPT Automation Pipeline")
    print("=" * 80)

    # 1) Generate raw data
    print("\n[1/4] Generating raw survey data...")
    df_raw = generate_sample_data()

    # 2) Aggregate raw → processed metrics
    print("\n[2/4] Preparing department-level metrics...")
    df_processed = prepare_department_metrics(
        raw_path="data/raw/survey_responses.xlsx",
        output_path="data/processed/department_metrics.xlsx",
    )

    # 3) Create PPT template
    print("\n[3/4] Creating PPT template...")
    create_template()

    # 4) Batch-generate PPT reports
    print("\n[4/4] Generating PPT reports...")
    results = batch_generate(
        template_path="data/raw/template.pptx",
        excel_path="data/processed/department_metrics.xlsx",
        output_dir="data/output/reports",
    )

    success = len(results["success"]) if results else 0
    failed = len(results["failed"]) if results else 0

    print("\n" + "=" * 80)
    print("✅ Pipeline Finished")
    print("=" * 80)
    print(f"Reports generated: {success} success, {failed} failed")
    print("Raw data:        data/raw/survey_responses.xlsx")
    print("Processed data:  data/processed/department_metrics.xlsx")
    print("Template:        data/raw/template.pptx")
    print("Reports folder:  data/output/reports/")
    print("=" * 80 + "\n")


if __name__ == "__main__":
    main()

