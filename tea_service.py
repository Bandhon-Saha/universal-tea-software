
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List
import shutil
import tempfile

import pandas as pd

from engine import universal_tea_engine_v06 as tea


@dataclass
class AnalysisResult:
    case_name: str
    process_category: str
    primary_product: str
    currency: str
    product_unit: str
    target_return: float
    warnings: List[str]

    baseline: Dict[str, object]
    sensitivity: pd.DataFrame
    scenarios: pd.DataFrame
    break_even: pd.DataFrame
    scale: pd.DataFrame
    environmental: pd.DataFrame
    validation: pd.DataFrame
    uncertainty_summary: pd.DataFrame
    uncertainty_samples: pd.DataFrame
    uncertainty_drivers: pd.DataFrame
    rd_priority: pd.DataFrame
    decision_summary: pd.DataFrame

    results_excel: Path
    word_report: Path
    figures: Dict[str, Path]
    work_dir: Path


def _figure_key(path: Path) -> str:
    name = path.stem.lower()
    mappings = [
        ("uncertainty_drivers", "uncertainty_drivers"),
        ("uncertainty_npv", "uncertainty_npv"),
        ("uncertainty_msp", "uncertainty_msp"),
        ("sensitivity_tornado", "sensitivity_tornado"),
        ("sensitivity", "sensitivity"),
        ("scenarios", "scenarios"),
        ("break_even", "break_even"),
        ("scale", "scale"),
        ("environmental", "environmental"),
        ("validation", "validation"),
        ("rnd_priority", "rd_priority"),
        ("capex", "capex"),
        ("opex", "opex"),
        ("cash_flow", "cash_flow"),
    ]
    for token, key in mappings:
        if token in name:
            return key
    return name


def run_analysis(workbook_path: str | Path, work_dir: str | Path | None = None) -> AnalysisResult:
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    if work_dir is None:
        work_dir = Path(tempfile.mkdtemp(prefix="universal_tea_v06_"))
    else:
        work_dir = Path(work_dir)
        work_dir.mkdir(parents=True, exist_ok=True)

    local_workbook = work_dir / workbook_path.name
    if workbook_path.resolve() != local_workbook.resolve():
        shutil.copy2(workbook_path, local_workbook)

    ctx, rd, decision_df, validation_targets = tea.parse_model(str(local_workbook))

    baseline = tea.evaluate_case(ctx)
    sensitivity = tea.run_sensitivity(ctx, decision_df)
    scenarios = tea.run_scenarios(ctx, decision_df)
    break_even = tea.run_break_even(ctx, decision_df)
    scale = tea.run_scale(ctx, decision_df)

    uncertainty_summary, uncertainty_samples, uncertainty_drivers = tea.run_uncertainty(
        ctx, rd, baseline
    )
    rd_priority = tea.build_rd_priority(
        ctx, baseline, sensitivity, rd, uncertainty_drivers
    )
    environmental = tea.calculate_environmental(ctx, rd, baseline)
    validation = tea.validation_comparison(
        ctx, validation_targets, baseline, sensitivity
    )
    decision_summary = tea.build_decision_summary(
        ctx,
        baseline,
        sensitivity,
        scenarios,
        break_even,
        uncertainty_summary,
        rd_priority,
        environmental,
        validation,
        rd,
    )

    stem = tea.slugify(local_workbook.stem)
    results_excel = work_dir / f"{stem}_TEA_Results_V0_6.xlsx"
    word_report = work_dir / f"{stem}_TEA_Report_V0_6.docx"

    tea.write_results_excel(
        results_excel,
        ctx,
        baseline,
        sensitivity,
        scenarios,
        break_even,
        scale,
        environmental,
        validation,
        uncertainty_summary,
        uncertainty_samples,
        uncertainty_drivers,
        rd_priority,
        decision_summary,
    )

    tea.write_word_report(
        word_report,
        ctx,
        baseline,
        sensitivity,
        scenarios,
        break_even,
        environmental,
        validation,
        uncertainty_summary,
        rd_priority,
        decision_summary,
    )

    figure_paths = tea.save_charts(
        ctx,
        work_dir,
        stem,
        baseline,
        sensitivity,
        scenarios,
        validation,
        break_even=break_even,
        scale=scale,
        environmental=environmental,
        uncertainty_summary=uncertainty_summary,
        uncertainty_samples=uncertainty_samples,
        uncertainty_drivers=uncertainty_drivers,
        rd_priority=rd_priority,
    )

    figures = {_figure_key(Path(p)): Path(p) for p in figure_paths}

    return AnalysisResult(
        case_name=ctx.case_name,
        process_category=ctx.process_category,
        primary_product=ctx.primary_product,
        currency=ctx.currency,
        product_unit=ctx.product_unit,
        target_return=ctx.target_return,
        warnings=list(ctx.warnings),
        baseline=baseline,
        sensitivity=sensitivity,
        scenarios=scenarios,
        break_even=break_even,
        scale=scale,
        environmental=environmental,
        validation=validation,
        uncertainty_summary=uncertainty_summary,
        uncertainty_samples=uncertainty_samples,
        uncertainty_drivers=uncertainty_drivers,
        rd_priority=rd_priority,
        decision_summary=decision_summary,
        results_excel=results_excel,
        word_report=word_report,
        figures=figures,
        work_dir=work_dir,
    )
