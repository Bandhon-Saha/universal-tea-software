
from __future__ import annotations

from pathlib import Path
import io
import shutil
import tempfile
import zipfile

import numpy as np
import pandas as pd
import streamlit as st

from tea_service import run_analysis


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = APP_DIR / "templates"

DEMO_CASES = {
    "Generic TEA demonstration": TEMPLATE_DIR / "Universal_TEA_Input_Workbook_V0_6.xlsx",
    "Validation Case 1 — H2A biomass gasification": TEMPLATE_DIR / "H2A_Biomass_Validation.xlsx",
    "Validation Case 2 — PEM electrolysis": TEMPLATE_DIR / "PEM_Electrolysis_Validation.xlsx",
}

st.set_page_config(
    page_title="Universal TEA Decision Support",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.6rem; padding-bottom: 2.5rem; max-width: 1500px;}
    [data-testid="stMetric"] {
        border: 1px solid rgba(120,120,120,0.22);
        border-radius: 12px;
        padding: 0.75rem 0.9rem;
        background: rgba(127,127,127,0.035);
    }
    [data-testid="stMetricLabel"] {font-weight: 600;}
    h1, h2, h3 {letter-spacing: -0.01em;}
    .small-note {font-size: 0.88rem; opacity: 0.78;}
    .status-box {
        border: 1px solid rgba(120,120,120,0.22);
        border-radius: 12px;
        padding: 0.9rem 1rem;
        margin-bottom: 0.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def fmt_money(value, currency):
    try:
        if value is None or np.isnan(value):
            return "n/a"
    except Exception:
        pass
    value = float(value)
    a = abs(value)
    if a >= 1e9:
        return f"{currency} {value/1e9:,.2f} B"
    if a >= 1e6:
        return f"{currency} {value/1e6:,.2f} M"
    if a >= 1e3:
        return f"{currency} {value/1e3:,.1f} k"
    return f"{currency} {value:,.2f}"


def fmt_num(value, digits=3):
    try:
        if value is None or np.isnan(value):
            return "n/a"
    except Exception:
        pass
    return f"{float(value):,.{digits}f}"


def fmt_pct(value, digits=1):
    try:
        if value is None or np.isnan(value):
            return "n/a"
    except Exception:
        pass
    return f"{float(value)*100:.{digits}f}%"


def dataframe(df, height=420):
    if df is None or df.empty:
        st.info("No results are available for this module with the current input workbook.")
        return
    st.dataframe(df, use_container_width=True, height=height, hide_index=True)


def show_image(result, key, caption=None):
    path = result.figures.get(key)
    if path and path.exists():
        st.image(str(path), caption=caption, use_container_width=True)
        return True
    return False


def indicator_value(df, indicator):
    if df is None or df.empty:
        return np.nan
    first_col = df.columns[0]
    rows = df[df[first_col].astype(str).str.strip().str.lower() == indicator.lower()]
    if rows.empty:
        return np.nan
    return rows.iloc[0, 1]


def decision_value(result, indicator, default="Not evaluated"):
    df = result.decision_summary
    if df is None or df.empty:
        return default
    rows = df[df["Decision Indicator"].astype(str).str.lower() == indicator.lower()]
    if rows.empty:
        return default
    value = rows.iloc[0]["Value"]
    if pd.isna(value):
        return default
    return str(value)


def make_download_zip(result):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as z:
        if result.results_excel.exists():
            z.write(result.results_excel, result.results_excel.name)
        if result.word_report.exists():
            z.write(result.word_report, result.word_report.name)
        for path in result.figures.values():
            if path.exists():
                z.write(path, path.name)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------
# Sidebar: case selection / upload
# ---------------------------------------------------------------------
with st.sidebar:
    st.title("Universal TEA")
    st.caption("Decision-Support Software · UI V0.1 · Engine V0.6")
    st.divider()

    source_mode = st.radio(
        "Start analysis from",
        ["Demo / validation case", "Upload universal TEA workbook"],
        help="The calculation engine always reads the same Universal TEA workbook structure.",
    )

    selected_path = None
    selected_label = None

    if source_mode == "Demo / validation case":
        selected_label = st.selectbox("Select case", list(DEMO_CASES.keys()))
        selected_path = DEMO_CASES[selected_label]

        with open(selected_path, "rb") as f:
            st.download_button(
                "Download selected input workbook",
                data=f.read(),
                file_name=selected_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        uploaded = st.file_uploader(
            "Upload Universal TEA .xlsx",
            type=["xlsx"],
            accept_multiple_files=False,
        )
        if uploaded is not None:
            upload_dir = Path(tempfile.mkdtemp(prefix="tea_upload_"))
            selected_path = upload_dir / uploaded.name
            selected_path.write_bytes(uploaded.getvalue())
            selected_label = uploaded.name

    run_clicked = st.button(
        "Run TEA Analysis",
        type="primary",
        use_container_width=True,
        disabled=selected_path is None,
    )

    st.divider()
    st.caption(
        "The interface does not replace the calculation model. "
        "It calls the frozen V0.6 workbook-driven engine and presents its outputs."
    )


# ---------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------
st.title("Universal Techno-Economic Analysis & Decision Support")
st.write(
    "A workbook-driven research tool for **CAPEX/OPEX**, discounted cash flow, "
    "**NPV–IRR–payback–MSP**, sensitivity, scenarios, break-even, scale, "
    "uncertainty, R&D prioritization, environmental screening and external validation."
)

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown("**Calculation core**  \nUniversal TEA Engine V0.6")
with c2:
    st.markdown("**Current interface**  \nSoftware UI V0.1")
with c3:
    st.markdown("**Workflow**  \nWorkbook → Engine → Dashboard → Reports")

if run_clicked:
    try:
        with st.status("Running Universal TEA V0.6…", expanded=True) as status:
            st.write("Reading and validating workbook inputs…")
            run_dir = Path(tempfile.mkdtemp(prefix="tea_ui_run_"))
            result = run_analysis(selected_path, run_dir)
            st.write("Calculating TEA and decision-analysis modules…")
            st.write("Generating Excel results, Word report and high-resolution figures…")
            st.session_state["tea_result"] = result
            st.session_state["tea_source_label"] = selected_label
            status.update(label="Analysis completed", state="complete", expanded=False)
    except Exception as exc:
        st.session_state.pop("tea_result", None)
        st.error("The analysis could not be completed.")
        st.exception(exc)

result = st.session_state.get("tea_result")

if result is None:
    st.divider()
    st.subheader("How to start")
    a, b, c = st.columns(3)
    with a:
        st.markdown(
            "### 1 · Select inputs\n"
            "Choose the generic demo, one of the validated DOE/NREL cases, "
            "or upload your own Universal TEA workbook."
        )
    with b:
        st.markdown(
            "### 2 · Run analysis\n"
            "The interface sends the workbook to the V0.6 engine. "
            "No TEA equations are reimplemented in the UI."
        )
    with c:
        st.markdown(
            "### 3 · Review & export\n"
            "Use the dashboard and analysis tabs, then download the complete "
            "Excel results, Word decision report and journal-quality figures."
        )
    st.info("Select a case in the left sidebar and click **Run TEA Analysis**.")
    st.stop()


# ---------------------------------------------------------------------
# Analysis header + warnings
# ---------------------------------------------------------------------
st.divider()
st.subheader(result.case_name)
st.caption(
    f"{result.process_category} · Primary product: {result.primary_product} · "
    f"Input source: {st.session_state.get('tea_source_label','')}"
)

if result.warnings:
    with st.expander(f"Methodology / input warnings ({len(result.warnings)})", expanded=False):
        for warning in result.warnings:
            st.warning(warning)


# ---------------------------------------------------------------------
# Main KPI dashboard
# ---------------------------------------------------------------------
baseline = result.baseline
currency = result.currency

st.markdown("### Executive dashboard")

r1 = st.columns(4)
r1[0].metric("Total capital investment", fmt_money(baseline["TCI"], currency))
r1[1].metric("Annual OPEX", fmt_money(baseline["Annual OPEX"], currency))
r1[2].metric("Annual revenue", fmt_money(baseline["Annual Revenue"], currency))
r1[3].metric("NPV", fmt_money(baseline["NPV"], currency))

r2 = st.columns(4)
r2[0].metric("IRR", fmt_pct(baseline["IRR"]))
r2[1].metric(
    "Payback",
    "Not reached" if pd.isna(baseline["Payback"]) else f"{baseline['Payback']:.2f} y",
)
r2[2].metric("MSP / LCOX", f"{fmt_num(baseline['MSP'],4)} {currency}/unit")
r2[3].metric("Benefit-cost ratio", fmt_num(baseline["BCR"], 3))

robustness = decision_value(result, "Economic robustness class")
top_sens = decision_value(result, "Top deterministic sensitivity")
top_rnd = decision_value(result, "Top R&D priority")

st.markdown(
    f"""
    <div class="status-box">
    <b>Decision summary:</b> {robustness}<br>
    <b>Most influential deterministic parameter:</b> {top_sens}<br>
    <b>Top R&amp;D priority:</b> {top_rnd}
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------
tabs = st.tabs([
    "Overview",
    "CAPEX & OPEX",
    "Financial",
    "Sensitivity",
    "Scenarios",
    "Break-even & Scale",
    "Uncertainty",
    "R&D Priority",
    "Environmental",
    "Validation",
    "Downloads",
])

with tabs[0]:
    st.markdown("#### Consolidated decision indicators")
    dataframe(result.decision_summary, height=500)

    left, right = st.columns(2)
    with left:
        show_image(result, "capex", "Installed CAPEX breakdown")
    with right:
        show_image(result, "opex", "Annual OPEX breakdown")

with tabs[1]:
    left, right = st.columns(2)
    with left:
        st.markdown("#### CAPEX")
        if not show_image(result, "capex"):
            st.info("No CAPEX figure is available.")
    with right:
        st.markdown("#### OPEX")
        if not show_image(result, "opex"):
            st.info("No OPEX figure is available.")

    st.markdown("#### Equipment / capital-cost details")
    equipment = result.baseline.get("Equipment Data")
    if isinstance(equipment, pd.DataFrame):
        dataframe(equipment)
    else:
        st.caption("Detailed equipment rows are available in the downloadable results Excel.")

with tabs[2]:
    st.markdown("#### Cumulative project cash flow")
    show_image(result, "cash_flow")
    st.markdown("#### Year-by-year cash flow")
    cashflow = result.baseline.get("Cash Flow")
    if isinstance(cashflow, pd.DataFrame):
        dataframe(cashflow, height=460)

with tabs[3]:
    if result.sensitivity.empty:
        st.info("Sensitivity analysis is not active in this workbook.")
    else:
        show_image(result, "sensitivity")
        show_image(result, "sensitivity_tornado")
        st.markdown("#### Sensitivity results")
        dataframe(result.sensitivity, height=520)

with tabs[4]:
    if result.scenarios.empty:
        st.info("Scenario analysis is not active in this workbook.")
    else:
        show_image(result, "scenarios")
        dataframe(result.scenarios, height=450)

with tabs[5]:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Break-even")
        show_image(result, "break_even")
        dataframe(result.break_even, height=360)
    with col2:
        st.markdown("#### Scale analysis")
        show_image(result, "scale")
        dataframe(result.scale, height=360)

with tabs[6]:
    if result.uncertainty_summary.empty:
        st.info(
            "Monte Carlo uncertainty is disabled for this workbook. "
            "Enable it in `Uncertainty_Settings` and activate distributions in "
            "`Uncertainty_Parameters`."
        )
    else:
        p_npv = indicator_value(result.uncertainty_summary, "P(NPV > 0)")
        p_irr = indicator_value(result.uncertainty_summary, "P(IRR >= target return)")
        p_msp = indicator_value(result.uncertainty_summary, "P(MSP <= entered selling price)")
        ucols = st.columns(3)
        ucols[0].metric("P(NPV > 0)", fmt_pct(p_npv))
        ucols[1].metric("P(IRR ≥ target)", fmt_pct(p_irr))
        ucols[2].metric("P(MSP ≤ selling price)", fmt_pct(p_msp))

        a, b = st.columns(2)
        with a:
            show_image(result, "uncertainty_npv")
        with b:
            show_image(result, "uncertainty_msp")
        show_image(result, "uncertainty_drivers")

        st.markdown("#### Probabilistic summary")
        dataframe(result.uncertainty_summary, height=520)
        st.markdown("#### Uncertainty drivers")
        dataframe(result.uncertainty_drivers, height=400)

with tabs[7]:
    if result.rd_priority.empty:
        st.info(
            "R&D priority analysis requires active uncertainty parameters in the input workbook."
        )
    else:
        show_image(result, "rd_priority")
        dataframe(result.rd_priority, height=520)
        st.caption(
            "Risk Priority and R&D Priority are transparent software-defined decision-support "
            "indices, not published industry-standard thresholds."
        )

with tabs[8]:
    if result.environmental.empty:
        st.info(
            "Environmental screening is not active. Add and enable boundary-consistent factors "
            "in the Environmental sheet before using this module."
        )
    else:
        show_image(result, "environmental")
        dataframe(result.environmental, height=500)
        st.caption("Environmental results are screening indicators and are not a full LCA.")

with tabs[9]:
    if result.validation.empty:
        st.info(
            "No external validation targets are active in this workbook. "
            "Use either validation demo case to show published-vs-calculated performance."
        )
    else:
        show_image(result, "validation")
        dataframe(result.validation, height=520)

with tabs[10]:
    st.markdown("#### Download complete outputs")
    d1, d2, d3 = st.columns(3)

    with d1:
        st.download_button(
            "Download Results Excel",
            data=result.results_excel.read_bytes(),
            file_name=result.results_excel.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d2:
        st.download_button(
            "Download Word Report",
            data=result.word_report.read_bytes(),
            file_name=result.word_report.name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with d3:
        st.download_button(
            "Download Complete Output ZIP",
            data=make_download_zip(result),
            file_name=f"{Path(result.results_excel).stem}_complete_outputs.zip",
            mime="application/zip",
            use_container_width=True,
        )

    st.markdown("#### Generated journal-quality figures")
    if result.figures:
        for key, path in sorted(result.figures.items()):
            with open(path, "rb") as f:
                st.download_button(
                    f"Download {path.name}",
                    data=f.read(),
                    file_name=path.name,
                    mime="image/png",
                    key=f"download_{key}_{path.name}",
                )
    else:
        st.info("No figures were generated for this case.")

st.divider()
st.caption(
    "Universal TEA Software UI V0.1 · Calculation Engine V0.6 · "
    "Research decision-support prototype. Review assumptions and source provenance before publication or investment use."
)
