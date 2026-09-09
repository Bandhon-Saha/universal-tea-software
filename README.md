# Universal TEA Decision-Support Software — UI V0.1

This is the first graphical software interface around the validated **Universal TEA Engine V0.6**.

## What the app does

The interface accepts the same Universal TEA Excel workbook used by the Python engine and presents:

- CAPEX and OPEX
- annual revenue
- NPV
- IRR
- payback
- MSP / LCOX
- benefit-cost ratio
- cash flow
- sensitivity analysis
- scenarios
- break-even analysis
- scale analysis
- Monte Carlo uncertainty
- probabilistic viability
- uncertainty drivers
- Risk Priority and R&D Priority
- environmental screening
- external validation
- automatic Excel results
- automatic Word report
- downloadable journal-quality PNG figures

The interface **does not duplicate or replace the TEA equations**. It calls the frozen V0.6 calculation engine under `engine/`.

---

## Windows — first run

### Step 1 — install Python

Install a recent 64-bit Python 3 release from Python.org.

During installation, enable the option to add Python to PATH if offered.

### Step 2 — unzip this software folder

Keep all files and folders together.

### Step 3 — install the requirements once

Double-click:

`SETUP_WINDOWS.bat`

Wait until it says setup is complete.

### Step 4 — launch the software

Double-click:

`RUN_UNIVERSAL_TEA.bat`

A terminal window will remain open while the application is running. Your browser should open the local Streamlit interface.

If the browser does not open automatically, use the local address printed in the terminal, normally:

`http://localhost:8501`

Do not close the terminal while using the application.

---

## Manual launch

From a terminal opened inside the software folder:

```bash
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

---

## Suggested university demonstration

1. Open the software.
2. Select **Validation Case 1 — H2A biomass gasification**.
3. Click **Run TEA Analysis**.
4. Show the executive KPI dashboard.
5. Open **Validation** and show published vs calculated deviation.
6. Open **Sensitivity**, **Scenarios**, and **Break-even & Scale**.
7. Return to the generic demo.
8. Show **Uncertainty** and **R&D Priority**.
9. Open **Downloads** and download the Excel results / Word report.

---

## Included input workbooks

`templates/Universal_TEA_Input_Workbook_V0_6.xlsx`  
Generic demonstration with uncertainty enabled.

`templates/H2A_Biomass_Validation.xlsx`  
DOE/NREL biomass-to-hydrogen validation case.

`templates/PEM_Electrolysis_Validation.xlsx`  
DOE/NREL central PEM-electrolysis validation case.

---

## Current prototype scope

UI V0.1 is a local Streamlit research-software prototype. It is intentionally workbook-driven so the interface and the validated calculation model remain separate.

V0.6 explicitly treats Monte Carlo input distributions, robustness classes, Risk Priority and R&D Priority as transparent software-defined decision-support assumptions. Environmental results are screening indicators, not a full LCA.

Debt financing and equipment degradation remain planned refinements in the underlying analytical engine; they should not be presented as fully implemented calculations in the current software.
