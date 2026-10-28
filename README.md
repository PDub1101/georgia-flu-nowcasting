# ISyE 4031 Homework 4 & Course Challenge Round 3: Retrospective and Prospective Nowcasts

**Course**: ISyE 4031: Regression and Forecasting — Fall 2026  
**Assignment**: Homework 4 (Course Challenge Round 3: Retrospective and Prospective Nowcasts)  
**Team Name**: Gibson, Carnes, Lutz (`Gibson_Carnes_Lutz`)  
**Team Members**: Paul Gibson, Ava Carnes, Abigail Lutz (Section T08)  
**Location**: Georgia (Location Code `13`, FIPS `13`)  
**Official Model ID**: `temperature_slr_v1`  
**Training Data Vintage**: `hw2_2026-09-09` (Frozen Homework 2 course history, $n = 210$ weeks ending `2026-08-29`)  
**Prospective Issue Timestamp**: `2026-10-12T21:15:00-04:00` (Monday, October 12, 2026 at 9:15 p.m. Eastern Time)  
**Submission Due Date**: Tuesday, October 13, 2026 at 11:59 p.m. Eastern Time  

---

## 1. Exact Offline Reproduction Command

To reproduce the complete pipeline, perform all numerical integrity and nesting validations, recreate the timeline and diagnostic figures, emit both submission CSV files, and print evaluation metrics completely offline from local saved inputs in `data/`, run:

```bash
python analysis.py
```

To recompile `report.docx` and regenerate the 5-page `report.pdf` using Microsoft Word COM automation:

```bash
python make_report_docx.py
```

---

## 2. Software Package & Environment Record

This analysis was developed and audited under Windows 11 with the following software environment:

| Software / Package | Version | Primary Purpose in Reproduction |
| :--- | :--- | :--- |
| **Python** | 3.11+ / 3.13 | Main script runtime environment (`analysis.py`, `make_report_docx.py`) |
| **NumPy** | 1.26.4 / 2.x | First-principles matrix algebra, ANOVA decomposition, and quantiles |
| **Pandas** | 2.2.2 | Cutoff-quarantined dataset loading, CSV emission, and assertion checks |
| **Matplotlib** | 3.8.4 | Diagnostic figures and publication timeline graphic generation |
| **python-docx** | 1.1.2 | Programmatic compilation of publication-quality `report.docx` |
| **pywin32 / Word COM** | 306+ | Automated lossless PDF export and page-count verification |

---

## 3. Data-Vintage Identifiers & Provenance Record

| Stream / Dataset | Vintage Identifier | Time Window Covered | Local Archived Path |
| :--- | :--- | :--- | :--- |
| **Frozen Course Baseline** | `hw2_2026-09-09` | 2022-02-19 to 2026-08-29 (210 weeks) | `data/hw2_ga_flu_temperature_history_2026-09-09.csv` |
| **FluSight Cutoff 1 Snapshot** | `flusight_0eff5838f5` | Through 2026-07-04 (pre-Cutoff 1) | `data/flusight_ga_snapshot_2026-09-24_0eff5838f5.csv` |
| **FluSight Cutoff 2 Snapshot** | `flusight_5268daf5e6` | Through 2026-09-26 (initial Sept 26 = 28) | `data/flusight_ga_snapshot_2026-09-30_5268daf5e6.csv` |
| **CDC NHSN Preliminary** | `cdc_prelim_2026-10-07` | Through 2026-10-03 (revised Sept 26 = 30) | `data/raw_cdc_nhsn_prelim_ga_2026-10-07.json` |
| **CDC NHSN Official** | `cdc_official_2026-10-09` | Through 2026-10-03 (finalized Sept 26 = 31) | `data/raw_cdc_nhsn_official_ga_2026-10-09.json` |
| **NOAA Daily Summaries** | `noaa_katl_hw4_2026-10-12` | 2026-09-13 to 2026-10-09 (Sat Oct 10 = `NA`) | `data/noaa_daily_katl_2026-09-13_to_2026-10-10.csv` |

---

## 4. Run Execution Timestamps, Predictor Links & Retrospective Labels

Every row in `retrospective_forecasts.csv` and `forecast.csv` is linked below to its training vintage, local predictor snapshot file, simulated/actual cutoff, actual generation timestamp (never backdated), and exact/approximate retrospective classification:

| Output CSV File | `forecast_date` (Cutoff/Issue) | `target_date` (Week Ending) | `model_id` | Training `data_vintage` & Local File | Predictor Snapshot ID & Local File | Actual Generation Timestamp (EDT) | Classification Label | Submitted Values (`pt`, 50%, 80%, 95%) |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`retrospective_forecasts.csv`** (Row 1) | `2026-09-29` (Sim. Cutoff: `2026-09-29T23:59:00-04:00`) | `2026-09-26` | `temperature_slr_v1` | `hw2_2026-09-09` (`data/cutoff_2026-09-29_training.csv`) | `noaa_katl_2026-09-26_retro_v20261012` (`data/cutoff_2026-09-29_predictor_2026-09-26.csv`, $T_0 = 74.1429^\circ\text{F}$, 7/7 days) | `2026-10-12T21:54:21-04:00` | **`approximate retrospective`** | `61`, `[0, 214]`, `[0, 352]`, `[0, 508]` |
| **`retrospective_forecasts.csv`** (Row 2) | `2026-10-06` (Sim. Cutoff: `2026-10-06T23:59:00-04:00`) | `2026-10-03` | `temperature_slr_v1` | `hw2_2026-09-09` (`data/cutoff_2026-10-06_training.csv`) | `noaa_katl_2026-10-03_retro_v20261012` (`data/cutoff_2026-10-06_predictor_2026-10-03.csv`, $T_0 = 73.2143^\circ\text{F}$, 7/7 days) | `2026-10-12T21:54:21-04:00` | **`approximate retrospective`** | `71`, `[0, 224]`, `[0, 363]`, `[0, 518]` |
| **`forecast.csv`** (Row 1) | `2026-10-12` (Actual Issue: `2026-10-12T21:15:00-04:00`) | `2026-10-10` | `temperature_slr_v1` | `hw2_2026-09-09` (`data/cutoff_2026-10-12_training.csv`) | `noaa_katl_2026-10-10_prosp_v20261012` (`data/cutoff_2026-10-12_predictor_2026-10-10.csv`, $T_0 = 70.5833^\circ\text{F}$, 6/7 days; Sat Oct 10 = `NA`) | `2026-10-12T21:54:21-04:00` | **`prospective nowcast`** (Genuinely Unseen) | `101`, `[0, 254]`, `[0, 392]`, `[0, 548]` |

- **Labeling Rationale**: Both retrospective runs are designated **`approximate retrospective`** because daily weather observations were retrieved post-hoc on October 12, 2026, even though strict information isolation and the frozen August 29 baseline were maintained. The prospective run is **`prospective nowcast`** because it was issued before the October 10 Georgia hospital admission count was observed or released by CDC/NHSN.

---

## 5. Mathematical Results & SE Contrast

- **Baseline Equation**: $\hat{Y}_t = 901.3096 - 11.3346 \cdot T_t \quad (s = 225.7383, \ R^2 = 0.3073, \ n = 210)$
- **SE Contrast (Prediction vs Confidence Interval)**:
  - *Retrospective 1*: $SE_{\text{pred}} = 226.53$ vs. $SE_{\text{mean}} = 18.90$ (Mean CI: $[23.7, 98.2]$, width = $74.5$)
  - *Retrospective 2*: $SE_{\text{pred}} = 226.48$ vs. $SE_{\text{mean}} = 18.30$ (Mean CI: $[35.4, 107.5]$, width = $72.2$)
  - *Prospective*: $SE_{\text{pred}} = 226.37$ vs. $SE_{\text{mean}} = 16.88$ (Mean CI: $[68.0, 134.6]$, width = $66.6$)
- **Retrospective 1 (2026-09-26)**: $T_0 = 74.1429^\circ\text{F} \rightarrow \hat{y} = 61, \ 50\%: [0, 214], \ 80\%: [0, 352], \ 95\%: [0, 508]$
  - *Observed Count*: 31 admissions (revision: $28 \rightarrow 30 \rightarrow 31$) $\rightarrow$ Error: $+30.0$, Absolute Error: $30.0$ admissions. All intervals capture observed count.
- **Retrospective 2 (2026-10-03)**: $T_0 = 73.2143^\circ\text{F} \rightarrow \hat{y} = 71, \ 50\%: [0, 224], \ 80\%: [0, 363], \ 95\%: [0, 518]$
  - *Observed Count*: 52 admissions (revision: $49 \rightarrow 52$) $\rightarrow$ Error: $+19.0$, Absolute Error: $19.0$ admissions. All intervals capture observed count.
- **Prospective (2026-10-10)**: $T_0 = 70.5833^\circ\text{F} \rightarrow \hat{y} = 101, \ 50\%: [0, 254], \ 80\%: [0, 392], \ 95\%: [0, 548]$
  - *Observed Count*: Unreleased at issuance (evaluation pending).
- **Strict Nesting Assertion**:
  $$0 \le L_{95} \le L_{80} \le L_{50} \le \hat{y} \le U_{50} \le U_{80} \le U_{95} \quad \text{(PASSED FOR ALL 3 WEEKS)}$$

---

## 6. Contact Information

- **Paul Gibson**: `pgibson@gatech.edu`
- **Ava Carnes**: `acarnes@gatech.edu`
- **Abigail Lutz**: `alutz@gatech.edu`
- **Section**: ISyE 4031 T08, Fall 2026
