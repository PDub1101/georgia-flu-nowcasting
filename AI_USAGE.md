# AI-Agent and Reproducibility Audit: Homework 4 (Course Challenge Round 3)

**Course**: ISyE 4031: Regression and Forecasting — Fall 2026  
**Assignment**: Homework 4 & Course Challenge Round 3: Retrospective and Prospective Nowcasts  
**Team Name**: Gibson, Carnes, Lutz (`Gibson_Carnes_Lutz`)  
**Team Members**: Paul Gibson, Ava Carnes, Abigail Lutz (Section T08)  
**Location**: Georgia (Location Code `13`)  
**Due Date**: Tuesday, October 13, 2026 at 11:59 p.m. Eastern Time  

---

## 1. Tool and Model Designation
- **Agentic Environment**: Antigravity Agentic IDE
- **Underlying Large Language Model**: Gemini 3.8 Flash (High)
- **Runtime Environment**: Windows 11 / PowerShell / Python 3.11+

---

## 2. Substantive Tasks Delegated
We delegated the automated data retrieval, computational pipeline, and document compilation to the agent:
1. **Automated Cutoff Quarantining & Data Archiving**: Programmatic construction of physically separated cutoff files in `data/`, local caching of the NOAA NCEI Daily Summaries API (`station: USW00013874`), retrieval of dated FluSight GitHub commit snapshots (`0eff5838f5`, `5268daf5e6`, `38023fd5e0`), and parsing of CDC NHSN preliminary (`mpgq-jmmr`) and official (`ua7e-t2fy`) surveillance feeds.
2. **First-Principles Statistical Pipeline**: Implementation of ordinary least squares (OLS) matrix equations, residual standard error derivation ($s = 225.7383$, $df = 208$), explicit mathematical contrast between $SE_{\text{pred}}$ (single observation, $\approx 226.5$) and $SE_{\text{mean}}$ (fitted regression line, $\approx 16.9$), Student's $t$ prediction intervals (50%, 80%, 95%), non-negativity clipping, and automated nesting assertions.
3. **Reproducibility & Deliverable Emission**: Automated emission of `forecast.csv` (1 prospective row, Point = 101) and `retrospective_forecasts.csv` (2 retrospective rows, Points = 61, 71) adhering to the course schema with `team_name = Gibson_Carnes_Lutz` (no internal commas).
4. **Figure & Report Compilation**: Generation of high-resolution vector figures (`fig1_nowcast_timeline.png` and `fig2_seasonal_heteroskedasticity.png`), programmatic construction of `report.docx` via `python-docx`, and compilation to a 5-page `report.pdf` via Microsoft Word COM automation.

---

## 3. Suggestion Accepted or Revised
- **Agent Suggestion**: The agent derived raw point extrapolations ($\hat{y}_{\text{retro1}} = 60.93$, $\hat{y}_{\text{retro2}} = 71.45$, $\hat{y}_{\text{prosp}} = 101.27$) and continuous lower prediction bounds that extended substantially below zero (e.g., raw 95% lower bounds: $-385.66$, $-375.04$, and $-345.00$).
- **Team Revision**: We accepted the mathematical derivation of $SE_{\text{pred}}$, but revised the operational CSV emission to enforce discrete integer rounding and non-negativity truncation ($\max(0, \cdot)$) for both the point nowcasts and interval bounds ($L_{95} = 0, L_{80} = 0, L_{50} = 0$). While the continuous negative numbers are mathematically implied by the linear regression plane, reporting negative hospital admissions violates physical reality and the course submission contract. Both unrounded raw results and clipped discrete submission values are fully disclosed in the technical report.

---

## 4. Suggestion Rejected with Statistical Reason
- **Agent Suggestion**: To resolve the severe overstatement of uncertainty in late-summer weeks (where the global $s = 225.74$ produces massive $\approx 893$-admission prediction intervals) and avoid negative lower bounds, the agent suggested abandoning the simple linear regression baseline and refitting an expanding-window log-transformed model ($\log(Y_t + 1)$) or a polynomial spline model using the newly released September surveillance data.
- **Statistical Reason for Rejection**:
  1. *Benchmark Contract Integrity*: Section 3 of the course handout explicitly guarantees that *"A team using the original frozen reference may keep it throughout and earn full credit."* Abandoning the frozen benchmark destroys longitudinal comparability across Course Challenge Rounds 1, 2, and 3.
  2. *Prevention of Hindsight Cherry-Picking*: Adopting an alternative functional form after knowing that late-summer influenza activity remained subdued would introduce subtle hindsight bias.
  3. *Transformation Bias & Exponentiation Hazards*: Log-transforming hospital admissions on extreme seasonal counts introduces severe retransformation bias ($\mathbb{E}[e^{\epsilon}] \neq e^{\mathbb{E}[\epsilon]}$) and compresses peak winter volatility, understating epidemic tail risk during cold snaps. We maintained `temperature_slr_v1` as the official submission model, while presenting the expanding-window refit as a transparent sensitivity analysis in Table 3.

---

## 5. Independent Numerical and Graphical Checks Performed
We performed the following independent checks without relying on agent-generated claims:
1. **ANOVA Partition Identity**: We independently verified the fundamental ANOVA partition identity $SST = SSR + SSE$ from raw vector sums:
   $$SST = 15,300,755.57 = SSR (4,701,540.10) + SSE (10,599,215.47) \quad (\text{absolute difference} < 10^{-8})$$
2. **Monotonicity & Nesting Verification**: Confirmed that prediction interval widths strictly expand with confidence level ($W_{50} = 305.91 < W_{80} = 582.07 < W_{95} = 892.54$) and verified strict non-negative nesting:
   - Retro 1: $0 \le 0 \le 0 \le 0 \le 61 \le 214 \le 352 \le 508$ (PASSED)
   - Retro 2: $0 \le 0 \le 0 \le 0 \le 71 \le 224 \le 363 \le 518$ (PASSED)
   - Prospective: $0 \le 0 \le 0 \le 0 \le 101 \le 254 \le 392 \le 548$ (PASSED)
3. **Manual Weather Aggregation & Missing-Day Audit**: We independently audited the temperature calculations:
   - Sept 20–26 (7 days): $519.0^\circ\text{F} / 7 = 74.1429^\circ\text{F}$
   - Sept 27–Oct 3 (7 days): $512.5^\circ\text{F} / 7 = 73.2143^\circ\text{F}$
   - Oct 4–10 (6 available days in NOAA NCEI, Saturday = `NA`): $423.5^\circ\text{F} / 6 = 70.5833^\circ\text{F}$
4. **Offline Run Verification**: Executed `python analysis.py` with Wi-Fi/network connectivity disabled to verify 100% offline self-containment.
