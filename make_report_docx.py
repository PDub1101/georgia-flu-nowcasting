"""
Generate publication-quality report.docx and compile to report.pdf for Homework 4.
ISyE 4031: Regression and Forecasting - Fall 2026
Team: Gibson, Carnes, Lutz (Section T08)
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCX_PATH = os.path.join(BASE_DIR, 'report.docx')
PDF_PATH = os.path.join(BASE_DIR, 'report.pdf')

doc = docx.Document()

# Page Margins: 0.5 in top/bottom, 0.6 in left/right
for s in doc.sections:
    s.top_margin = Inches(0.48)
    s.bottom_margin = Inches(0.48)
    s.left_margin = Inches(0.58)
    s.right_margin = Inches(0.58)

NAVY = RGBColor(0, 48, 87)       # Georgia Tech Navy #003057
GOLD = RGBColor(179, 163, 105)   # Georgia Tech Gold #B3A369
CHARCOAL = RGBColor(35, 35, 35)
GRAY = RGBColor(90, 90, 90)
CRIMSON = RGBColor(160, 20, 20)

def set_cell_background(cell, fill_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=30, bottom=30, left=60, right=60):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_p(text='', space_after=2.0, space_before=0.0, line_spacing=1.04, bold=False, italic=False, color=CHARCOAL, size=8.8):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return p

def add_section_header(title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5.5)
    p.paragraph_format.space_after = Pt(2.0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(title)
    r.font.name = 'Calibri'
    r.font.size = Pt(10.8)
    r.font.bold = True
    r.font.color.rgb = NAVY
    bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="B3A369"/></w:pBdr>')
    p._p.get_or_add_pPr().append(bdr)

def add_sub_header(title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3.5)
    p.paragraph_format.space_after = Pt(1.0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(title)
    r.font.name = 'Calibri'
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = NAVY

# -----------------------------------------------------------------------------
# Document Header
# -----------------------------------------------------------------------------
title_p = doc.add_paragraph()
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after = Pt(1)
r_title = title_p.add_run('ISyE 4031: Regression and Forecasting — Fall 2026')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(13.5)
r_title.font.bold = True
r_title.font.color.rgb = NAVY

sub_p = doc.add_paragraph()
sub_p.paragraph_format.space_before = Pt(0)
sub_p.paragraph_format.space_after = Pt(3)
r_sub = sub_p.add_run('Course Challenge Round 3: Retrospective and Prospective Nowcasts (Homework 4)\n')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(10.2)
r_sub.font.bold = True
r_sub.font.color.rgb = RGBColor(140, 115, 45)

r_meta = sub_p.add_run('Team: Gibson, Carnes, Lutz (Team ID: Gibson_Carnes_Lutz)  |  Authors: Paul Gibson, Ava Carnes, Abigail Lutz (Section T08)\n'
                       'Location: Georgia (Code 13)  |  Due Date: October 13, 2026  |  Model ID: temperature_slr_v1  |  Vintage: hw2_2026-09-09')
r_meta.font.name = 'Calibri'
r_meta.font.size = Pt(8.0)
r_meta.font.italic = True
r_meta.font.color.rgb = GRAY

# Header divider
p_div = doc.add_paragraph()
p_div.paragraph_format.space_after = Pt(3)
p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="1" w:color="003057"/></w:pBdr>')
p_div._p.get_or_add_pPr().append(p_div_border)

# -----------------------------------------------------------------------------
# 1. THREE TARGETS, THREE INFORMATION CUTOFFS & PROTOCOL AUDIT
# -----------------------------------------------------------------------------
add_section_header('1. Three Targets, Information Cutoffs & Forecasting Contract')
add_p('In Homework 4 (Course Challenge Round 3), our team reconstructs two missed weekly Georgia influenza nowcasts and issues a prospective nowcast for the current epidemiological week. All nowcasts predict weekly new laboratory-confirmed influenza hospital admissions across Georgia (Location Code 13, CDC field totalconfflunewadm). Under course challenge protocols, we strictly segregate historical information cutoffs, store separate cutoff-quarantined datasets in data/, freeze all three sets of predictions prior to inspecting retrospective outcomes, and evaluate retrospective reliability.')

cutoff_table_data = [
    ["Run Type", "Target Week & Date", "Information Cutoff (Eastern)", "Legally Available Surveillance", "Excluded / Prohibited Information", "Operational Status & Data File"],
    ["Retrospective 1", "Sept 20–26, 2026\n(2026-09-26)", "Simulated: Sept 29, 2026\n11:59 p.m. ET", "Daily weather through Sept 26 (7/7 days); training admissions through Aug 29.", "Sept 26 target admissions count and all CDC releases after Sept 29.", "Approximate retrospective;\ncutoff_2026-09-29_predictor.csv"],
    ["Retrospective 2", "Sept 27–Oct 3, 2026\n(2026-10-03)", "Simulated: Oct 6, 2026\n11:59 p.m. ET", "Daily weather through Oct 3 (7/7 days); training admissions through Aug 29.", "Oct 3 target admissions count and all CDC releases after Oct 6.", "Approximate retrospective;\ncutoff_2026-10-06_predictor.csv"],
    ["Prospective", "Oct 4–10, 2026\n(2026-10-10)", "Actual Issue Time:\nOct 12, 2026 3:00 p.m. ET", "Daily weather through Oct 9 (6/7 days; Sat Oct 10 = NA in NOAA NCEI).", "Oct 10 target count (strictly unreleased/unseen at issuance).", "Genuinely prospective; frozen before evaluation; cutoff_2026-10-12_predictor.csv"]
]

t_cut = doc.add_table(rows=len(cutoff_table_data), cols=6)
t_cut.alignment = WD_TABLE_ALIGNMENT.CENTER
for col_idx, text in enumerate(cutoff_table_data[0]):
    cell = t_cut.rows[0].cells[col_idx]
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(7.8)
    r.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, '003057')
    set_cell_margins(cell, 20, 20, 40, 40)

col_widths_cut = [Inches(1.0), Inches(1.1), Inches(1.2), Inches(1.5), Inches(1.3), Inches(1.2)]
for row_idx, row in enumerate(cutoff_table_data[1:], start=1):
    for col_idx, val in enumerate(row):
        cell = t_cut.rows[row_idx].cells[col_idx]
        cell.width = col_widths_cut[col_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(7.6)
        bg = 'F8F9FA' if row_idx % 2 == 1 else 'FFFFFF'
        set_cell_background(cell, bg)
        set_cell_margins(cell, 20, 20, 40, 40)

add_p('', space_after=1)
add_p('Disclosure of Prior Knowledge: As required by Section 3, we candidly disclose that these retrospective runs were executed on October 12, 2026. While CDC published surveillance tables through October 3 on October 8–9, our computational pipeline programmatically isolated inputs into separate cutoff files in data/, prohibiting target counts or post-cutoff revisions from entering model fitting. All three predictions were completely frozen before opening released surveillance outcomes.', size=8.3, italic=True, color=GRAY)

# -----------------------------------------------------------------------------
# 2. DATA REFRESH, PROVENANCE & METEOROLOGICAL LOG
# -----------------------------------------------------------------------------
add_section_header('2. Data Refresh, Provenance & Recent History Documentation')
add_p('Recent Admissions & Weather History: We ingested Georgia NHSN influenza hospital admissions and Atlanta ambient surface temperature history covering the three completed weeks as of October 8 (weeks ending September 19, September 26, and October 3) and the prospective week ending October 10. In strict accordance with course rules, unreleased values are treated as unavailable (NA), never imputed as zero.')

add_sub_header('A. Target Week Daily Meteorological Logs (NOAA Station USW00013874, Atlanta KATL)')
add_p('Daily temperature records were retrieved from the designated NOAA NCEI Daily Summaries API for station USW00013874. For each day, midrange temperature T_mid = (TMAX + TMIN)/2 was computed. For Saturday, October 10, data remained unreleased in NOAA NCEI due to the standard 2–3 day publication lag. In strict compliance with Section 2, October 10 is recorded as NA (unavailable), and missing-day handling is executed by computing the mean of the 6 available days (Sunday through Friday), yielding T0 = 70.5833°F.')

temp_table_data = [
    ["Target Week Ending", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "7-Day Sum", "Weekly Mean (T0)", "Missing Days & Handling Policy"],
    ["2026-09-19 (HW3 Ref)", "81.0", "82.5", "82.5", "81.0", "79.5", "82.0", "81.5", "570.0°F", "81.43°F (81.57)", "0 missing; full 7-day GHCN verified."],
    ["2026-09-26 (Retro 1)", "80.0", "80.5", "79.0", "74.5", "66.0", "68.5", "70.5", "519.0°F", "74.1429°F", "0 missing; NOAA NCEI daily summaries (7/7 days)."],
    ["2026-10-03 (Retro 2)", "68.0", "67.5", "72.0", "74.5", "75.5", "77.0", "78.0", "512.5°F", "73.2143°F", "0 missing; NOAA NCEI daily summaries (7/7 days)."],
    ["2026-10-10 (Prospective)", "74.0", "71.5", "70.0", "69.5", "71.5", "67.0", "NA", "423.5°F", "70.5833°F", "1 missing (Sat Oct 10 NA in NCEI lag); 6-day mean evaluated."]
]

t_temp = doc.add_table(rows=len(temp_table_data), cols=11)
t_temp.alignment = WD_TABLE_ALIGNMENT.CENTER
for col_idx, text in enumerate(temp_table_data[0]):
    cell = t_temp.rows[0].cells[col_idx]
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(7.4)
    r.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, '003057')
    set_cell_margins(cell, 18, 18, 30, 30)

for row_idx, row in enumerate(temp_table_data[1:], start=1):
    for col_idx, val in enumerate(row):
        cell = t_temp.rows[row_idx].cells[col_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(7.3)
        bg = 'F8F9FA' if row_idx % 2 == 1 else 'FFFFFF'
        set_cell_background(cell, bg)
        set_cell_margins(cell, 18, 18, 30, 30)

add_p('', space_after=1)
add_sub_header('B. Compact Data Provenance, FluSight Git Hashes & Triple-Timestamp Log')
add_p('To enforce provenance integrity, our protocol distinguishes event occurrence, public release vintage, and local download timestamps, incorporating exact FluSight GitHub commit hashes and CDC Socrata endpoints:')

data_log_rows = [
    ["Dataset Name / Stream", "Source Entity & URL / Git Commit", "Dates Covered", "Release Status & Vintage", "Public Availability Date", "Team Download Timestamp"],
    ["Frozen Course Baseline", "ISyE 4031 Course Distribution\n(hw2_ga_flu_history.csv)", "2022-02-19 to 2026-08-29\n(210 observed weeks)", "Official Course Frozen Baseline\n(vintage: hw2_2026-09-09)", "2026-09-09 12:00 ET", "2026-09-09 Pre-archived\n(Local: data/hw2_ga_flu_history.csv)"],
    ["FluSight Cutoff 1 Snapshot", "github.com/cdcepi/FluSight-forecast-hub\nCommit: 0eff5838f5 (2026-09-24)", "Through 2026-07-04\n(latest commit pre-Cutoff 1)", "Official FluSight target-data commit\n(vintage: flusight_0eff5838f5)", "2026-09-24 18:52 UTC", "2026-10-12 21:01 EDT\nLocal: data/flusight_...0eff5838f5.csv"],
    ["FluSight Cutoff 2 Snapshot", "github.com/cdcepi/FluSight-forecast-hub\nCommit: 5268daf5e6 (2026-09-30)", "Through 2026-09-26\n(initial unrevised Sept 26 = 28)", "Official FluSight target-data commit\n(vintage: flusight_5268daf5e6)", "2026-09-30 19:49 UTC", "2026-10-12 21:01 EDT\nLocal: data/flusight_...5268daf5e6.csv"],
    ["CDC NHSN Preliminary", "data.cdc.gov/resource/mpgq-jmmr.json\nFluSight Commit: 38023fd5e0", "Through 2026-10-03\n(revised Sept 26 = 30, Oct 3 = 49)", "Preliminary weekly surveillance\n(vintage: cdc_prelim_2026-10-07)", "2026-10-07 16:17 UTC", "2026-10-12 14:53 EDT\nLocal: data/raw_cdc_nhsn_prelim_ga...json"],
    ["CDC NHSN Official", "data.cdc.gov/resource/ua7e-t2fy.json\n(jurisdiction=GA, totalconfflunewadm)", "Through 2026-10-03\n(finalized Sept 26 = 31, Oct 3 = 52)", "Official weekly surveillance\n(vintage: cdc_official_2026-10-09)", "2026-10-09 14:07 UTC", "2026-10-12 14:53 EDT\nLocal: data/raw_cdc_nhsn_official_ga...json"],
    ["NOAA Daily Summaries", "ncei.noaa.gov/access/services/data/v1\n(Station USW00013874)", "2026-09-13 to 2026-10-09\n(Daily TMAX, TMIN; Oct 10 = NA)", "Official GHCN-Daily summaries\n(vintage: noaa_katl_hw4_2026-10-12)", "Daily at ~06:00 ET\n(2-day publication lag)", "2026-10-12 14:48 EDT\nLocal: data/noaa_daily_katl_...csv"]
]

t_log = doc.add_table(rows=len(data_log_rows), cols=6)
t_log.alignment = WD_TABLE_ALIGNMENT.CENTER
for col_idx, text in enumerate(data_log_rows[0]):
    cell = t_log.rows[0].cells[col_idx]
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(7.4)
    r.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, '003057')
    set_cell_margins(cell, 18, 18, 30, 30)

for row_idx, row in enumerate(data_log_rows[1:], start=1):
    for col_idx, val in enumerate(row):
        cell = t_log.rows[row_idx].cells[col_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(7.2)
        bg = 'F8F9FA' if row_idx % 2 == 1 else 'FFFFFF'
        set_cell_background(cell, bg)
        set_cell_margins(cell, 18, 18, 30, 30)

# -----------------------------------------------------------------------------
# 3. RECONSTRUCTION METHODOLOGY & STATED MODELING DECISIONS
# -----------------------------------------------------------------------------
add_section_header('3. Reconstruction Methodology & Stated Modeling Decisions')
add_p('Reconstruction Plan Written Ex-Ante: As instructed by Section 3, our reconstruction plan was established prior to inspecting the two target outcomes. We start from the model and decision rule submitted in Homework 3:')

add_p('1. Stated Model & Fixed Recipe: We preserve our official baseline: model_id = temperature_slr_v1. The training window is held fixed at the 210 observed weeks ending August 29, 2026 (data_vintage = hw2_2026-09-09). Per Section 3 guidelines ("A team using the original frozen reference may keep it throughout and earn full credit"), keeping this frozen course reference across all three runs guarantees longitudinal benchmark integrity and ensures that our reconstructions are 100% immune to historical data revisions, reporting backfills, and vintage drift.')

add_p('2. Construction of Separate Legal Input Sets: In data/, we generated separate legal input files for each cutoff. For Retrospective 1 (cutoff Sept 29, 11:59 p.m. ET), target week admissions for September 20–26 were strictly prohibited; historical admissions beyond August 29 were excluded to avoid unverified revision artifacts. For Retrospective 2 (cutoff Oct 6, 11:59 p.m. ET), October 3 admissions were strictly excluded. For the Prospective run (issue time Oct 12), October 10 admissions are completely unobserved and unreleased.')

add_p('3. Vintage Integrity & Historical Revision Tracking: Querying dated FluSight Git commits revealed that CDC admissions undergo active multi-stage revisions: for the week ending September 26, the initial count was 28 (FluSight commit 5268daf5e6), revised to 30 in preliminary release (commit 38023fd5e0), and finalized at 31 in the official release (28 -> 30 -> 31). Filtering modern data to past dates would silently leak finalized values (31) into historical training sets. Our frozen August 29 baseline protects against revision leakage.')

add_p('4. Exact Mathematical Formulation & Prediction Intervals: For untransformed temperature regression, ordinary least squares parameter estimation yields: b0 = 901.3096, b1 = -11.3346, s = 225.7383 (df = 208). For any target week mean temperature T0, the point extrapolation is ŷ0 = b0 + b1 * T0. The standard error for one new observation retains the residual variation term: SE_pred = s * sqrt(1 + 1/n + (T0 - T_bar)^2 / Sxx). Central prediction intervals are constructed using Student\'s t critical values: t_50 = 0.675704, t_80 = 1.285621, and t_95 = 1.971435. Discrete submission values are generated by integer rounding and physical non-negativity clipping (Y >= 0).')

# -----------------------------------------------------------------------------
# 4. PROSPECTIVE NOWCAST & THREE-WEEK COMPARISON TABLE
# -----------------------------------------------------------------------------
add_section_header('4. Prospective Nowcast, CI vs. PI Contrast & Forecast Comparison')
add_p('Issuance of Prospective Nowcast: For the week ending October 10, 2026, ambient temperature across the 6 available days averaged T0 = 70.5833°F (Saturday Oct 10 = NA in NOAA NCEI). Evaluating our frozen model yields an unrounded point extrapolation of ŷ0 = 101.2729 admissions, with SE_pred = 226.3686. Applying integer rounding yields a point nowcast of 101 admissions, with nested prediction intervals of [0, 254] (50%), [0, 392] (80%), and [0, 548] (95%). Finite values and non-negative nesting (0 <= 0 <= 0 <= 0 <= 101 <= 254 <= 392 <= 548) are strictly verified.')

add_sub_header('A. Essential Terminology Distinction: Prediction Interval vs. Confidence Interval for the Mean Response')
add_p('Section 4 and Terminology guidelines require explicitly distinguishing a prediction interval for one new count from a confidence interval for the fitted mean response E[Y|T0]. The standard error of the fitted mean regression line is SE_mean = s * sqrt(1/n + (T0 - T_bar)^2 / Sxx). For the prospective week (T0 = 70.58°F), SE_mean = 16.88 admissions, yielding a 95% Confidence Interval of [68.0, 134.6] admissions (width = 66.6). This CI expresses sampling uncertainty in estimating the average historical relationship between temperature and admissions. In contrast, predicting a single weekly count requires accounting for individual weekly residual variation sigma^2, adding +1 under the radical: SE_pred = s * sqrt(1 + 1/n + (T0 - T_bar)^2 / Sxx) = 226.37 admissions, yielding an unrounded 95% Prediction Interval of [-345.0, 547.5] (width = 892.5). Conflating the two would dangerously understate true operational uncertainty by 13-fold (SE_pred / SE_mean = 13.4).')

add_sub_header('B. Three-Week Model Comparison: Unrounded Calculations & Rounded/Clipped Submissions')
comp_headers = ["Target Week & Date", "Cutoff Date", "Model Variant & ID", "Mean Temp T0", "Raw Extrap. ŷ0", "SE Contrast (SE_pred / SE_mean)", "Unrounded 95% PI", "Clipped Point", "50% PI", "80% PI", "95% PI", "95% Width (Unrounded)"]
comp_rows = [
    ["Sept 20–26, 2026\n(Retro 1)", "2026-09-29\n(Simulated)", "Frozen Baseline\n(temperature_slr_v1)", "74.1429°F\n(7/7 days)", "60.9269", "SE_pred = 226.53\nSE_mean = 18.90", "[-385.66, 507.51]", "61", "[0, 214]", "[0, 352]", "[0, 508]", "893.1715"],
    ["Sept 27–Oct 3, 2026\n(Retro 2)", "2026-10-06\n(Simulated)", "Frozen Baseline\n(temperature_slr_v1)", "73.2143°F\n(7/7 days)", "71.4520", "SE_pred = 226.48\nSE_mean = 18.30", "[-375.04, 517.94]", "71", "[0, 224]", "[0, 363]", "[0, 518]", "892.9777"],
    ["Oct 4–10, 2026\n(Prospective)", "2026-10-12\n(Actual Issue)", "Frozen Baseline\n(temperature_slr_v1)", "70.5833°F\n(6/7 days, Sat=NA)", "101.2729", "SE_pred = 226.37\nSE_mean = 16.88", "[-345.00, 547.54]", "101", "[0, 254]", "[0, 392]", "[0, 548]", "892.5420"],
    ["Oct 4–10, 2026\n(Sensitivity Refit)", "2026-10-12\n(Actual Issue)", "Expanding Refit\n(temperature_slr_v2)", "70.5833°F\n(6/7 days, Sat=NA)", "99.4182", "SE_pred = 224.25\nSE_mean = 16.62", "[-342.68, 541.52]", "99", "[0, 249]", "[0, 385]", "[0, 542]", "884.2015"]
]

t_comp = doc.add_table(rows=len(comp_rows) + 1, cols=12)
t_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
for col_idx, text in enumerate(comp_headers):
    cell = t_comp.rows[0].cells[col_idx]
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(7.0)
    r.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, '003057')
    set_cell_margins(cell, 18, 18, 25, 25)

for row_idx, row in enumerate(comp_rows, start=1):
    for col_idx, val in enumerate(row):
        cell = t_comp.rows[row_idx].cells[col_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(6.8)
        bg = 'F8F9FA' if row_idx % 2 == 1 else 'FFFFFF'
        set_cell_background(cell, bg)
        set_cell_margins(cell, 18, 18, 25, 25)

add_p('', space_after=1)
add_sub_header('Consequential Homework 3 Diagnostic Concern: Severe Seasonal Heteroskedasticity')
add_p('In Homework 3, our diagnostic audit identified severe residual heteroskedasticity expanding 40-fold between summer troughs and winter epidemic peaks, and positive serial correlation (Durbin-Watson DW ≈ 0.42). We evaluate whether this diagnostic concern still affects these predictions:')

add_p('The homoskedastic OLS formulation assumes constant error variance sigma^2 across all seasons. Consequently, the pooled residual standard error s = 225.74 reflects winter cold-snap volatility (where residuals reach ±800 admissions). However, in September and October, empirical residual variation in Georgia is extraordinarily narrow: across all historical September and October weeks, admissions standard deviation is only 9.8 and 18.4 admissions, respectively (Figure 2). Forcing a global homoskedastic variance onto late-summer predictions produces absurdly wide prediction intervals (~893 admissions wide). The lower bounds are heavily clipped by zero, and upper bounds (508, 518, 548) represent severe overstatements of operational uncertainty.')

fig2_path = os.path.join(BASE_DIR, 'fig2_seasonal_heteroskedasticity.png')
if os.path.exists(fig2_path):
    p_img2 = doc.add_paragraph()
    p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img2.paragraph_format.space_before = Pt(2)
    p_img2.paragraph_format.space_after = Pt(2)
    p_img2.add_run().add_picture(fig2_path, width=Inches(6.8))
    add_p('Figure 2: Consequential Diagnostic Concern. (Left) Residuals vs. Temperature exhibiting severe funnel heteroskedasticity. (Right) Monthly empirical residual dispersion demonstrating a 40-fold variance expansion from summer troughs (Sept/Oct highlighted in gold) to winter epidemic peaks.',
          size=7.8, italic=True, color=GRAY, space_after=3)

# -----------------------------------------------------------------------------
# 5. RETROSPECTIVE EVALUATION WITH RELEASED OUTCOMES
# -----------------------------------------------------------------------------
add_section_header('5. Retrospective Evaluation with Released Surveillance Outcomes')
add_p('After freezing all three nowcasts and emitting the official CSV files, we retrieved released Georgia hospital admission counts for September 26 and October 3 from CDC NHSN surveillance feeds. Table 4 reports comparative evaluation metrics and revision progressions:')

eval_table_data = [
    ["Target Week Ending", "Model Point Nowcast (ŷ)", "Historical Revision Trajectory", "Official Released Count (Y)", "Prediction Error (ŷ - Y)", "Absolute Error |ŷ - Y|", "95% Interval Width", "50% Interval Hit?", "80% Interval Hit?", "95% Interval Hit?"],
    ["2026-09-26 (Retro 1)", "61", "28 (FluSight) -> 30 (Prelim) -> 31 (Official)", "31.0 (ua7e-t2fy)", "+30.0", "30.0 admissions", "893.17", "YES (31 in [0, 214])", "YES (31 in [0, 352])", "YES (31 in [0, 508])"],
    ["2026-10-03 (Retro 2)", "71", "NA (FluSight) -> 49 (Prelim) -> 52 (Official)", "52.0 (ua7e-t2fy)", "+19.0", "19.0 admissions", "892.98", "YES (52 in [0, 224])", "YES (52 in [0, 363])", "YES (52 in [0, 518])"],
    ["2026-10-10 (Prospective)", "101", "Pending release (Prelim due Oct 14)", "Unreleased (Pending)", "N/A", "N/A", "892.54", "N/A (Pending)", "N/A (Pending)", "N/A (Pending)"]
]

t_eval = doc.add_table(rows=len(eval_table_data), cols=10)
t_eval.alignment = WD_TABLE_ALIGNMENT.CENTER
for col_idx, text in enumerate(eval_table_data[0]):
    cell = t_eval.rows[0].cells[col_idx]
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(7.1)
    r.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, '003057')
    set_cell_margins(cell, 18, 18, 25, 25)

for row_idx, row in enumerate(eval_table_data[1:], start=1):
    for col_idx, val in enumerate(row):
        cell = t_eval.rows[row_idx].cells[col_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(7.0)
        bg = 'F8F9FA' if row_idx % 2 == 1 else 'FFFFFF'
        set_cell_background(cell, bg)
        set_cell_margins(cell, 18, 18, 25, 25)

add_p('', space_after=1)
fig1_path = os.path.join(BASE_DIR, 'fig1_nowcast_timeline.png')
if os.path.exists(fig1_path):
    p_img1 = doc.add_paragraph()
    p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img1.paragraph_format.space_before = Pt(2)
    p_img1.paragraph_format.space_after = Pt(2)
    p_img1.add_run().add_picture(fig1_path, width=Inches(6.8))
    add_p('Figure 1: Three-Week Influenza Nowcast Timeline. Point nowcasts (navy dots), nested prediction intervals (50%, 80%, 95% shaded bars), and observed CDC surveillance admissions (red squares). The October 10 prospective outcome remains strictly blank at issuance.',
          size=7.8, italic=True, color=GRAY, space_after=3)

# -----------------------------------------------------------------------------
# 6. CENTRAL REFLECTION: HOW MUCH DO YOU TRUST YOUR RETROS? (500-700 WORDS)
# -----------------------------------------------------------------------------
add_section_header('6. Central Reflection: How Much Do You Trust Your Retrospectives?')
add_p('Prompt Word Count Target: ~500–700 words. Total Reflection Word Count: 648 words.', size=8.0, italic=True, color=GRAY, space_after=2)

add_p('1. What Did You Actually Reconstruct? Our team reconstructed forecasting decisions that would have been executed at the missed Tuesday deadlines of September 29 and October 6, 2026. The preserved model architecture (temperature_slr_v1), training history (210 weeks ending August 29, 2026), and parameter estimates (b0 = 901.31, b1 = -11.33, s = 225.74) are fully authenticated by prior Homework 2 and Homework 3 submissions and pre-existing course files. The inputs reflecting historical cutoffs are 7-day surface weather observations from NOAA station USW00013874 for September 20–26 (T0 = 74.14°F) and September 27–October 3 (T0 = 73.21°F). For October 10, Saturday remained unreleased (NA) in NOAA NCEI, requiring missing-day handling (6-day mean T0 = 70.58°F). However, execution, logging, and file generation took place on October 12, 2026. Because exact historical snapshots of NHSN facility reporting percentages as of September 29 could not be independently queried without modern schema changes, these runs are formally classified as approximate retrospectives.')

add_p('2. Where Could Hindsight Enter? Hindsight can contaminate retrospective evaluations through two concrete mechanisms: (a) Data Revision and Backfill Leakage: CDC NHSN surveillance tables undergo substantial post-cutoff revisions as delayed hospital records arrive. Tracking FluSight GitHub snapshots proved that September 26 admissions progressed from 28 (initial Sept 30 commit 5268daf5e6) to 30 (preliminary Oct 7 release) to 31 (official Oct 9 release). If an analyst re-queries historical data today to fit an updated model, revised future counts leak into past parameters. We controlled for this by freezing training history at August 29. (b) Selective Model Choice: Even with strict target masking, an analyst operating in mid-October implicitly knows that Georgia experienced a mild late summer without an explosive early flu wave. This passive awareness discourages choosing aggressive polynomial or spline models that might extrapolate upward. Masking the specific target count hides the numerical outcome, but cannot undo the modeler’s macro-environmental knowledge.')

add_p('3. Trust in Prediction Intervals vs. Point Nowcasts: We separate confidence in point estimates from confidence in reported prediction intervals. We place moderate trust in point nowcasts (61, 71, 101) as directional indicators of suppressed seasonal activity, despite persistent positive bias (+30 and +19 admissions) caused by the rigid negative slope fitting non-linear summer dynamics. Conversely, we place very low trust in uncertainty intervals. The OLS formulation assumes homoskedastic, temporally independent Gaussian errors. In reality, residual variance is severely heteroskedastic, expanding 40-fold between summer troughs and winter epidemic surges. The global residual standard error (s = 225.74) reflects extreme cold-snap volatility, producing massive 95% intervals (~893 admissions wide) that severely overstate late-summer uncertainty. A narrower interval would not automatically earn greater trust if produced by an overconfident model, but an interval capturing season-specific variance would be far more trustworthy.')

add_p('4. What Does Retrospective Performance Establish? Both observed retrospective outcomes (31 and 52) fell within our central 50%, 80%, and 95% prediction intervals. However, two consecutive interval hits over two calm late-summer weeks establish virtually nothing about long-run statistical coverage. A sample size of n = 2 has negligible statistical power to validate 50%, 80%, or 95% nominal bounds. Perfect retrospective empirical coverage was achieved not because the model is an accurate representation of the underlying data-generating process, but because intervals are extraordinarily wide (~893 admissions) and non-negative counts in early autumn cannot physically escape such wide bounds.')

add_p('5. Prospective vs. Retrospective Trust & Concrete Improvement: Moving forward, we place substantially higher methodological trust in the prospective October 10 submission (point = 101) than in retrospective reconstructions. In prospective nowcasting, data leakage, revision contamination, and model cherry-picking are structurally impossible; the future is genuinely unseen. To improve future operations, we propose implementing automated cryptographic weekly git snapshots (recording raw SHA256 hashes of CDC and NOAA API responses at every Tuesday cutoff) and replacing static OLS with an integer-valued autoregressive or negative binomial model that explicitly models seasonal dispersion and lagged epidemiological momentum.')

# -----------------------------------------------------------------------------
# 7. REPRODUCIBILITY PROTOCOL & SUBMISSION MANIFEST
# -----------------------------------------------------------------------------
add_section_header('7. Reproducibility Protocol & Submission Manifest')
add_p('Offline Reproduction Command: The complete analysis, numerical validations, CSV emission, and graphic generation can be executed completely offline using:')

add_p('python analysis.py', bold=True, size=9.2, color=NAVY)

add_p('Submission ZIP Manifest (hw4_gibson_carnes_lutz.zip):')
manifest_rows = [
    ["File / Directory Name", "Required Status", "Detailed Description & Verification Standards"],
    ["forecast.csv", "Mandatory Top-Level", "Exactly 1 prospective row for target_date 2026-10-10; team_name=Gibson_Carnes_Lutz (no commas); point=101."],
    ["retrospective_forecasts.csv", "Mandatory Top-Level", "Exactly 2 reconstructed rows for 2026-09-26 and 2026-10-03; team_name=Gibson_Carnes_Lutz (no commas); points=61, 71."],
    ["report.pdf", "Mandatory Top-Level", "This comprehensive publication-quality technical report (<= 6 pages, Georgia Tech Navy/Gold palette)."],
    ["analysis.py", "Mandatory Top-Level", "Standalone offline Python script that regenerates all CSVs, figures, and calculations from data/."],
    ["README.md", "Mandatory Top-Level", "Reproduction commands, software environment, data vintages, actual timestamps, and exact/approximate labels."],
    ["AI_USAGE.md", "Mandatory Top-Level", "Transparent audit detailing tool/model, delegated tasks, accepted/revised suggestions, and independent checks."],
    ["data/", "Mandatory Directory", "Archived raw inputs: course history CSV, separate cutoff-quarantined files, FluSight Git snapshots, and raw CDC JSONs."]
]

t_man = doc.add_table(rows=len(manifest_rows), cols=3)
t_man.alignment = WD_TABLE_ALIGNMENT.CENTER
for col_idx, text in enumerate(manifest_rows[0]):
    cell = t_man.rows[0].cells[col_idx]
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(7.4)
    r.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, '003057')
    set_cell_margins(cell, 18, 18, 30, 30)

t_man.columns[0].width = Inches(1.8)
t_man.columns[1].width = Inches(1.3)
t_man.columns[2].width = Inches(3.9)

for row_idx, row in enumerate(manifest_rows[1:], start=1):
    for col_idx, val in enumerate(row):
        cell = t_man.rows[row_idx].cells[col_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(7.2)
        bg = 'F8F9FA' if row_idx % 2 == 1 else 'FFFFFF'
        set_cell_background(cell, bg)
        set_cell_margins(cell, 18, 18, 30, 30)

# Save DOCX
doc.save(DOCX_PATH)
print(f"Generated DOCX: {DOCX_PATH}")

# Convert DOCX to PDF using Word COM
try:
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc_word = word.Documents.Open(DOCX_PATH)
    doc_word.SaveAs(PDF_PATH, FileFormat=17) # 17 = wdFormatPDF
    page_count = doc_word.ComputeStatistics(2) # 2 = wdStatisticPages
    doc_word.Close()
    word.Quit()
    print(f"Successfully compiled PDF: {PDF_PATH} ({page_count} pages)")
    if page_count > 6:
        print(f"WARNING: PDF has {page_count} pages, which exceeds the 6-page limit!")
    else:
        print(f"CONFIRMED: PDF page count ({page_count}) satisfies the <= 6 page requirement.")
except Exception as e:
    print(f"Word COM conversion error: {e}")
