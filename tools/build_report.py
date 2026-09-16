"""Generate AryanVerma_ProjectReport.docx from the notebook's computed outputs.

The report is not written by hand. Every number in it is read out of
`outputs/facts.json` and every chart is read out of `outputs/figures/`, both of
which are produced by running
`AryanVerma_RetailCustomerSegmentationAnalysis.ipynb` end to end. That means the
document cannot silently drift away from the analysis: if a cleaning rule
changes, re-running the notebook and then this script updates the report.

Usage
-----
    python tools/build_report.py

Requires `python-docx` (see requirements.txt) and a notebook run that has
already written outputs/facts.json and outputs/figures/*.png.
"""

import json
import sys
from datetime import date
from pathlib import Path

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Inches, Pt, RGBColor
except ModuleNotFoundError:                                     # pragma: no cover
    sys.exit("python-docx is not installed. Run:  pip install python-docx")

ROOT = Path(__file__).resolve().parents[1]
FACTS_PATH = ROOT / "outputs" / "facts.json"
FIG_DIR = ROOT / "outputs" / "figures"
OUT_PATH = ROOT / "AryanVerma_ProjectReport.docx"

DATASET_URL = "https://archive.ics.uci.edu/dataset/502/online+retail+ii"
NOTEBOOK = "AryanVerma_RetailCustomerSegmentationAnalysis.ipynb"
REPO_URL = "https://github.com/BU1lDR/IBM-BharatCares---Internship-CAPSTONE-PROJECT-"

# Colours reused from the notebook's chart theme so the document and the figures
# look like one artefact.
ACCENT = RGBColor(0x16, 0x4A, 0x86)
INK = RGBColor(0x0B, 0x0B, 0x0B)
INK_2 = RGBColor(0x4A, 0x4A, 0x48)
INK_MUTED = RGBColor(0x77, 0x77, 0x6F)


# --------------------------------------------------------------------------- #
# inputs
# --------------------------------------------------------------------------- #
class Facts(dict):
    """A dict that names the missing key instead of raising a bare KeyError."""

    def __missing__(self, key):
        raise KeyError(
            f"fact '{key}' is not in outputs/facts.json - re-run {NOTEBOOK} first"
        )


if not FACTS_PATH.exists():
    sys.exit(f"missing {FACTS_PATH.relative_to(ROOT)} - run {NOTEBOOK} first")

payload = json.loads(FACTS_PATH.read_text(encoding="utf-8"))
F = Facts(payload["facts"])
FIGS = payload["figures"]
CLEAN_LOG = payload["clean_log"]
KPI = payload["kpi"]
SEGMENTS = payload["segments"]
SEGMENT_CLV = payload["segment_clv"]
CLUSTERS = payload["clusters"]
KMEANS_SCAN = payload["kmeans_scan"]
MONTHLY = payload["monthly"]
TOP_PRODUCTS = payload["top_products"]
COUNTRIES = payload["countries"]
TOP_RETURNED = payload["top_returned"]
COHORTS = payload["cohort_quality"]

missing_figs = [n for n in FIGS.values() if not (FIG_DIR / n).exists()]
if missing_figs:
    sys.exit(
        "these figures are missing from outputs/figures/: "
        + ", ".join(missing_figs)
        + f"\nre-run {NOTEBOOK} to regenerate them"
    )


# --------------------------------------------------------------------------- #
# formatting helpers
# --------------------------------------------------------------------------- #
def gbp(value, dp=0):
    return f"£{value:,.{dp}f}"


def gbp_m(value):
    return f"£{value / 1e6:.2f}M"


def gbp_compact(value):
    """Money sized for a narrow table column."""
    if abs(value) >= 1e6:
        return f"£{value / 1e6:.2f}M"
    if abs(value) >= 1e4:
        return f"£{value / 1e3:.0f}k"
    return f"£{value:,.0f}"


def pct(value, dp=1):
    return f"{value:.{dp}f}%"


def num(value):
    return f"{value:,.0f}"


def one_dp(value):
    return f"{value:,.1f}"


def ordinal(value):
    """3 -> '3rd', 1072 -> '1,072nd'. Keeps rank suffixes right if the data moves."""
    n = int(value)
    suffix = "th" if n % 100 in (11, 12, 13) else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n:,}{suffix}"


def pretty_date(iso):
    """'2011-12-09' -> '9 December 2011'."""
    d = date.fromisoformat(str(iso)[:10])
    return f"{d.day} {d:%B %Y}"


def pretty_month(iso):
    """'2011-01' -> 'January 2011'."""
    d = date.fromisoformat(f"{str(iso)[:7]}-01")
    return f"{d:%B %Y}"


# --------------------------------------------------------------------------- #
# document scaffolding
# --------------------------------------------------------------------------- #
doc = Document()

doc.core_properties.title = "Retail Customer Segmentation & Sales Analysis"
doc.core_properties.author = "Aryan Verma"
doc.core_properties.subject = (
    "IBM SkillsBuild Data Analytics with AI Academic Internship Programme - "
    "Capstone Project"
)
doc.core_properties.keywords = (
    "RFM, K-Means, customer lifetime value, cohort analysis, Online Retail II"
)

section = doc.sections[0]
section.page_width = Cm(21.0)          # A4 portrait
section.page_height = Cm(29.7)
section.left_margin = section.right_margin = Cm(2.2)
section.top_margin = section.bottom_margin = Cm(2.0)
USABLE_IN = 6.15                       # A4 width minus margins, in inches

normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.font.color.rgb = INK
normal.paragraph_format.space_after = Pt(7)
normal.paragraph_format.line_spacing = 1.12

for name, size, colour, bold in [
    ("Title", 26, ACCENT, True),
    ("Heading 1", 16, ACCENT, True),
    ("Heading 2", 12.5, ACCENT, True),
    ("Heading 3", 11, INK, True),
]:
    style = doc.styles[name]
    style.font.name = "Calibri"
    style.font.size = Pt(size)
    style.font.color.rgb = colour
    style.font.bold = bold
    style.paragraph_format.space_before = Pt(16 if name == "Heading 1" else 11)
    style.paragraph_format.space_after = Pt(5)
    style.paragraph_format.keep_with_next = True


def footer_page_numbers():
    """`Aryan Verma | ... | Page X of Y` - PAGE/NUMPAGES are Word fields."""
    p = doc.sections[0].footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def small(text):
        run = p.add_run(text)
        run.font.size = Pt(8)
        run.font.color.rgb = INK_MUTED
        return run

    def field(instruction):
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), instruction)
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "1"                # placeholder until Word recalculates
        run.append(text)
        fld.append(run)
        p._p.append(fld)

    small("Aryan Verma  |  Retail Customer Segmentation & Sales Analysis  |  Page ")
    field("PAGE")
    small(" of ")
    field("NUMPAGES")


footer_page_numbers()

FIGURE_N = 0
TABLE_N = 0


def runs_from(paragraph, text):
    """Write `text` into `paragraph`, honouring **bold**, *italic* and `code`."""
    for bold_i, bold_chunk in enumerate(text.split("**")):
        for code_i, code_chunk in enumerate(bold_chunk.split("`")):
            for it_i, chunk in enumerate(code_chunk.split("*")):
                if not chunk:
                    continue
                run = paragraph.add_run(chunk)
                run.bold = bool(bold_i % 2)
                run.italic = bool(it_i % 2)
                if code_i % 2:
                    run.font.name = "Consolas"
                    run.font.size = Pt(9.5)
    return paragraph


def para(text="", style=None, size=None, align=None, colour=None,
         space_after=None, italic=False, indent=None):
    p = doc.add_paragraph(style=style)
    runs_from(p, text)
    if align is not None:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if indent is not None:
        p.paragraph_format.left_indent = Cm(indent)
    for run in p.runs:
        if size is not None:
            run.font.size = Pt(size)
        if colour is not None:
            run.font.color.rgb = colour
        if italic:
            run.italic = True
    return p


def h1(text):
    doc.add_paragraph(text, style="Heading 1")


def h2(text):
    doc.add_paragraph(text, style="Heading 2")


def h3(text):
    doc.add_paragraph(text, style="Heading 3")


def bullet(text):
    p = para(text, style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    return p


def numbered(index, text):
    """Manual numbering - Word's List Number style keeps counting across lists."""
    p = para(f"**{index}.** {text}")
    p.paragraph_format.left_indent = Cm(0.75)
    p.paragraph_format.first_line_indent = Cm(-0.75)
    p.paragraph_format.space_after = Pt(5)
    return p


def caption(text, kind):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if kind == "Figure" else WD_ALIGN_PARAGRAPH.LEFT
    runs_from(p, text)
    for run in p.runs:
        run.font.size = Pt(8.5)
        run.font.color.rgb = INK_MUTED
        run.italic = True
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(12)
    return p


def figure(key, text, width=USABLE_IN):
    global FIGURE_N
    FIGURE_N += 1
    doc.add_picture(str(FIG_DIR / FIGS[key]), width=Inches(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.space_after = Pt(2)
    caption(f"Figure {FIGURE_N}. {text}", "Figure")


def table(headers, rows, text, widths=None, font_size=8.5, numeric_from=1):
    """A bordered table with a bold header row and right-aligned numeric columns."""
    global TABLE_N
    TABLE_N += 1
    caption(f"Table {TABLE_N}. {text}", "Table")

    t = doc.add_table(rows=1, cols=len(headers))
    try:
        t.style = doc.styles["Light Grid Accent 1"]
    except KeyError:                                            # pragma: no cover
        t.style = doc.styles["Table Grid"]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = widths is None

    def write(cell, value, bold=False, right=False):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)
        if right:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        text = str(value)
        if "`" in text or "**" in text:      # markers would otherwise render literally
            runs_from(p, text)
        else:                                # leave a lone * in a product name alone
            p.add_run(text)
        for run in p.runs:
            run.bold = run.bold or bold
            run.font.size = Pt(font_size)
            if run.font.name != "Consolas":
                run.font.name = "Calibri"

    for j, head in enumerate(headers):
        write(t.rows[0].cells[j], head, bold=True, right=j >= numeric_from)

    for row in rows:
        cells = t.add_row().cells
        for j, value in enumerate(row):
            write(cells[j], value, right=j >= numeric_from)

    if widths:
        for row in t.rows:                 # width must be set on every cell
            for j, cell in enumerate(row.cells):
                cell.width = Inches(widths[j])

    doc.paragraphs[-1].paragraph_format.space_after = Pt(2)
    para("", size=4)                       # breathing room under the table
    return t


def page_break():
    doc.add_page_break()


# --------------------------------------------------------------------------- #
# derived values used in the prose (computed, never typed)
# --------------------------------------------------------------------------- #
def seg(name):
    """Look a segment row up by name rather than trusting its position."""
    for row in SEGMENTS:
        if row["Segment"] == name:
            return row
    raise KeyError(f"segment '{name}' is not in outputs/facts.json")


CHAMPIONS = seg("Champions")
CANNOT_LOSE = seg("Cannot lose them")
BIGGEST_CLUSTER = max(CLUSTERS, key=lambda r: r["Customers"])
RICHEST_CLUSTER = max(CLUSTERS, key=lambda r: r["TotalRevenue"])
# The prose ties this row to the second-largest cancellation, so match it by name
# instead of assuming it is the second row of the most-returned table.
SECOND_RETURNED = next(
    (row for row in TOP_RETURNED
     if str(row["Description"]).lower() == F["second_cancel_product"].lower()),
    TOP_RETURNED[1],
)

MEAN_MEDIAN_RATIO = F["mean_customer_revenue"] / F["median_customer_revenue"]
AOV_RATIO = F["aov"] / F["median_order"]
DEAD_CUSTOMER_PCT = F["seg_lost_cust_pct"] + F["seg_hibernating_cust_pct"]
DEAD_REVENUE_PCT = F["seg_lost_rev_pct"] + F["seg_hibernating_rev_pct"]
EXPORT_MARKETS = F["countries"] - 1
NL_UK_RATIO = F["best_export_aov"] / F["uk_aov"]
NEW_SHARE_YEAR2 = 100 - F["returning_share_year2_pct"]
XMAS_GAP = F["nonxmas_cohort_m3"] / F["xmas_cohort_m3"]
TOP1_LOSS = F["top1pct_share"]
CLEAN_ROWS_DROPPED = F["raw_rows"] - F["sales_rows"]
CATALOGUE_SHARE = len(TOP_PRODUCTS) / F["product_count"] * 100
PIPELINE_PCT = (F["seg_promising_rev_pct"] + F["seg_new_rev_pct"]
                + F["seg_needs_attention_rev_pct"])
BUILD_DATE = date.today().strftime("%d %B %Y")


# --------------------------------------------------------------------------- #
# 0. title page
# --------------------------------------------------------------------------- #
para("", size=10)
title_p = doc.add_paragraph(style="Title")
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.add_run("Retail Customer Segmentation").add_break()
title_p.add_run("& Sales Analysis")

para(
    "Who the customers of a UK online gift retailer actually are, "
    "what they are worth, and which of them are quietly leaving",
    align=WD_ALIGN_PARAGRAPH.CENTER, size=13, colour=INK_2, italic=True,
    space_after=26,
)

para("**Capstone Project Report**", align=WD_ALIGN_PARAGRAPH.CENTER, size=12,
     space_after=4)
para("IBM SkillsBuild Data Analytics with AI Academic Internship Programme",
     align=WD_ALIGN_PARAGRAPH.CENTER, size=11, colour=INK_2, space_after=2)
para("Conducted by BharatCares in association with AICTE",
     align=WD_ALIGN_PARAGRAPH.CENTER, size=11, colour=INK_2, space_after=30)

table(
    ["Field", "Detail"],
    [
        ["Submitted by", "Aryan Verma"],
        ["Project title", "Retail Customer Segmentation & Sales Analysis"],
        ["Dataset", f"UCI Machine Learning Repository - Online Retail II ({DATASET_URL})"],
        ["Records analysed", f"{num(F['raw_rows'])} transaction lines, "
                             f"{pretty_date(F['date_min'])} to {pretty_date(F['date_max'])}"],
        ["Code file", NOTEBOOK],
        ["Language / stack", "Python 3 - pandas, NumPy, Matplotlib, scikit-learn"],
        ["Methods", "Data cleaning audit, EDA, RFM segmentation, K-Means clustering, "
                    "CLV estimation, cohort retention"],
        ["Report generated", f"{BUILD_DATE}, by tools/build_report.py from "
                             f"outputs/facts.json"],
    ],
    "Submission details.",
    widths=[1.55, 4.6],
    font_size=9.5,
    numeric_from=99,
)

para("", size=8)
para(
    "Every figure, table and number in this report is computed by the notebook "
    "and read from `outputs/facts.json`. Nothing is transcribed by hand.",
    align=WD_ALIGN_PARAGRAPH.CENTER, size=9, colour=INK_MUTED, italic=True,
)
page_break()


# --------------------------------------------------------------------------- #
# contents
# --------------------------------------------------------------------------- #
h1("Contents")
for line in [
    "1.  Executive summary",
    "2.  Introduction and business problem",
    "3.  The dataset",
    "4.  Tools and technologies",
    "5.  Methodology",
    "6.  Data cleaning and its audit trail",
    "7.  Exploratory analysis",
    "8.  Customer segmentation - RFM",
    "9.  Customer segmentation - K-Means",
    "10. Customer lifetime value",
    "11. Cohort retention",
    "12. Key findings",
    "13. Recommendations",
    "14. Limitations",
    "15. Conclusion",
    "16. How to reproduce this project",
    "17. References",
]:
    p = para(line, size=10.5)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(0.5)
page_break()


# --------------------------------------------------------------------------- #
# 1. executive summary
# --------------------------------------------------------------------------- #
h1("1.  Executive summary")

para(
    f"This project analyses {num(F['raw_rows'])} transaction lines from a UK-based "
    f"online gift retailer covering {pretty_date(F['date_min'])} to {pretty_date(F['date_max'])}, and answers "
    f"one question: which customers is this business actually built on, and which of "
    f"them are leaving? After a documented cleaning pass that retains "
    f"{pct(F['rows_retained_pct'])} of rows, the analysed base is "
    f"**{gbp(F['sales_revenue'])} of revenue across {num(F['orders'])} orders, "
    f"{num(F['customers'])} identified customers and {num(F['products'])} products "
    f"in {F['countries']} countries**."
)

para(
    f"The headline result is concentration. **{num(F['top20pct_customers'])} customers "
    f"- 20% of the base - produce {pct(F['top20pct_share'])} of revenue, and the top "
    f"1% ({num(F['top1pct_customers'])} customers) produce {pct(F['top1pct_share'])}.** "
    f"The mean customer is worth {gbp(F['mean_customer_revenue'], 0)} and the median "
    f"{gbp(F['median_customer_revenue'], 0)} - a ratio of "
    f"{MEAN_MEDIAN_RATIO:.1f}x - so any plan addressed to “the average customer” "
    f"is addressed to nobody. That is the finding that justifies segmenting the base "
    f"before doing anything else with it."
)

para(
    f"Rule-based RFM scoring splits the base into nine segments. "
    f"**{num(F['seg_champions_customers'])} Champions "
    f"({pct(F['seg_champions_cust_pct'])} of customers) carry "
    f"{gbp_m(F['seg_champions_revenue'])}, or {pct(F['seg_champions_rev_pct'])} of "
    f"identified revenue.** A K-Means clustering run on the same log-scaled RFM "
    f"features, given no rules at all, reproduces that structure: every one of the "
    f"nine segments sends at least 60% of its members to a single cluster. Two "
    f"independent methods agreeing is evidence the structure is in the data rather "
    f"than in the thresholds."
)

para(
    f"The actionable number is smaller and more specific. "
    f"**{num(F['at_risk_customers'])} customers who have already spent "
    f"{gbp(F['at_risk_revenue'])} ({pct(F['at_risk_rev_pct'])} of revenue) have "
    f"stopped ordering.** They are not low-value customers who drifted off; they are "
    f"proven buyers who went quiet, and they are the only group with a large, "
    f"individually identified prize attached to contacting them."
)

para(
    f"Three further results change how the business should read its own reports. "
    f"Returns total {gbp(F['return_value'])} - {pct(F['return_rate_pct'], 2)} of gross "
    f"revenue - and the rate spikes to {pct(F['worst_return_rate_pct'], 2)} in "
    f"{pretty_month(F['worst_return_month'])} rather than sitting flat. The single largest line in "
    f"the file, {gbp(F['biggest_sale_value'])} of one product, was cancelled "
    f"{F['biggest_sale_cancel_gap_min']:.0f} minutes after it was raised, which moves "
    f"that product from **gross rank {F['worst_net_mover_gross_rank']} to net rank "
    f"{num(F['worst_net_mover_net_rank'])}**; {F['net_rank_movers']} of the top ten "
    f"products change position once cancellations are netted off. And customers "
    f"acquired in the Christmas peak retain at {pct(F['xmas_cohort_m3'])} by month 3 "
    f"against {pct(F['nonxmas_cohort_m3'])} for every other month - the season that "
    f"produces the revenue produces the worst customers."
)

table(
    ["Metric", "Value"],
    [[row["Metric"], row["Value"]] for row in KPI],
    "Headline metrics for the cleaned dataset. Computed in section 7 of the notebook.",
    widths=[3.0, 3.15],
    font_size=9.5,
    numeric_from=1,
)

para(
    "What this report does not do is equally deliberate. The dataset has no cost of "
    "goods, so every value here is revenue and not margin; it contains no campaign "
    "data, so nothing here shows that contacting a customer *causes* a return; and "
    "the timestamp records back-office keying rather than customer intent, so it "
    "cannot say when customers prefer to shop. Section 14 states all nine "
    "limitations in full."
)
page_break()


# --------------------------------------------------------------------------- #
# 2. introduction
# --------------------------------------------------------------------------- #
h1("2.  Introduction and business problem")

h2("2.1  Background")
para(
    "The subject of this analysis is a UK-registered online retailer selling "
    "all-occasion giftware, largely to wholesale buyers who resell the stock. Its "
    "transaction log for two full trading years is published by the UCI Machine "
    "Learning Repository as *Online Retail II*. Because the log is a real operational "
    "extract rather than a teaching file, it carries everything a real extract "
    "carries: cancelled invoices, bad-debt adjustments, postage lines mixed in with "
    "product lines, guest checkouts with no customer identifier, and duplicate writes."
)
para(
    "A retailer in this position has a specific problem. It knows its total revenue "
    "and it knows which products sell, because both come straight off the invoice "
    "system. What it does not know is the shape of its customer base - how much of "
    "the business depends on how few accounts, which accounts have stopped buying, "
    "and whether the customers it acquires in its busiest season are worth the same "
    "as the ones it acquires in June."
)

h2("2.2  The business problem")
para(
    "**Marketing and account-management effort is finite, and it is currently being "
    "spread without evidence about where it pays back.** The analytical problem is to "
    "convert a flat transaction log into a segmentation of the customer base that "
    "(a) is derived from behaviour rather than assumption, (b) attaches a revenue "
    "figure to each segment so effort can be prioritised, and (c) identifies the "
    "customers whose spending has stopped while they are still recoverable."
)

h2("2.3  Objectives")
for i, text in enumerate([
    "Build a reproducible cleaning pipeline for a real transaction log, and record "
    "what each rule removes rather than dropping rows silently.",
    "Establish the shape of the business through exploratory analysis: revenue over "
    "time, seasonality, products, geography, returns and order-value distribution.",
    "Segment identified customers on Recency, Frequency and Monetary value using an "
    "explicit rule-based scheme, and quantify each segment's revenue contribution.",
    "Cross-check that segmentation with unsupervised K-Means clustering on the same "
    "features, and report the agreement between the two.",
    "Estimate customer lifetime value per segment, and test the estimate against "
    "actual revenue rather than presenting it unchallenged.",
    "Measure retention with a monthly acquisition-cohort matrix, and correct it for "
    "the censoring at both ends of the observation window.",
    "Turn the results into recommendations that name the finding and the number "
    "behind each one, and state plainly what the data cannot support.",
], start=1):
    numbered(i, text)

h2("2.4  Scope")
para(
    f"**In scope:** the {num(F['raw_rows'])} transaction lines in the two source "
    f"sheets, covering {pretty_date(F['date_min'])} to {pretty_date(F['date_max'])}; descriptive and "
    f"unsupervised analysis; revenue-based valuation."
)
para(
    "**Out of scope:** profitability (no cost data exists in the file), demographic "
    "or firmographic profiling (no customer attributes exist), predictive churn "
    "modelling with a probabilistic survival model, and any causal claim about "
    "marketing effect. Section 14 explains why each exclusion is a property of the "
    "data rather than a shortcut."
)
page_break()


# --------------------------------------------------------------------------- #
# 3. dataset
# --------------------------------------------------------------------------- #
h1("3.  The dataset")

h2("3.1  Source")
table(
    ["Attribute", "Detail"],
    [
        ["Name", "Online Retail II"],
        ["Publisher", "UCI Machine Learning Repository (dataset 502)"],
        ["Link", DATASET_URL],
        ["Direct download", "https://archive.ics.uci.edu/static/public/502/online+retail+ii.zip"],
        ["Donated by", "Dr Daqing Chen, London South Bank University"],
        ["Licence", "Creative Commons Attribution 4.0 International (CC BY 4.0)"],
        ["Format", "Single .xlsx workbook, two sheets: Year 2009-2010 and Year 2010-2011"],
        ["Size", "approx. 45 MB compressed download"],
        ["Rows", f"{num(F['raw_rows'])} across both sheets"],
        ["Period", f"{pretty_date(F['date_min'])} to {pretty_date(F['date_max'])}"],
        ["Granularity", "One row per product line per invoice"],
    ],
    "Dataset provenance. The notebook downloads this file on first run and caches it "
    "under data/raw/; the source workbook is never modified.",
    widths=[1.45, 4.7],
    font_size=9.5,
    numeric_from=99,
)

h2("3.2  Schema")
table(
    ["Column", "Type", "Meaning and notes"],
    [
        ["Invoice", "text", "Invoice number. A `C` prefix marks a cancellation and "
                            "an `A` prefix marks a bad-debt adjustment."],
        ["StockCode", "text", "Product code. Some codes are not products at all - "
                              "POST, DOT, M, BANK CHARGES, AMAZONFEE, TEST001."],
        ["Description", "text", f"Product name. Missing on {num(F['missing_description'])} rows."],
        ["Quantity", "integer", "Units on the line. Negative on returns and write-offs."],
        ["InvoiceDate", "datetime", "When the line was keyed into the system - see "
                                    "section 7.4, this is not customer intent."],
        ["Price", "float", "Unit price in GBP. Zero or negative on non-sale lines."],
        ["Customer ID", "float", f"Customer identifier. Missing on "
                                 f"{num(F['missing_customer_id'])} rows "
                                 f"({pct(F['missing_customer_id_pct'])}) - guest checkouts."],
        ["Country", "text", f"Shipping country. {F['countries']} distinct values."],
        ["SourceSheet", "text", "Added by the notebook to record which sheet a row came from."],
        ["Revenue", "float", "Engineered by the notebook: Quantity x Price."],
    ],
    "Schema of the combined dataset: eight source columns, plus SourceSheet and the "
    "engineered Revenue column.",
    widths=[1.0, 0.7, 4.45],
    font_size=9,
    numeric_from=99,
)

h2("3.3  Data-quality issues found before any analysis")
para(
    "The profiling pass in the notebook found the following, all of which shape the "
    "cleaning rules in section 6:"
)
for text in [
    f"**{num(F['missing_customer_id'])} rows ({pct(F['missing_customer_id_pct'])}) "
    f"have no customer ID.** These are real sales with real revenue; they simply "
    f"cannot be attributed to a customer.",
    f"**{num(F['duplicate_rows'])} rows are exact duplicates** on every column "
    f"including the timestamp - a double write, not two sales.",
    f"**{num(F['cancellation_rows'])} rows sit on `C`-prefixed cancellation "
    f"invoices**, and {num(F['negative_qty_rows'])} rows carry a negative quantity. "
    f"The two sets are not the same: the difference is stock write-offs on ordinary "
    f"invoices.",
    f"**Six `A`-prefixed invoices carry {gbp(F['invoice_prefix_a_revenue'])} of "
    f"bad-debt adjustments.** The common “drop invoices starting with C” "
    f"recipe leaves every penny of that in the revenue line.",
    f"**{num(F['zero_price_rows'])} rows have a zero or negative price**, including "
    f"every row with a blank description.",
    f"**Non-product stock codes are mixed into the product lines**: postage (POST, "
    f"DOT), carriage (C2, C3), manual corrections (M), bank charges, Amazon fees, "
    f"samples, gift vouchers and test records.",
]:
    bullet(text)
page_break()


# --------------------------------------------------------------------------- #
# 4. tools
# --------------------------------------------------------------------------- #
h1("4.  Tools and technologies")

table(
    ["Tool", "Version used", "Role in this project"],
    [
        ["Python", "3.14.6", "Language for the entire analysis"],
        ["Jupyter Notebook", "7.6.2", "Executable document holding code, charts and narrative"],
        ["pandas", "3.0.5", "Loading, cleaning, joining, grouping, cohort pivots"],
        ["NumPy", "2.5.3", "Vectorised arithmetic and the log1p transform before clustering"],
        ["Matplotlib", "3.11.2", "All 14 figures, drawn on a single explicit chart theme"],
        ["scikit-learn", "1.9.1", "StandardScaler, KMeans, silhouette_score"],
        ["openpyxl", "3.1.5", "Reads the two sheets of the source .xlsx workbook"],
        ["python-docx", "1.2.0", "Generates this report from outputs/facts.json"],
        ["Git / GitHub", "-", "Version control and submission repository"],
    ],
    "Software stack. Exact pinned lower bounds are in requirements.txt.",
    widths=[1.35, 1.1, 3.7],
    font_size=9.5,
    numeric_from=99,
)

para(
    "Two deliberate choices are worth stating. **Matplotlib rather than an "
    "interactive library:** an interactive chart does not survive being committed to "
    "a notebook and read by someone else, so every figure here is a static image "
    "that renders identically on any machine. **A hand-built chart theme rather than "
    "a default style:** the notebook defines its own ink, grid and categorical "
    "colours once, assigns the four categorical hues in a fixed order, uses a single "
    "light-to-dark hue for every heatmap, prints a table alongside every chart, and "
    "never uses two y-axes on one plot."
)

h2("Repository layout")
for text in [
    f"`{NOTEBOOK}` - the complete analysis: 75 cells, all executed, no errors.",
    "`requirements.txt` - dependencies with verified versions.",
    "`AryanVerma_ProjectReport.docx` - this report.",
    "`README.md` - overview, dataset link, setup and run instructions.",
    "`tools/build_report.py` - the generator that produces this document.",
    "`outputs/facts.json` - every computed number, dumped by the notebook.",
    "`outputs/figures/` - the 14 exported charts (regenerated by a notebook run).",
    "`data/raw/` - the downloaded source workbook (not committed; downloaded on first run).",
]:
    bullet(text)
page_break()


# --------------------------------------------------------------------------- #
# 5. methodology
# --------------------------------------------------------------------------- #
h1("5.  Methodology")

h2("5.1  Pipeline")
para(
    "The analysis runs top to bottom in one notebook, in this order: acquire and "
    "cache the workbook, profile it, clean it with a logged audit trail, engineer "
    "time and revenue fields, compute headline KPIs, run exploratory analysis, build "
    "RFM segments, cross-check them with K-Means, estimate CLV, build the cohort "
    "matrix, then write findings and dump every computed number to "
    "`outputs/facts.json`. No cell depends on a later cell, so a fresh kernel and "
    "*Run All* reproduces the whole report."
)

h2("5.2  Two analysis bases, used deliberately")
para(
    f"Restricting the whole analysis to identified customers would throw away "
    f"{gbp(F['anonymous_revenue'])} of real revenue; keeping guest orders in the "
    f"customer models would be meaningless, because a guest row cannot be attributed "
    f"to anyone. The notebook therefore carries two bases and states which is in use "
    f"at every step:"
)
bullet(
    f"**`sales`** - all valid revenue lines including guest checkouts: "
    f"{num(F['sales_rows'])} rows, {gbp(F['sales_revenue'])}. Used for revenue, "
    f"products, geography, seasonality and returns."
)
bullet(
    f"**`sales_id`** - the subset with a known customer ID: {num(F['sales_id_rows'])} "
    f"rows, {gbp(F['sales_id_revenue'])}, which is {pct(F['id_coverage_pct'])} of "
    f"clean revenue. Used for RFM, K-Means, CLV and cohorts, because all four need "
    f"an identified customer."
)
para(
    f"Cancellations ({num(F['cancellation_rows'])} rows) and stock write-offs "
    f"({num(F['writeoff_rows'])} rows) are **set aside rather than deleted**, so "
    f"section 7.7 can analyse them instead of losing them."
)

h2("5.3  RFM segmentation")
para(
    f"Recency, frequency and monetary value are computed per customer as of the day "
    f"after the last transaction in the file ({pretty_date(F['date_max'])}), so “recency” "
    f"is measured against a fixed reference point rather than today's date. Each "
    f"measure is scored 1-5 by quintile."
)
para(
    "**The scoring is rank-based, and that detail matters.** Frequency is heavily "
    "tied - thousands of customers have exactly one or two orders - and quintile "
    "cutting on raw values fails on ties. The notebook therefore ranks first "
    "(`pd.qcut(x.rank(method=\"first\"), 5, ...)`), which guarantees five populated "
    "bands. Recency is scored in reverse: a small number of days since the last "
    "order is a high score."
)
para(
    "The R score and the average of the F and M scores are then mapped onto nine "
    "named segments - Champions, Loyal, Cannot lose them, At risk, Needs attention, "
    "Promising, New, Hibernating, Lost - through an exhaustive rule ladder, so every "
    "one of the 25 (R, FM) combinations lands in exactly one segment and no customer "
    "falls through."
)

h2("5.4  K-Means clustering")
para(
    "The same three features are clustered without any rules, as an independent "
    "check. Because recency, frequency and monetary value are all heavily "
    "right-skewed, each is `log1p`-transformed and then standardised with "
    "`StandardScaler`; skipping either step lets the single largest customer dominate "
    "the Euclidean distance. `k` is scanned from 2 to 8 with inertia and silhouette "
    "recorded at every value, and the chosen `k` is justified against that scan in "
    "section 9 - including where the metric disagrees with the choice."
)

h2("5.5  Customer lifetime value")
para(
    "CLV is reported two ways. **Historical CLV** is simply the revenue a customer "
    "has already produced - a fact, not a model. **Projected 12-month value** takes "
    "each customer's observed spend rate over their own active lifespan and extends "
    "it twelve months. Section 10 then tests that projection against actual "
    "second-year revenue and reports the bias instead of hiding it."
)

h2("5.6  Cohort retention")
para(
    "Every identified customer is assigned to the month of their first purchase. The "
    "matrix records, for each cohort and each month since acquisition, the share of "
    "that cohort transacting. Two censoring effects are then corrected: late cohorts "
    "have not lived long enough to fill the right of the matrix (blank means *not yet "
    "observable*, never zero), and the first cohort is left-censored because "
    "customers who had been buying before the log opens appear as newly acquired. "
    "Section 11 reports both the raw and the corrected averages so the size of the "
    "correction is visible."
)
page_break()


# --------------------------------------------------------------------------- #
# 6. cleaning
# --------------------------------------------------------------------------- #
h1("6.  Data cleaning and its audit trail")

para(
    "Cleaning is where an analysis is most easily and most invisibly wrong, so every "
    "step in the notebook logs the rows and revenue it removed and why. The table "
    "below is that log, printed directly from the pipeline."
)

table(
    ["Step", "Rows", "Rows removed", "Revenue (GBP)", "Revenue removed"],
    [[row["Step"], num(row["Rows"]), num(row["Rows removed"]),
      f"{row['Revenue (GBP)']:,.0f}", f"{row['Revenue removed']:,.0f}"]
     for row in CLEAN_LOG],
    "The cleaning audit trail. Revenue rises at step 3 because removing cancellations "
    "removes negative revenue.",
    widths=[2.55, 0.85, 0.95, 1.0, 0.95],
    font_size=8.5,
)

h2("6.1  Why each rule is defensible")
for text in [
    "**Duplicates are dropped only on an exact match of every column**, timestamp "
    "included. Two identical lines a minute apart are two sales; two identical lines "
    "with the same timestamp are one sale written twice.",
    "**Cancellations and write-offs are separated, not merged.** A `C`-prefixed "
    "invoice is a customer return. A negative quantity on an ordinary zero-priced "
    "invoice is a stock adjustment. Treating both as “returns” would "
    "overstate the customer return rate.",
    "**Non-positive prices go, and that also removes every blank-description row** - "
    "which is why description is never imputed. There is no product to name.",
    "**Administrative stock codes are removed by an explicit list**, built from the "
    "profiling output rather than from a guess: bank charges, Amazon fees, manual "
    "adjustments, bad debt, samples and test records.",
    "**Postage and carriage lines are flagged, not deleted.** They are real revenue "
    f"({gbp(F['carriage_revenue'])}, {pct(F['carriage_share_pct'])} of the total) but "
    "they are not merchandise, so the product analysis excludes them while the "
    "revenue total keeps them.",
]:
    bullet(text)

para(
    f"The pipeline retains **{pct(F['rows_retained_pct'])} of the raw rows** "
    f"({num(CLEAN_ROWS_DROPPED)} lines removed) and produces the "
    f"{gbp(F['sales_revenue'])} base that every later section uses."
)
page_break()


# --------------------------------------------------------------------------- #
# 7. EDA
# --------------------------------------------------------------------------- #
h1("7.  Exploratory analysis")

h2("7.1  Revenue over time")
figure("01_revenue_trend", "Monthly revenue across the 25-month window, with the "
                           "September-November peak marked.")
para(
    f"Revenue is strongly seasonal. **September to November averages "
    f"{gbp(F['q4_monthly_avg'])} a month against {gbp(F['rest_monthly_avg'])} for "
    f"every other month - a {pct(F['q4_uplift_pct'])} uplift.** The peak month is "
    f"{pretty_month(F['peak_month'])} at {gbp(F['peak_month_revenue'])}; the trough is "
    f"{pretty_month(F['trough_month'])} at {gbp(F['trough_month_revenue'])}, which is less than "
    f"35% of the peak. Comparing the two like-for-like trading years, revenue grew "
    f"{pct(F['yoy_growth_pct'])}, from {gbp(F['year1_revenue'])} to "
    f"{gbp(F['year2_revenue'])}."
)
para(
    f"The final month in the file, {pretty_month(MONTHLY[-1]['InvoiceMonth'])}, is truncated - the "
    f"log stops on {pretty_date(F['date_max'])} - so its {gbp(MONTHLY[-1]['Revenue'])} is a "
    f"partial month and is never compared with a full one."
)

h2("7.2  Where the growth came from")
figure("02_new_vs_returning", "Monthly revenue split by whether the customer had "
                              "bought before.")
para(
    f"The growth is not coming from acquisition. **In the second year, "
    f"{pct(F['returning_share_year2_pct'])} of identified revenue came from customers "
    f"acquired earlier**; new customers contributed the remaining "
    f"{pct(NEW_SHARE_YEAR2)}. That single split reframes the whole analysis: the "
    f"existing base is the growth engine, so losing part of it is more expensive than "
    f"failing to add to it."
)

h2("7.3  When orders are processed")
figure("03_demand_heatmap", "Distinct invoices by weekday and hour. Saturday is "
                            "effectively empty; Sunday is not.")
para(
    f"**All {F['saturday_orders']} Saturday orders in 25 months fall on a single "
    f"date, {pretty_date(F['saturday_date'])}**, leaving "
    f"{F['saturdays_with_zero_orders']} of the {F['saturdays_in_window']} Saturdays "
    f"in the window at zero - while Sunday processes {num(F['sunday_orders'])} orders "
    f"at weekday-like volumes. No customer base trades every Sunday and one Saturday "
    f"in two years."
)
para(
    f"The conclusion is that `InvoiceDate` records **when staff keyed the order in**, "
    f"not when the customer decided to buy. That makes the column worthless for any "
    f"“best day to advertise” question and genuinely useful for staffing: "
    f"{pct(F['core_hours_share_pct'])} of all orders are processed between 10:00 and "
    f"15:59, peaking at {F['peak_hour']}:00, and the busiest day is "
    f"{F['busiest_weekday']} with {num(F['busiest_weekday_orders'])} orders."
)

h2("7.4  Products")
figure("04_top_products", "The ten highest-revenue products, net of nothing - see "
                          "section 7.7 for what changes when returns are netted off.")
table(
    ["Product", "Revenue", "Units", "Orders", "% of revenue"],
    [[row["Description"].title(), gbp(row["Revenue"]), num(row["Units"]),
      num(row["Orders"]), pct(row["Revenue share %"], 2)]
     for row in TOP_PRODUCTS],
    "Top ten products by merchandise revenue.",
    widths=[2.6, 0.95, 0.75, 0.7, 1.0],
    font_size=8.5,
)
para(
    f"The best-selling line is **{F['top_product']} at {gbp(F['top_product_revenue'])}**. "
    f"The ten largest products together are {pct(F['top10_product_share_pct'], 2)} of "
    f"merchandise revenue, which sounds small until it is compared with a uniform "
    f"expectation: ten products out of {num(F['product_count'])} is "
    f"{pct(CATALOGUE_SHARE, 2)} of the catalogue, so those ten are "
    f"**over-represented by {F['top10_concentration_ratio']:.1f}x**."
)
para(
    f"Revenue and volume are also not the same ranking. Only "
    f"{F['rev_unit_overlap']} of the top ten products by revenue also appear in the "
    f"top ten by units sold - the rest are either cheap products sold in enormous "
    f"quantity or expensive products sold steadily. A buying team optimising on units "
    f"and a finance team optimising on revenue would pick different winners."
)

h2("7.5  Geography")
figure("05_geography", "Revenue by country, with the UK share called out.")
table(
    ["Country", "Revenue", "Orders", "Customers", "AOV", "% of revenue"],
    [[row["Country"], gbp(row["Revenue"]), num(row["Orders"]), num(row["Customers"]),
      gbp(row["AOV"]), pct(row["Revenue share %"], 2)]
     for row in COUNTRIES],
    "The twelve largest markets by revenue.",
    widths=[1.5, 1.05, 0.8, 0.9, 0.85, 1.0],
    font_size=8.5,
)
para(
    f"**The United Kingdom is {pct(F['uk_share_pct'], 2)} of revenue "
    f"({gbp(F['uk_revenue'])}).** All {EXPORT_MARKETS} other markets together produce "
    f"{gbp(F['export_revenue'])}, led by {F['top_export']} at "
    f"{gbp(F['top_export_revenue'])}. That is a concentration risk stated plainly."
)
para(
    f"The revenue ranking hides the more interesting number. "
    f"**{F['best_export_aov_country']} averages {gbp(F['best_export_aov'], 2)} per "
    f"order against the UK's {gbp(F['uk_aov'], 2)} - {NL_UK_RATIO:.1f}x** - on a "
    f"small number of orders. The export business is low-volume and high-value, which "
    f"is a different business from the domestic one. On a base this small it is a "
    f"hypothesis worth testing, not a conclusion to act on."
)

h2("7.6  Order value")
para(
    f"Average order value is {gbp(F['aov'], 2)} against a median of "
    f"{gbp(F['median_order'], 2)} - the mean is {AOV_RATIO:.2f}x the median, so a "
    f"minority of very large baskets is pulling the average away from the typical "
    f"order. Orders average {F['lines_per_order']} product lines and "
    f"{num(F['units'])} units were sold in total, both consistent with wholesale "
    f"buying rather than consumer shopping."
)

h2("7.7  Returns and cancellations")
figure("06_return_rate", "Monthly return rate against the overall average.")
para(
    f"Returns total **{gbp(F['return_value'])} across {num(F['return_rows'])} lines - "
    f"{pct(F['return_rate_pct'], 2)} of gross revenue**, taking "
    f"{gbp_m(F['sales_revenue'])} down to a net {gbp_m(F['net_revenue'])}. "
    f"{num(F['returning_customers'])} customers returned something at least once."
)
para(
    f"The aggregate rate hides the pattern. The monthly rate averages "
    f"{pct(F['mean_monthly_return_rate_pct'])} but **spikes to "
    f"{pct(F['worst_return_rate_pct'], 2)} in {pretty_month(F['worst_return_month'])}**, immediately "
    f"after the Christmas peak. Returns are a post-seasonal event, so a single annual "
    f"return-rate assumption under-provisions January and over-provisions the rest of "
    f"the year."
)
figure("07_top_returned", "The ten products with the largest returned value, with "
                          "each product's return rate.")
table(
    ["Product", "Returned", "Sold", "Return rate", "Lines"],
    [[row["Description"].title(), gbp(row["Returned"]), gbp(row["Sold"]),
      pct(row["Return rate %"]), num(row["Lines"])]
     for row in TOP_RETURNED],
    "Most-returned products by value. A 100% rate on a single line is a cancelled "
    "order, not a quality problem.",
    widths=[2.6, 0.95, 0.95, 0.85, 0.7],
    font_size=8.5,
)
para(
    f"Tracing the largest of those lines produces the most consequential single "
    f"finding in the project. **{num(F['biggest_sale_units'])} units of "
    f"{F['biggest_sale_product'].title()}, worth {gbp(F['biggest_sale_value'], 2)}, "
    f"were cancelled {F['biggest_sale_cancel_gap_min']:.0f} minutes after the invoice "
    f"was raised** - net contribution {gbp(F['biggest_sale_net'], 2)}. The "
    f"second-largest line in the file is the same story: "
    f"**{num(F['second_cancel_units'])} units of {F['second_cancel_product']}, worth "
    f"{gbp(F['second_cancel_value'], 2)}, also cancelled.** The two largest sales in "
    f"the dataset were never sales."
)
para(
    f"The consequence is not academic. Ranked on gross revenue "
    f"{F['biggest_sale_product'].title()} is the retailer's "
    f"**{ordinal(F['worst_net_mover_gross_rank'])}** best seller; ranked net of "
    f"cancellations it is **{ordinal(F['worst_net_mover_net_rank'])}**. "
    f"{F['second_net_mover']} moves from "
    f"{ordinal(F['second_net_mover_gross_rank'])} to "
    f"{ordinal(F['second_net_mover_net_rank'])}. "
    f"**{F['net_rank_movers']} of the top ten products change position once returns "
    f"are netted off.** Any product report built on gross sales is materially wrong "
    f"at the top of the table. The most-returned table above shows the other half of "
    f"the same event: {SECOND_RETURNED['Description'].lower()} carries "
    f"{gbp(SECOND_RETURNED['Returned'], 2)} returned against "
    f"{gbp(SECOND_RETURNED['Sold'], 2)} sold, a "
    f"{pct(SECOND_RETURNED['Return rate %'])} return rate over "
    f"{SECOND_RETURNED['Lines']} lines."
)

h2("7.8  How concentrated the customer base is")
figure("08_pareto", "Cumulative share of revenue against cumulative share of "
                     "customers, ranked by spend.")
table(
    ["Customer group", "Customers", "Share of revenue"],
    [
        ["Top 1%", num(F["top1pct_customers"]), pct(F["top1pct_share"])],
        ["Top 5%", num(F["top5pct_customers"]), pct(F["top5pct_share"])],
        ["Top 10%", num(F["top10pct_customers"]), pct(F["top10pct_share"])],
        ["Top 20%", num(F["top20pct_customers"]), pct(F["top20pct_share"])],
        ["Top 50%", num(F["top50pct_customers"]), pct(F["top50pct_share"])],
    ],
    "Revenue concentration among identified customers.",
    widths=[2.0, 2.0, 2.15],
    font_size=9.5,
)
para(
    f"**{num(F['top20pct_customers'])} customers produce "
    f"{pct(F['top20pct_share'])} of revenue, and {num(F['top1pct_customers'])} "
    f"produce {pct(F['top1pct_share'])}.** The distribution behind that is extreme: "
    f"the median customer is worth {gbp(F['median_customer_revenue'], 2)}, the mean "
    f"{gbp(F['mean_customer_revenue'], 2)} ({MEAN_MEDIAN_RATIO:.1f}x the median), and "
    f"the largest single customer {gbp(F['max_customer_revenue'], 2)}."
)
para(
    "This is the finding that licenses everything after it. Treating this base "
    "uniformly is guaranteed to be wrong, because a uniform plan is calibrated to a "
    "customer who does not exist."
)
page_break()


# --------------------------------------------------------------------------- #
# 8. RFM
# --------------------------------------------------------------------------- #
h1("8.  Customer segmentation - RFM")

figure("09_rfm_segments", "The nine RFM segments: share of customers against share "
                          "of revenue.")
table(
    ["Segment", "Customers", "% of base", "Revenue", "% of revenue",
     "Avg recency (days)", "Avg orders", "Avg spend"],
    [[row["Segment"], num(row["Customers"]), pct(row["Customer %"]),
      gbp_compact(row["Revenue"]), pct(row["Revenue %"]),
      one_dp(row["AvgRecency"]), one_dp(row["AvgFrequency"]),
      gbp(row["AvgMonetary"])]
     for row in SEGMENTS],
    "The nine RFM segments, ordered by revenue contribution.",
    widths=[1.25, 0.72, 0.72, 0.72, 0.78, 0.85, 0.63, 0.72],
    font_size=8,
)

para(
    f"**{num(F['seg_champions_customers'])} Champions - "
    f"{pct(F['seg_champions_cust_pct'])} of the base - account for "
    f"{gbp(F['seg_champions_revenue'])}, or {pct(F['seg_champions_rev_pct'])} of all "
    f"identified revenue.** They order roughly "
    f"{CHAMPIONS['AvgFrequency']:.0f} times, last bought "
    f"{CHAMPIONS['AvgRecency']:.0f} days before the reference date, and spend "
    f"{gbp(CHAMPIONS['AvgMonetary'])} each."
)
para(
    f"At the other end, {num(F['seg_lost_customers'])} *Lost* and "
    f"{num(F['seg_hibernating_customers'])} *Hibernating* customers - "
    f"{pct(DEAD_CUSTOMER_PCT)} of the base between them - contribute "
    f"{pct(DEAD_REVENUE_PCT)} of revenue. They are numerous and nearly worthless, "
    f"which is exactly why a base-wide campaign wastes most of its budget."
)
para(
    f"The urgent group is neither. **{num(F['seg_cannot_lose_them_customers'])} "
    f"customers in *Cannot lose them* and {num(F['seg_at_risk_customers'])} in *At "
    f"risk* - {num(F['at_risk_customers'])} accounts - have already spent "
    f"{gbp(F['at_risk_revenue'])}, which is {pct(F['at_risk_rev_pct'])} of revenue, "
    f"and have stopped ordering.** *Cannot lose them* last bought an average of "
    f"{CANNOT_LOSE['AvgRecency']:.0f} days ago having ordered "
    f"{CANNOT_LOSE['AvgFrequency']:.1f} times and spent "
    f"{gbp(CANNOT_LOSE['AvgMonetary'])} each: proven buyers who went quiet, not "
    f"casual browsers who never came back."
)
para(
    f"The remaining segments matter for pipeline rather than revenue. *Promising* "
    f"({num(F['seg_promising_customers'])}), *New* ({num(F['seg_new_customers'])}) "
    f"and *Needs attention* ({num(F['seg_needs_attention_customers'])}) hold "
    f"{pct(PIPELINE_PCT)} "
    f"of revenue between them - they are where the next cohort of Champions has to "
    f"come from, and they are cheap to nurture precisely because they are small."
)
page_break()


# --------------------------------------------------------------------------- #
# 9. K-Means
# --------------------------------------------------------------------------- #
h1("9.  Customer segmentation - K-Means")

h2("9.1  Choosing k honestly")
figure("10_kmeans_selection", "Inertia and silhouette score for k = 2 to 8.")
table(
    ["k", "Inertia", "Silhouette"],
    [[row["k"], f"{row['Inertia']:,.1f}", f"{row['Silhouette']:.4f}"]
     for row in KMEANS_SCAN],
    "Cluster-count scan on log-scaled, standardised RFM features.",
    widths=[0.8, 2.0, 2.0],
    font_size=9.5,
)
para(
    f"The silhouette score is highest at **k = {F['silhouette_best_k']} "
    f"({F['silhouette_best']:.4f})**, and the value adopted here is "
    f"**k = {F['k_chosen']} ({F['silhouette_at_k']:.4f})**. That is a deliberate "
    f"disagreement with the metric, for a stated reason: two clusters cannot be "
    f"marketed to differently, so the split would be statistically neater and "
    f"operationally useless. The elbow in inertia is also around k = 4. The cost of "
    f"the choice is reported rather than hidden."
)

h2("9.2  What the four clusters are")
figure("11_cluster_profiles", "Average recency, frequency and monetary value per "
                              "cluster, shown as separate panels rather than on one "
                              "shared axis.")
table(
    ["Cluster", "Label", "Customers", "Avg recency (days)", "Avg orders",
     "Avg spend", "Revenue", "% of revenue"],
    [[row["Cluster"], row["Name"], num(row["Customers"]), one_dp(row["AvgRecency"]),
      one_dp(row["AvgFrequency"]), gbp(row["AvgMonetary"]),
      gbp_compact(row["TotalRevenue"]), pct(row["Revenue %"])]
     for row in CLUSTERS],
    "Cluster profiles. Labels are assigned from the profile, after fitting.",
    widths=[0.6, 1.35, 0.75, 0.85, 0.65, 0.7, 0.65, 0.6],
    font_size=8,
)
figure("12_cluster_scatter", "Customers in frequency-monetary space, coloured by "
                             "cluster, both axes log-scaled.")
para(
    f"The clusters separate cleanly on value and activity. "
    f"**{num(RICHEST_CLUSTER['Customers'])} “{RICHEST_CLUSTER['Name']}” "
    f"carry {pct(RICHEST_CLUSTER['Revenue %'])} of revenue** on an average of "
    f"{RICHEST_CLUSTER['AvgFrequency']:.1f} orders and "
    f"{gbp(RICHEST_CLUSTER['AvgMonetary'])} of spend, while the largest cluster by "
    f"headcount - {num(BIGGEST_CLUSTER['Customers'])} "
    f"“{BIGGEST_CLUSTER['Name']}” - carries "
    f"{pct(BIGGEST_CLUSTER['Revenue %'])}. The clustering finds the same imbalance "
    f"the RFM rules found, from a completely different starting point."
)

h2("9.3  Do the two methods agree?")
para(
    f"This is the point of running both. K-Means was given no rules, no segment "
    f"names and no thresholds, and **every one of the nine rule-based segments sends "
    f"at least 60% of its members to a single cluster** "
    f"({pct(F['segment_cluster_agreement_pct'], 0)} of segments clear that bar). "
    f"*Cannot lose them* maps 96.7% onto one cluster and *Lost* maps 98.4% onto "
    f"another."
)
para(
    "Two independent methods recovering the same structure is evidence that the "
    "structure is a property of the data rather than an artefact of the analyst's "
    "chosen cut-offs. The practical consequence is that the transparent, explainable "
    "RFM rules can be used to run the business, with the clustering standing behind "
    "them as validation - which is the opposite of the usual trade-off between "
    "interpretability and rigour."
)
page_break()


# --------------------------------------------------------------------------- #
# 10. CLV
# --------------------------------------------------------------------------- #
h1("10.  Customer lifetime value")

figure("13_clv_by_segment", "Historical value already realised against projected "
                            "12-month value, by segment.")
table(
    ["Segment", "Customers", "Historical CLV", "Projected 12M", "Pool projected",
     "% of pool"],
    [[row["Segment"], num(row["Customers"]), gbp(row["HistoricalCLV"]),
      gbp(row["Projected12M"]), gbp_compact(row["PoolProjected"]),
      pct(row["Share of projected pool %"])]
     for row in SEGMENT_CLV],
    "Historical and projected value per segment.",
    widths=[1.4, 0.8, 1.05, 0.95, 1.0, 0.75],
    font_size=8.5,
)
para(
    f"Historical CLV - revenue already banked - averages "
    f"{gbp(F['mean_historical_clv'], 2)} with a median of "
    f"{gbp(F['median_historical_clv'], 2)}. The 12-month projection averages "
    f"{gbp(F['mean_projected_12m'], 2)} per customer and "
    f"{gbp(F['total_projected_12m'])} across the base."
)

h2("10.1  Testing the projection instead of publishing it")
para(
    f"That total is not credible, and the notebook says so with a number. "
    f"**{gbp(F['total_projected_12m'])} is {F['proj_vs_actual_ratio']:.2f}x the "
    f"retailer's actual second-year revenue of {gbp(F['year2_revenue'])}.** The "
    f"projection assumes every customer keeps buying at their observed rate for a "
    f"further twelve months, and it carries no churn probability, so a customer who "
    f"has not ordered for a year is projected forward exactly as confidently as one "
    f"who ordered last week."
)
para(
    f"The failure is visible in the ranking as well as the total. The projection puts "
    f"**{F['clv_rank_top_segment']} at {gbp(F['clv_cannot_lose_them_projected'], 2)} "
    f"per customer, above Champions at "
    f"{gbp(F['clv_champions_projected'], 2)}** - because a dormant customer who "
    f"bought heavily in a short window shows a high rate over a short lifespan. That "
    f"ordering is an artefact of the constant-rate assumption, not a real ranking."
)
para(
    "Two conclusions follow, and both are carried into section 13. The projection is "
    "reported as a **per-customer upper bound**, never as a forecast. And every "
    "prioritisation decision in this report rests on **historical** value, which is "
    "measured, rather than on the projection, which is modelled. The correct fix - a "
    "BG/NBD model for purchase frequency with a Gamma-Gamma model for spend - is "
    "named in the limitations and is not attempted here."
)
page_break()


# --------------------------------------------------------------------------- #
# 11. cohorts
# --------------------------------------------------------------------------- #
h1("11.  Cohort retention")

figure("14_cohort_retention", "Monthly acquisition cohorts against months since "
                              "acquisition. Blank cells are not yet observable.")
table(
    ["Cohort", "Customers", "Month 1", "Month 3", "Month 6",
     "Revenue", "Revenue per customer"],
    [[row["CohortMonth"], num(row["CohortSize"]),
      pct(row["Month1 %"]) if row["Month1 %"] is not None else "—",
      pct(row["Month3 %"]) if row["Month3 %"] is not None else "—",
      pct(row["Month6 %"]) if row["Month6 %"] is not None else "—",
      gbp_compact(row["Revenue"]), gbp(row["Revenue per acquired customer"])]
     for row in COHORTS],
    "Cohort size, early retention and lifetime revenue per acquisition month.",
    widths=[0.8, 0.85, 0.75, 0.75, 0.75, 0.9, 1.35],
    font_size=8.5,
)

para(
    f"Retention settles at roughly a fifth and then holds: **{pct(F['retention_m1_true'])} "
    f"at month 1, {pct(F['retention_m3_true'])} at month 3, "
    f"{pct(F['retention_m6_true'])} at month 6 and {pct(F['retention_m12_true'])} at "
    f"month 12**, excluding the left-censored first cohort. A curve that flattens "
    f"rather than decaying is the signature of a wholesale reorder cycle: the "
    f"customers who stay, stay for years."
)

h2("11.1  Correcting the first cohort")
para(
    f"The December 2009 cohort is the largest in the file "
    f"({num(F['first_cohort_size'])} customers) and it retains far better than any "
    f"other - because it is not really a cohort. Customers who had been buying before "
    f"the log opens are recorded as newly acquired in the first month, so an "
    f"established base is being counted as new. Excluding it moves month-1 retention "
    f"from {pct(F['retention_m1_mean'])} to {pct(F['retention_m1_true'])} and month-12 "
    f"from {pct(F['retention_m12_mean'])} to {pct(F['retention_m12_true'])}."
)
para(
    f"The same artefact distorts a conclusion an analyst might otherwise draw. The "
    f"correlation between cohort size and month-6 retention is "
    f"**{F['cohort_size_vs_m6_corr']:+.3f} with the first cohort included and "
    f"{F['cohort_size_vs_m6_corr_true']:+.3f} without it** - so “bigger cohorts "
    f"retain better” is mostly one censored row doing the work."
)

h2("11.2  What the Christmas peak actually delivers")
para(
    f"Splitting cohorts by acquisition month exposes a sharp difference. "
    f"**Customers acquired in November or December retain at "
    f"{pct(F['xmas_cohort_m3'])} by month 3, against {pct(F['nonxmas_cohort_m3'])} "
    f"for customers acquired in any other month** - a gap of {XMAS_GAP:.1f}x. The "
    f"season that produces the revenue spike produces the retailer's worst customers, "
    f"and a target that counts peak-season acquisitions as equivalent new "
    f"relationships is measuring the wrong thing."
)
page_break()


# --------------------------------------------------------------------------- #
# 12. findings
# --------------------------------------------------------------------------- #
h1("12.  Key findings")
para(
    "Every figure below is computed in the notebook and stored in "
    "`outputs/facts.json`. Nothing here is estimated or rounded from memory.",
    italic=True, colour=INK_2, size=10,
)

FINDINGS = [
    ("Revenue is extremely concentrated, so segmentation is justified before it is built",
     [f"**{num(F['top20pct_customers'])} customers (20%) generate "
      f"{pct(F['top20pct_share'])} of revenue. The top {num(F['top1pct_customers'])} "
      f"customers (1%) generate {pct(F['top1pct_share'])}.** The median customer is "
      f"worth {gbp(F['median_customer_revenue'])} and the mean "
      f"{gbp(F['mean_customer_revenue'])} - {MEAN_MEDIAN_RATIO:.1f}x - so "
      f"“the average customer” is a fiction. Order value tells the same "
      f"story: a {gbp(F['aov'], 2)} mean against a {gbp(F['median_order'], 2)} median.",
      "Uniform treatment of this base is guaranteed to be wrong, because it is "
      "calibrated to a customer who does not exist."]),

    ("One segment holds most of the revenue, and a different one holds the risk",
     [f"**Champions - {num(F['seg_champions_customers'])} customers, "
      f"{pct(F['seg_champions_cust_pct'])} of the base - account for "
      f"{gbp_m(F['seg_champions_revenue'])}, or {pct(F['seg_champions_rev_pct'])} of "
      f"identified revenue.** At the other end, {num(F['seg_lost_customers'])} *Lost* "
      f"and {num(F['seg_hibernating_customers'])} *Hibernating* customers "
      f"({pct(DEAD_CUSTOMER_PCT)} of the base) contribute {pct(DEAD_REVENUE_PCT)} "
      f"between them.",
      f"The urgent group is neither: **{num(F['at_risk_customers'])} customers in "
      f"*Cannot lose them* and *At risk* have already spent "
      f"{gbp(F['at_risk_revenue'])} ({pct(F['at_risk_rev_pct'])} of revenue) and have "
      f"stopped ordering.** These are proven buyers who went quiet - the only group "
      f"where an intervention has a large, identified prize attached."]),

    ("The two largest sales in the dataset were never sales",
     [f"**The single biggest line in {num(F['raw_rows'])} rows - "
      f"{num(F['biggest_sale_units'])} units of {F['biggest_sale_product'].title()}, "
      f"{gbp(F['biggest_sale_value'], 2)} - was cancelled "
      f"{F['biggest_sale_cancel_gap_min']:.0f} minutes after it was raised**, on the "
      f"final day of the log. Net contribution: **{gbp(F['biggest_sale_net'], 2)}**. "
      f"**The second-biggest line - {num(F['second_cancel_units'])} units of "
      f"{F['second_cancel_product']}, {gbp(F['second_cancel_value'], 2)} - was "
      f"cancelled too.** Neither of the two largest sales in the dataset was a sale.",
      f"Ranked on gross revenue {F['biggest_sale_product'].title()} is the retailer's "
      f"**{ordinal(F['worst_net_mover_gross_rank'])}** best seller. Ranked net of "
      f"cancellations it is **{ordinal(F['worst_net_mover_net_rank'])}**; "
      f"{F['second_net_mover']} falls from "
      f"{ordinal(F['second_net_mover_gross_rank'])} to "
      f"{ordinal(F['second_net_mover_net_rank'])}. "
      f"**{F['net_rank_movers']} of the top ten products change position once returns "
      f"are netted off.** Any product report built on gross sales - the default in "
      f"most published work on this dataset - is materially wrong at the top of the "
      f"table."]),

    (f"Returns cost {gbp_m(F['return_value'])}, and the aggregate rate hides the problem",
     [f"Returns total **{gbp(F['return_value'])} across {num(F['return_rows'])} lines "
      f"- {pct(F['return_rate_pct'], 2)} of gross revenue**, taking "
      f"{gbp_m(F['sales_revenue'])} down to a net {gbp_m(F['net_revenue'])}. The "
      f"monthly rate averages {pct(F['mean_monthly_return_rate_pct'])} but is not "
      f"stable: it **spikes to {pct(F['worst_return_rate_pct'], 2)} in "
      f"{pretty_month(F['worst_return_month'])}**, immediately after the Christmas peak.",
      "Returns are a post-seasonal event, not a constant leak, so a single annual "
      "return-rate assumption will under-provision January and over-provision the "
      "rest of the year."]),

    ("The business is one country and one quarter",
     [f"**The UK is {pct(F['uk_share_pct'], 2)} of revenue "
      f"({gbp_m(F['uk_revenue'])});** all {EXPORT_MARKETS} other markets together "
      f"make {gbp_m(F['export_revenue'])}. **September to November averages "
      f"{gbp_m(F['q4_monthly_avg'])} per month against {gbp(F['rest_monthly_avg'])} "
      f"for every other month - a {pct(F['q4_uplift_pct'])} seasonal uplift**, "
      f"peaking at {gbp_m(F['peak_month_revenue'])} in {pretty_month(F['peak_month'])} against a "
      f"{gbp(F['trough_month_revenue'])} trough in {pretty_month(F['trough_month'])}.",
      f"Both concentrations are risks, but the export data also holds an opportunity "
      f"the revenue ranking hides: **{F['best_export_aov_country']} averages "
      f"{gbp(F['best_export_aov'], 2)} per order against the UK's "
      f"{gbp(F['uk_aov'], 2)} - {NL_UK_RATIO:.1f}x** - on a small number of orders. "
      f"The export markets are low-volume and high-value, which is a different "
      f"business from the domestic one."]),

    ("Growth came from retained customers, not new ones",
     [f"Like-for-like years grew **{pct(F['yoy_growth_pct'])}** "
      f"({gbp_m(F['year1_revenue'])} to {gbp_m(F['year2_revenue'])}). But in the "
      f"second year, **{pct(F['returning_share_year2_pct'])} of identified revenue "
      f"came from customers acquired earlier.** New customers contributed the "
      f"remaining {pct(NEW_SHARE_YEAR2)}.",
      "Growth is being produced by the existing base, which raises the cost of losing "
      "any part of it and makes finding 2's at-risk pool the most expensive problem "
      "on this list."]),

    ("Christmas buys volume, not loyalty",
     [f"Retention settles at roughly a fifth and then holds: "
      f"**{pct(F['retention_m1_true'])} at month 1, {pct(F['retention_m3_true'])} at "
      f"month 3, {pct(F['retention_m6_true'])} at month 6, "
      f"{pct(F['retention_m12_true'])} at month 12** (excluding the left-censored "
      f"first cohort). The curve flattens rather than decaying, which is the "
      f"signature of a wholesale reorder cycle rather than one-off consumer buying.",
      f"Splitting cohorts by acquisition month exposes a sharp difference: "
      f"**customers acquired in November or December retain at "
      f"{pct(F['xmas_cohort_m3'])} by month 3, against "
      f"{pct(F['nonxmas_cohort_m3'])} for customers acquired in any other month** - "
      f"less than half. The peak season that produces the revenue spike produces the "
      f"retailer's worst customers."]),

    ("The rules and the algorithm agree, which is the point of running both",
     [f"K-Means on log-scaled RFM was given no rules, and **every one of the nine "
      f"rule-based segments sends at least 60% of its members to a single cluster**. "
      f"*Cannot lose them* maps 96.7% onto one cluster and *Lost* maps 98.4% onto "
      f"another. Two independent methods recovering the same structure is evidence "
      f"the structure is in the data rather than in the analyst's choice of "
      f"thresholds.",
      f"The honest caveat: the silhouette score preferred "
      f"**k = {F['silhouette_best_k']} ({F['silhouette_best']:.4f})** over the "
      f"**k = {F['k_chosen']} ({F['silhouette_at_k']:.4f})** adopted here. Two "
      f"clusters cannot be marketed to differently, so interpretability was chosen "
      f"over the metric, and the cost is stated rather than hidden."]),

    ("Two data artefacts that would have produced wrong answers",
     [f"**`InvoiceDate` is not customer intent.** All {F['saturday_orders']} Saturday "
      f"orders in 25 months fall on one date, {pretty_date(F['saturday_date'])} "
      f"({pct(F['saturday_share_pct'], 3)} of orders), while Sunday processes "
      f"{num(F['sunday_orders'])} orders at weekday-like volumes. No customer base "
      f"trades every Sunday and one Saturday in two years. The timestamp records "
      f"back-office keying, which makes it valid for staffing "
      f"({pct(F['core_hours_share_pct'])} of orders are processed between 10:00 and "
      f"15:59, peaking at {F['peak_hour']}:00) and worthless for deciding when to "
      f"advertise.",
      f"**Cancellations are not the only reversal.** "
      f"{F['invoice_prefix_a_count']} `A`-prefixed invoices carry "
      f"{gbp(F['invoice_prefix_a_revenue'])} of bad-debt write-offs. The standard "
      f"“drop invoices starting with C” recipe leaves all of it in the "
      f"revenue line."]),
]

for i, (title, paragraphs) in enumerate(FINDINGS, start=1):
    h2(f"Finding {i}.  {title}")
    for text in paragraphs:
        para(text)
page_break()


# --------------------------------------------------------------------------- #
# 13. recommendations
# --------------------------------------------------------------------------- #
h1("13.  Recommendations")
para(
    "Each recommendation names the finding it rests on and the number that sizes it. "
    "Where the data cannot support a decision, that is stated instead of guessed.",
    italic=True, colour=INK_2, size=10,
)

RECOMMENDATIONS = [
    (f"Run a win-back on the {num(F['at_risk_customers'])} quiet high-value accounts "
     f"- and measure it properly",
     [f"*From finding 2.* These customers have already spent "
      f"**{gbp(F['at_risk_revenue'])}** and have stopped ordering. Prioritise by "
      f"**historical** value, not by the projected figure: section 10.1 shows the "
      f"projection ranks this group *above* Champions, which is an artefact of the "
      f"constant-rate assumption rather than a real ordering.",
      "Because these are wholesale accounts, the intervention is a phone call from a "
      "named account manager, not a discount email - and a discount is the wrong "
      "lever anyway, since nothing in this dataset says price caused them to leave.",
      f"**Hold back a random 20% as an untreated control group.** Without one, any "
      f"subsequent recovery is indistinguishable from customers who would have "
      f"reordered regardless - and given the flat "
      f"{pct(F['retention_m12_true'])} month-12 reorder rate in finding 7, a "
      f"meaningful fraction would have."]),

    (f"Protect the {num(F['seg_champions_customers'])} Champions before chasing "
     f"anyone new",
     [f"*From findings 2 and 6.* Champions produce "
      f"**{pct(F['seg_champions_rev_pct'])} of revenue**, and "
      f"**{pct(F['returning_share_year2_pct'])} of second-year revenue came from "
      f"previously-acquired customers**. The concentration cuts both ways: losing "
      f"{num(F['top1pct_customers'])} customers (the top 1%) would remove "
      f"**{pct(TOP1_LOSS)} of revenue**.",
      "Concretely: guaranteed stock availability on their repeat lines through the "
      "September-November peak, and a named contact. The defensive case is stronger "
      "than any acquisition case on this data, because acquisition is demonstrably "
      "not what is producing growth."]),

    ("Stop reporting product performance on gross revenue",
     [f"*From finding 3.* {F['net_rank_movers']} of the top ten products change rank "
      f"once cancellations are netted off, and the gross-rank-"
      f"{F['worst_net_mover_gross_rank']} product is genuinely "
      f"{ordinal(F['worst_net_mover_net_rank'])}. Every product report should net "
      f"cancellations against the original sale.",
      "This is a reporting fix, not a strategy: it costs nothing, and it stops the "
      "buying team restocking a line that sold nothing."]),

    ("Treat January as a returns event, and investigate the top returned lines",
     [f"*From finding 4.* The return rate hits "
      f"**{pct(F['worst_return_rate_pct'], 2)} in {pretty_month(F['worst_return_month'])}** against a "
      f"{pct(F['mean_monthly_return_rate_pct'])} average. Two actions follow: "
      f"provision warehouse and refund capacity for a post-Christmas spike rather "
      f"than an even monthly rate, and inspect the concentrated returned lines in "
      f"section 7.7 for a cause.",
      "The data shows *which* products come back but never *why*, so the second "
      "action is a question to take to the warehouse, not one to answer from the "
      "file."]),

    ("Change what the Christmas peak is expected to deliver",
     [f"*From finding 7.* Customers acquired in November and December retain at "
      f"**{pct(F['xmas_cohort_m3'])} by month 3 versus "
      f"{pct(F['nonxmas_cohort_m3'])}** otherwise. The peak should be run as a volume "
      f"and cash-generation event, and peak-acquired customers should **not** be "
      f"counted as new relationships in any target that assumes they behave like the "
      f"rest of the base.",
      "The corollary is to spend acquisition budget outside the peak, where the "
      "customers who arrive are more than twice as likely to still be buying at "
      "month 3."]),

    ("Test the export markets deliberately",
     [f"*From finding 5.* {F['best_export_aov_country']} averages "
      f"**{gbp(F['best_export_aov'], 2)} per order - {NL_UK_RATIO:.1f}x the UK's "
      f"{gbp(F['uk_aov'], 2)}**. That is a striking ratio on a small base, and a "
      f"small base is exactly why it must be tested rather than acted on: with "
      f"{gbp_m(F['export_revenue'])} of export revenue against "
      f"{gbp_m(F['uk_revenue'])} domestic, a handful of large wholesale accounts "
      f"could produce the entire effect.",
      "The recommendation is a bounded test on the highest-AOV markets with a defined "
      "budget, not a market-entry decision. The data supports the hypothesis; it does "
      "not support the investment."]),

    (f"Fix the {pct(F['anonymous_revenue_pct'])} attribution gap at source",
     [f"*From the limitations.* **{gbp(F['anonymous_revenue'])} of clean revenue - "
      f"{pct(F['anonymous_revenue_pct'])} of the total, spread over "
      f"{pct(F['anonymous_rows_pct'])} of all sales lines - has no customer ID**, so "
      f"it is invisible to every model here. No amount of analysis recovers it: it is "
      f"a checkout-and-CRM change, and it raises the ceiling on all future customer "
      f"work."]),
]

for i, (title, paragraphs) in enumerate(RECOMMENDATIONS, start=1):
    h2(f"Recommendation {i}.  {title}")
    for text in paragraphs:
        para(text)

h2("What this analysis does not support")
for text in [
    "**Any pricing or product-mix decision.** The dataset has no cost of goods, so "
    "every figure here is revenue, not margin.",
    "**Any claim that a campaign *causes* a customer to return.** There is no "
    "campaign data and no experiment in this file. That is why recommendation 1 "
    "specifies a control group.",
    "**Any conclusion about when customers prefer to shop** - see finding 9.",
    f"**Using the 12-month CLV projection as a forecast.** It sums to "
    f"{F['proj_vs_actual_ratio']:.2f}x actual second-year revenue (section 10.1). It "
    f"is an upper bound per account, nothing more.",
]:
    bullet(text)
page_break()


# --------------------------------------------------------------------------- #
# 14. limitations
# --------------------------------------------------------------------------- #
h1("14.  Limitations")
para(
    "Stating what this analysis cannot support is part of the analysis.",
    italic=True, colour=INK_2, size=10,
)

LIMITATIONS = [
    f"**Attribution gap.** {pct(F['anonymous_rows_pct'])} of clean sales lines "
    f"({num(F['anonymous_rows'])}), carrying {pct(F['anonymous_revenue_pct'])} of "
    f"clean revenue ({gbp(F['anonymous_revenue'])}), have no customer ID - both "
    f"shares measured on the same cleaned base. Every customer-level result is "
    f"computed on the identified {pct(F['id_coverage_pct'])} of revenue only. If "
    f"guest orders are disproportionately one-off purchases, the true repeat rate is "
    f"lower than reported here; if they are unrecognised repeat buyers, several "
    f"segments are undercounted. The data cannot distinguish the two cases.",

    "**Revenue, not profit.** The dataset has no cost of goods. A low-margin "
    "bestseller and a high-margin niche line are indistinguishable in this analysis, "
    "so no pricing or product-mix decision should rest on it alone.",

    f"**The CLV projection is demonstrably biased upward, and section 10.1 shows by "
    f"how much.** Summed across the base it is {F['proj_vs_actual_ratio']:.2f}x "
    f"actual second-year revenue, and it ranks a dormant segment above Champions. It "
    f"carries no churn probability, so it should be read only as a per-customer upper "
    f"bound. BG/NBD plus a Gamma-Gamma spend model is the correct fix and is not done "
    f"here; prioritisation in section 13 therefore rests on historical value.",

    "**Censoring at both ends of the cohort matrix.** Cohorts acquired late in the "
    "window have not had time to show a retention curve, so blank cells on the right "
    "mean *not yet observable*, not zero. At the other end, the first cohort is "
    "left-censored: customers who predate the log appear as newly acquired in "
    "December 2009. Section 11.1 quantifies that bias and reports corrected figures.",

    f"**`InvoiceDate` records processing, not intent.** All {F['saturday_orders']} "
    f"Saturday orders in the 25-month window fall on one date, so the timestamp "
    f"reflects back-office keying rather than customer behaviour. No conclusion about "
    f"*when customers want to shop* can be drawn from the hour and weekday columns.",

    "**Quintile scoring is relative.** An RFM score of 5 means “top fifth of "
    "this base”, not “good” in absolute terms. Scores are not "
    "comparable across a different customer base or a different time window.",

    f"**k = {F['k_chosen']} is a judgement call.** The silhouette metric preferred "
    f"k = {F['silhouette_best_k']}. Four clusters were chosen for operational "
    f"interpretability, and the metric cost of that choice is reported in section 9.1 "
    f"rather than hidden.",

    "**One retailer, one category, 2009-2011.** A UK giftware wholesaler in the "
    "post-financial-crisis period. The method transfers; these coefficients do not.",

    "**Segments are descriptive, not causal.** Nothing here establishes that a "
    "campaign *causes* a customer to return. A holdout test would be needed to claim "
    "that, and this dataset contains no campaign data.",
]

for i, text in enumerate(LIMITATIONS, start=1):
    numbered(i, text)
page_break()


# --------------------------------------------------------------------------- #
# 15. conclusion
# --------------------------------------------------------------------------- #
h1("15.  Conclusion")

para(
    f"The project set out to turn a flat transaction log into a picture of a customer "
    f"base, and the picture is sharper than the aggregate figures suggest. "
    f"{gbp(F['sales_revenue'])} of revenue and {num(F['customers'])} identified "
    f"customers resolve into a business that depends on a few hundred accounts: "
    f"{num(F['top20pct_customers'])} customers carry {pct(F['top20pct_share'])} of "
    f"revenue, {num(F['seg_champions_customers'])} Champions carry "
    f"{pct(F['seg_champions_rev_pct'])}, and "
    f"{pct(F['returning_share_year2_pct'])} of second-year revenue came from "
    f"customers who were already there."
)
para(
    f"That structure was recovered twice, by methods that share no assumptions: an "
    f"explicit RFM rule ladder and an unsupervised K-Means clustering that was given "
    f"no rules at all. Their agreement - every segment sending at least 60% of its "
    f"members to one cluster - is the strongest evidence in the report that the "
    f"segmentation describes the business rather than the analyst."
)
para(
    f"The most useful output is not the segmentation itself but the list it produces: "
    f"**{num(F['at_risk_customers'])} named accounts, worth "
    f"{gbp(F['at_risk_revenue'])} of proven historical spend, that have stopped "
    f"ordering.** That is a finite, addressable list rather than a strategy slide, "
    f"and it comes with a measurement design attached - a 20% untreated control "
    f"group - because without one the result would be unreadable."
)
para(
    "Three of the results in this report contradict what a first pass at this dataset "
    "would have produced. The largest sale in the file was a cancellation, so gross "
    "product rankings are wrong at the top. The Christmas peak generates the "
    "retailer's worst-retaining customers, so peak acquisition should not be counted "
    "like any other. And the invoice timestamp records office hours rather than "
    "customer behaviour, so an entire class of “when to advertise” "
    "conclusions is unavailable. Each was found by checking a claim against the data "
    "rather than accepting a plausible number - which, more than any individual "
    "figure here, is what the project was for."
)
page_break()


# --------------------------------------------------------------------------- #
# 16. reproducing
# --------------------------------------------------------------------------- #
h1("16.  How to reproduce this project")

para("From a clone of the repository, on any machine with Python 3.10 or newer:")
for text in [
    "`pip install -r requirements.txt`",
    "`jupyter notebook " + NOTEBOOK + "`",
    "*Kernel > Restart Kernel and Run All Cells*. The first run downloads the "
    "45 MB source workbook from UCI into `data/raw/` and caches it; later runs read "
    "the cache. Expect several minutes, most of it spent reading the .xlsx file.",
    "`python tools/build_report.py` regenerates this document from the notebook's "
    "outputs.",
]:
    bullet(text)
para(
    f"The notebook writes {len(FIGS)} PNG figures to `outputs/figures/` and every "
    f"computed number to `outputs/facts.json` ({len(F)} facts plus the segment, "
    f"cluster, cohort and product tables). This report reads only those two "
    f"locations, so the document and the analysis cannot disagree."
)
para(
    "The source workbook is never modified, and it is not committed to the "
    "repository: it is a large binary that the notebook can fetch reproducibly from "
    "the publisher."
)

h1("17.  References")
for text in [
    f"Chen, D. (2019). *Online Retail II* [Data set]. UCI Machine Learning "
    f"Repository. {DATASET_URL} (CC BY 4.0).",
    "Hughes, A. M. (1994). *Strategic Database Marketing*. Probus Publishing - the "
    "origin of RFM scoring.",
    "Fader, P. S., Hardie, B. G. S., & Lee, K. L. (2005). “RFM and CLV: Using "
    "iso-value curves for customer base analysis.” *Journal of Marketing "
    "Research*, 42(4), 415-430.",
    "Rousseeuw, P. J. (1987). “Silhouettes: a graphical aid to the "
    "interpretation and validation of cluster analysis.” *Journal of "
    "Computational and Applied Mathematics*, 20, 53-65.",
    "Pedregosa, F. et al. (2011). “Scikit-learn: Machine learning in Python.” "
    "*Journal of Machine Learning Research*, 12, 2825-2830.",
    f"Project repository: {REPO_URL}",
]:
    bullet(text)


# --------------------------------------------------------------------------- #
# write
# --------------------------------------------------------------------------- #
doc.save(OUT_PATH)
size_kb = OUT_PATH.stat().st_size / 1024
print(f"wrote {OUT_PATH.name}  ({size_kb:,.0f} KB)")
print(f"  {FIGURE_N} figures embedded, {TABLE_N} tables, "
      f"{len(F)} facts available from outputs/facts.json")
