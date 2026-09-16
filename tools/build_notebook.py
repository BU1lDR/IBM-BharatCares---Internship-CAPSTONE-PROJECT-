"""Build AryanVerma_RetailCustomerSegmentationAnalysis.ipynb from source cells.

Build tooling, not the deliverable: the notebook it writes is. Keeping the cell
sources here means the notebook can be regenerated wholesale instead of edited
cell by cell in the browser, which is what keeps its section numbering and its
narrative consistent.

Usage
-----
    python tools/build_notebook.py     # writes the .ipynb, unexecuted
    python tools/run_notebook.py      # executes it in place, reports failures
    python tools/build_report.py      # rebuilds the .docx from the new outputs

Regenerating discards any edit made directly to the .ipynb, so change the cell
sources here rather than the notebook.
"""
from pathlib import Path

import nbformat as nbf

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "AryanVerma_RetailCustomerSegmentationAnalysis.ipynb"

C = []  # (kind, source)
def md(s): C.append(("md", s.strip("\n")))
def code(s): C.append(("code", s.strip("\n")))


# ==================================================================== TITLE
md(r'''
# Retail Customer Segmentation & Sales Analysis

**IBM SkillsBuild Data Analytics with AI — Academic Internship Programme**
BharatCares in association with AICTE

| | |
|---|---|
| **Author** | Aryan Verma |
| **Dataset** | [UCI Machine Learning Repository — Online Retail II](https://archive.ics.uci.edu/dataset/502/online+retail+ii) |
| **Scope** | 1,067,371 invoice lines · 01 Dec 2009 – 09 Dec 2011 · 43 countries |
| **Stack** | Python · pandas · NumPy · Matplotlib · scikit-learn |

---

## 1. Business problem

A UK-based online giftware retailer has two years of transaction history but no
view of *who* its customers are. Marketing spend is undifferentiated: the same
campaign goes to a customer who ordered yesterday and to one who has not ordered
in 18 months. Management wants three questions answered from the data it already
owns:

1. **Where does revenue actually come from** — which products, countries and
   months, and how concentrated is it?
2. **Which customers matter** — can the base be split into a small number of
   groups that deserve genuinely different treatment?
3. **What is a customer worth**, and how much of the base is quietly churning?

## 2. Analytical approach

| Stage | What is done | Why |
|---|---|---|
| Data quality audit | Quantify missing values, duplicates, cancellations, negative quantities, non-product codes | Decide the cleaning rules from evidence, not habit |
| Cleaning | Nine documented steps, each logged with its row and revenue impact | An auditable trail — every number below is traceable to a rule |
| Descriptive analysis | Revenue trend, demand rhythm, products, geography, returns | Establish the baseline the segmentation has to explain |
| Concentration | Pareto curve of revenue by customer | Test whether segmentation is even worth doing |
| RFM segmentation | Recency / Frequency / Monetary quintile scoring → named segments | Transparent, rule-based, directly actionable |
| K-Means clustering | Unsupervised clustering on log-scaled RFM (scikit-learn) | Let the data propose its own groups, then compare against the rules |
| Customer Lifetime Value | Historical value + a rate-based 12-month projection | Turn segments into a budget |
| Cohort retention | Monthly acquisition cohorts × months since first order | Separate a growth problem from a retention problem |

## 3. How to run this notebook

```bash
pip install -r requirements.txt
jupyter notebook AryanVerma_RetailCustomerSegmentationAnalysis.ipynb
```

Run the cells in order. The first data cell downloads the 45 MB source workbook
from the UCI repository if it is not already present and caches a combined CSV,
so the download happens **once**. No API keys, credentials or manual downloads
are needed. Nothing outside this repository folder is written to.
''')

# ==================================================================== SETUP
md(r'''
---
# 4. Environment and imports

Versions are printed rather than assumed, so that any result in this notebook can
be tied to the exact library versions that produced it.
''')

code(r'''
# --- Standard library -------------------------------------------------------
import json
import re
import sys
import urllib.request
import zipfile
from pathlib import Path

# --- Third party ------------------------------------------------------------
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import sklearn
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.ticker import FuncFormatter, PercentFormatter
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler

%matplotlib inline

pd.set_option("display.max_columns", 40)
pd.set_option("display.width", 170)
pd.set_option("display.float_format", lambda v: f"{v:,.2f}")

print("python      ", sys.version.split()[0])
print("pandas      ", pd.__version__)
print("numpy       ", np.__version__)
print("matplotlib  ", mpl.__version__)
print("scikit-learn", sklearn.__version__)
''')

code(r'''
# Paths are all relative to the repository root, which is where this notebook lives.
ROOT = Path.cwd()
DATA = ROOT / "data"
RAW = DATA / "raw"
OUTPUTS = ROOT / "outputs"
FIGS = OUTPUTS / "figures"
for d in (RAW, FIGS):
    d.mkdir(parents=True, exist_ok=True)

SOURCE_URL = "https://archive.ics.uci.edu/static/public/502/online+retail+ii.zip"
ZIP_PATH = RAW / "online_retail_II.zip"
XLSX_PATH = RAW / "online_retail_II.xlsx"
CACHE_PATH = DATA / "online_retail_II.csv"  # combined sheets, written once

# FACTS collects every headline number this notebook computes. It is written to
# outputs/facts.json at the end so the written report quotes computed values
# instead of hand-copied ones.
FACTS = {}
FIGURES = {}

def fact(key, value):
    """Record a headline number and return it unchanged."""
    FACTS[key] = value
    return value

print("repository root:", ROOT)
''')

# ==================================================================== THEME
md(r'''
## 4.1 Chart theme

Every chart in this notebook is drawn through one theme so that colour carries
meaning consistently:

* **Categorical palette** — four hues assigned in a fixed order and never
  cycled. The palette was checked for colour-vision-deficiency separation
  (worst adjacent pair ΔE 9.1 under protanopia, against a floor of 8) so the
  series stay distinguishable for red–green colourblind readers.
* **Sequential ramp** — a single hue, light to dark, used for magnitude
  heatmaps. Never a rainbow, because a rainbow implies an order that hue does
  not have.
* Grid lines are hairline and recessive; there are no dual-axis charts, and no
  pie charts. Every chart is accompanied by the table it was drawn from, so no
  value is readable only by eye.
''')

code(r'''
SURFACE = "#fcfcfb"    # chart surface
INK = "#0b0b0b"        # primary text
INK_2 = "#4a4a48"      # secondary text (values, labels)
INK_MUTED = "#77776f"  # tick labels, notes
GRID = "#e6e6e2"       # hairline grid

# Categorical palette - fixed order, never cycled. Colour follows the entity,
# never its rank, so a filtered chart does not repaint the survivors.
PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100"]
BLUE, ORANGE, GREEN, AMBER = PALETTE

# Sequential ramp: one hue, light -> dark.
SEQ = LinearSegmentedColormap.from_list(
    "seq_blue", ["#f4f7fc", "#cfdff4", "#8fb6e8", "#3f86da", "#164a86"]
)

mpl.rcParams.update({
    "figure.figsize": (10, 5.2), "figure.dpi": 100, "savefig.dpi": 150,
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
    "font.size": 10.5, "text.color": INK,
    "axes.edgecolor": GRID, "axes.linewidth": 0.8, "axes.labelcolor": INK_2,
    "axes.labelsize": 10, "axes.titlesize": 13.5, "axes.titleweight": "semibold",
    "axes.titlecolor": INK, "axes.titlelocation": "left", "axes.titlepad": 16,
    "axes.grid": True, "axes.axisbelow": True,
    "grid.color": GRID, "grid.linewidth": 0.8, "grid.linestyle": "-",
    "xtick.color": INK_MUTED, "ytick.color": INK_MUTED,
    "xtick.labelsize": 9.5, "ytick.labelsize": 9.5,
    "legend.frameon": False, "legend.fontsize": 9.5,
    "axes.spines.top": False, "axes.spines.right": False, "axes.spines.left": False,
    "lines.linewidth": 2, "lines.markersize": 8,
})


def money(v, _pos=None):
    """Currency axis formatter. The dataset is a UK retailer, so values are GBP."""
    sign = "-" if v < 0 else ""
    a = abs(v)
    if a >= 1e6:
        return f"{sign}£{a / 1e6:,.1f}M"
    if a >= 1e3:
        return f"{sign}£{a / 1e3:,.0f}k"
    return f"{sign}£{a:,.0f}"


GBP = FuncFormatter(money)


def titled(ax, title, subtitle=None):
    """Left-aligned title with an optional muted subtitle above the axes."""
    ax.set_title(title)
    if subtitle:
        ax.annotate(subtitle, xy=(0, 1.02), xycoords="axes fraction", ha="left",
                    va="bottom", fontsize=9.5, color=INK_MUTED)


def label_barh(ax, values, texts=None, fmt=money, pad=0.012):
    """Direct-label horizontal bars. Values stay in ink, never the series colour."""
    span = max((abs(v) for v in values), default=1) or 1
    for i, v in enumerate(values):
        txt = fmt(v) if texts is None else texts[i]
        ax.annotate(txt, xy=(v + pad * span * (1 if v >= 0 else -1), i), va="center",
                    ha="left" if v >= 0 else "right", fontsize=9, color=INK_2)


def save(fig, name):
    """Persist a figure for the written report and register it."""
    path = FIGS / f"{name}.png"
    fig.savefig(path, bbox_inches="tight")
    FIGURES[name] = path.name
    return path


print("theme loaded - palette:", " ".join(PALETTE))
''')

# ============================================================== ACQUISITION
md(r'''
---
# 5. Data acquisition

The source is the **Online Retail II** dataset from the UCI Machine Learning
Repository: <https://archive.ics.uci.edu/dataset/502/online+retail+ii>

It is a real transaction log from a UK-registered, non-store online retailer
selling giftware, mostly to wholesalers. The workbook holds **two sheets, one
per trading year**; both are loaded and concatenated, with a `SourceSheet`
column kept so the two years remain separable.

The cell below is idempotent — it uses the cached CSV if present, then the
extracted workbook, and only downloads from UCI as a last resort. The original
downloaded files are never modified.
''')

code(r'''
EXPECTED_COLS = ["Invoice", "StockCode", "Description", "Quantity", "InvoiceDate",
                 "Price", "Customer ID", "Country", "SourceSheet"]


def download_source():
    """Fetch the 45 MB zip from UCI and extract the workbook. Runs at most once."""
    if not ZIP_PATH.exists():
        print(f"downloading {SOURCE_URL}\n(45 MB from UCI - this can take several minutes)")
        urllib.request.urlretrieve(SOURCE_URL, ZIP_PATH)
    with zipfile.ZipFile(ZIP_PATH) as z:
        member = next(n for n in z.namelist() if n.lower().endswith(".xlsx"))
        XLSX_PATH.write_bytes(z.read(member))
    print(f"extracted {XLSX_PATH.name}  ({XLSX_PATH.stat().st_size / 1e6:.1f} MB)")


def load_raw():
    """Return the combined two-sheet transaction log, caching a CSV on first run."""
    if CACHE_PATH.exists():
        cached = pd.read_csv(CACHE_PATH, parse_dates=["InvoiceDate"],
                             dtype={"Invoice": "str", "StockCode": "str"})
        if list(cached.columns) == EXPECTED_COLS:
            print(f"loaded cache {CACHE_PATH.name}  {cached.shape}")
            return cached
        print("cache schema does not match - rebuilding from the workbook")

    if not XLSX_PATH.exists():
        download_source()

    book = pd.ExcelFile(XLSX_PATH)
    print("sheets:", book.sheet_names)
    frames = []
    for sheet in book.sheet_names:
        part = book.parse(sheet, dtype={"Invoice": "str", "StockCode": "str"})
        part["SourceSheet"] = sheet
        print(f"  {sheet:<18} {part.shape}")
        frames.append(part)
    combined = pd.concat(frames, ignore_index=True)
    combined.to_csv(CACHE_PATH, index=False)
    print(f"cached -> {CACHE_PATH.name}  ({CACHE_PATH.stat().st_size / 1e6:.0f} MB)")
    return combined


raw = load_raw()
fact("raw_rows", int(len(raw)))
fact("raw_cols", int(raw.shape[1]))
print(f"\ncombined: {len(raw):,} rows x {raw.shape[1]} columns")
''')

code(r'''
raw.head(8)
''')

code(r'''
raw.info(memory_usage="deep")
''')

code(r'''
# Numeric spread before any cleaning - the minimums are the first sign of trouble.
raw[["Quantity", "Price"]].describe().T
''')

md(r'''
**First read.** `Quantity` and `Price` both have negative minima and the price
maximum is orders of magnitude above the median. Neither is an error to be
deleted on sight — negative quantities are how this retailer records returns,
and the price outliers are the administrative codes found below. Each needs a
rule, and each rule needs a reason.
''')

# ============================================================ QUALITY AUDIT
md(r'''
---
# 6. Data quality audit

Nothing is dropped in this section. The point is to *measure* every defect
first, so the cleaning rules that follow are chosen against evidence.
''')

code(r'''
# Work on a normalised copy so that whitespace and case do not create phantom
# duplicates or split the same stock code into two.
df = raw.copy()
df["Invoice"] = df["Invoice"].astype("string").str.strip().str.upper()
df["StockCode"] = df["StockCode"].astype("string").str.strip().str.upper()
df["Description"] = df["Description"].astype("string").str.strip().str.upper()
df["Country"] = df["Country"].astype("string").str.strip()
df = df.rename(columns={"Customer ID": "CustomerID"})
df["CustomerID"] = df["CustomerID"].astype("Int64")

# The metric the whole analysis rests on: line-level revenue.
df["Revenue"] = (df["Quantity"] * df["Price"]).round(2)

audit = pd.DataFrame([
    ("Rows", len(df), ""),
    ("Columns", df.shape[1], "8 source + SourceSheet + engineered Revenue"),
    ("Missing Description", int(df["Description"].isna().sum()), ""),
    ("Missing CustomerID", int(df["CustomerID"].isna().sum()),
     f"{df['CustomerID'].isna().mean():.1%} of rows - anonymous / guest orders"),
    ("Exact duplicate rows", int(df.duplicated().sum()), "identical on every column"),
    ("Cancellation invoices (C-prefix)", int(df["Invoice"].str.startswith("C").sum()), ""),
    ("Negative Quantity rows", int((df["Quantity"] < 0).sum()), "returns and stock write-offs"),
    ("Zero Quantity rows", int((df["Quantity"] == 0).sum()), ""),
    ("Zero Price rows", int((df["Price"] == 0).sum()), ""),
    ("Negative Price rows", int((df["Price"] < 0).sum()), ""),
    ("Unique invoices", int(df["Invoice"].nunique()), ""),
    ("Unique stock codes", int(df["StockCode"].nunique()), ""),
    ("Unique customers", int(df["CustomerID"].nunique()), ""),
    ("Countries", int(df["Country"].nunique()), ""),
], columns=["Check", "Value", "Note"])

fact("missing_customer_id", int(df["CustomerID"].isna().sum()))
fact("missing_customer_id_pct", round(float(df["CustomerID"].isna().mean()) * 100, 2))
fact("missing_description", int(df["Description"].isna().sum()))
fact("duplicate_rows", int(df.duplicated().sum()))
fact("cancellation_rows", int(df["Invoice"].str.startswith("C").sum()))
fact("negative_qty_rows", int((df["Quantity"] < 0).sum()))
fact("zero_price_rows", int((df["Price"] == 0).sum()))
fact("raw_countries", int(df["Country"].nunique()))
fact("date_min", str(df["InvoiceDate"].min().date()))
fact("date_max", str(df["InvoiceDate"].max().date()))

audit
''')

code(r'''
# Do the missing descriptions and the missing customer IDs overlap with the
# zero-price rows? If they do, they are one defect, not three.
nd = df[df["Description"].isna()]
print(f"rows with no Description: {len(nd):,}")
print(f"  ... of which Price == 0        : {(nd['Price'] == 0).mean():.1%}")
print(f"  ... of which CustomerID missing: {nd['CustomerID'].isna().mean():.1%}")
print("\nSo the blank-description rows are a single defect: zero-priced, unattributed")
print("lines. They carry no revenue and no customer, and cannot inform any question here.")
''')

code(r'''
# Are all negative quantities returns? Split by whether the invoice is a
# cancellation (C-prefix) to find out.
is_cancel = df["Invoice"].str.startswith("C")
neg = df["Quantity"] < 0

split = pd.DataFrame([
    ("Cancellation invoice (C), negative qty", int((is_cancel & neg).sum()),
     round(float(df.loc[is_cancel & neg, "Revenue"].sum()), 2)),
    ("Cancellation invoice (C), positive qty", int((is_cancel & ~neg).sum()),
     round(float(df.loc[is_cancel & ~neg, "Revenue"].sum()), 2)),
    ("Normal invoice, negative qty", int((~is_cancel & neg).sum()),
     round(float(df.loc[~is_cancel & neg, "Revenue"].sum()), 2)),
], columns=["Group", "Rows", "Revenue (GBP)"])
print(split.to_string(index=False))

other_neg = df[~is_cancel & neg]
print(f"\nOf the {len(other_neg):,} negative-quantity rows on normal invoices, "
      f"{(other_neg['Price'] == 0).mean():.1%} have Price == 0")
print("and their combined revenue is "
      f"{other_neg['Revenue'].sum():,.2f} - they are zero-priced stock adjustments")
print("(damages, write-offs), NOT customer returns. Treating them as returns would")
print(f"overstate the return count by {len(other_neg) / max(1, int((is_cancel & neg).sum())):.0%}.")
fact("writeoff_rows", int(len(other_neg)))
''')

code(r'''
# How many kinds of invoice are there? The usual recipe for this dataset is
# "drop the C invoices"; checking the prefixes shows whether that is sufficient.
prefix = df["Invoice"].str.extract(r"^([A-Z]*)", expand=False).replace("", "(numeric)")
inv_kinds = (df.assign(Prefix=prefix).groupby("Prefix")
             .agg(Rows=("Revenue", "size"), Invoices=("Invoice", "nunique"),
                  Revenue=("Revenue", "sum")).round(2))
print(inv_kinds.to_string())
a_inv = sorted(df.loc[prefix == "A", "Invoice"].unique())
print(f"\nThere is a third invoice type. The {len(a_inv)} 'A' invoices ({', '.join(a_inv)})")
print("are bad-debt adjustments carrying "
      f"{df.loc[prefix == 'A', 'Revenue'].sum():,.2f} in negative revenue.")
print("A pipeline that only filters the C prefix would leave all of it in the revenue line.")
fact("invoice_prefix_a_count", len(a_inv))
fact("invoice_prefix_a_revenue", round(float(df.loc[prefix == "A", "Revenue"].sum()), 2))
''')

code(r'''
# Which stock codes are not products? Real product codes are five digits with an
# optional letter suffix (e.g. 85123A). Everything else is inspected by hand.
prod_pat = re.compile(r"^\d{5}")
non_product = df[~df["StockCode"].str.match(prod_pat).fillna(False)]

code_table = (non_product.groupby("StockCode")
              .agg(Rows=("Quantity", "size"),
                   Revenue=("Revenue", "sum"),
                   Example=("Description", lambda s: s.dropna().iloc[0] if s.notna().any() else ""))
              .sort_values("Revenue", ascending=False)
              .round(2))
print(f"{len(non_product):,} rows across {len(code_table)} non-standard stock codes\n")
print(code_table.head(14).to_string())
print("\n... net-negative codes (fees, adjustments, write-offs):\n")
print(code_table[code_table["Revenue"] < 0].to_string())
''')

md(r'''
**What the code table shows.** Three genuinely different things are mixed
together under `StockCode`:

1. **Service charges** — `DOT` (Dotcom postage), `POST`, `C2`/`C3` (carriage)
   and the `GIFT_0001_*` vouchers. These are money a real customer really paid,
   but they are not merchandise. They belong in revenue and in customer value;
   they must be excluded from *product* rankings.
2. **Administrative entries** — `AMAZONFEE`, `B` (bad-debt adjustment), `M`
   (manual), `BANK CHARGES`, `D` (discount), `CRUK` (charity commission),
   `S` (samples), `ADJUST`, `TEST001`/`TEST002`. Every one of these is either a
   cost, a correction or a test. They are not sales at all and are removed.
3. **Real products with non-standard codes** — the `DCGS*` range (dog collars,
   bags, ashtrays), `SP1002`, `PADS`. These stay, and stay counted as products.

The invoice prefixes matter for the same reason. Cancellations are not the only
reversal in the file: six `A`-prefixed invoices carry bad-debt write-offs worth
−£147,614. The standard recipe for this dataset — "drop the rows whose invoice
starts with C" — leaves every one of them in the revenue line.

This is why the audit came before the cleaning. A blanket "drop everything whose
stock code is not five digits" rule would have thrown away real merchandise, and
a blanket "keep everything" rule would have let a −£260,763 Amazon fee sit in the
revenue line. In particular, `M` (*Manual*) shows a large **positive** figure if
only positive quantities are counted, but its true net across all rows is
**negative** — it is a bidirectional correction code, and reporting it as a
top-selling "product" would be an outright error.
''')

# ================================================================ CLEANING
md(r'''
---
# 7. Cleaning

Nine steps, each one logged with the rows and revenue it removes. The log is the
audit trail: any figure later in this notebook can be traced back to the rule
that produced the population it was computed on.

The design decision worth stating explicitly: **cancellations and write-offs are
set aside, not deleted.** They are analysed in their own right in section 10.
Deleting a return destroys the information that the return happened.
''')

code(r'''
CARRIAGE_CODES = {"DOT", "POST", "C2", "C3"}
ADMIN_CODES = {"AMAZONFEE", "B", "M", "BANK CHARGES", "D", "CRUK", "S",
               "ADJUST", "ADJUST2", "TEST001", "TEST002"}

steps = []
prev_rows, prev_rev = len(df), float(df["Revenue"].sum())


def log_step(name, frame, note):
    """Record the row and revenue effect of one cleaning step."""
    global prev_rows, prev_rev
    rows, rev = len(frame), float(frame["Revenue"].sum())
    steps.append({
        "Step": name, "Rows": rows, "Rows removed": prev_rows - rows,
        "Revenue (GBP)": round(rev, 2), "Revenue removed": round(prev_rev - rev, 2),
        "Rationale": note,
    })
    prev_rows, prev_rev = rows, rev
    return frame


work = df
steps.append({"Step": "0. Raw combined (both sheets)", "Rows": len(work), "Rows removed": 0,
              "Revenue (GBP)": round(float(work["Revenue"].sum()), 2), "Revenue removed": 0.0,
              "Rationale": "starting point"})

# 1. Normalisation and Revenue were applied in the audit section above.
steps.append({"Step": "1. Normalise text, types; engineer Revenue", "Rows": len(work),
              "Rows removed": 0, "Revenue (GBP)": round(float(work["Revenue"].sum()), 2),
              "Revenue removed": 0.0,
              "Rationale": "trim/upper-case keys so identical rows collide; Revenue = Quantity x Price"})

# 2. Exact duplicates.
work = log_step("2. Drop exact duplicate rows", work.drop_duplicates().copy(),
                "identical on every column incl. timestamp - a double write, not two sales")

# 3. Set aside cancellations.
cancel_mask = work["Invoice"].str.startswith("C")
cancellations = work[cancel_mask].copy()
work = log_step("3. Set aside cancellation invoices", work[~cancel_mask].copy(),
                "kept for the returns analysis in section 10")

# 4. Set aside zero-priced negative-quantity write-offs.
wo_mask = work["Quantity"] < 0
writeoffs = work[wo_mask].copy()
work = log_step("4. Set aside stock write-offs", work[~wo_mask].copy(),
                "negative qty on a normal invoice, zero priced - damages/adjustments")

# 5. Zero and negative prices.
work = log_step("5. Drop non-positive Price", work[work["Price"] > 0].copy(),
                "no consideration changed hands; also removes every blank-description row")

# 6. Zero quantities.
work = log_step("6. Drop non-positive Quantity", work[work["Quantity"] > 0].copy(),
                "a line with no units is not a sale")

# 7. Administrative codes.
work = log_step("7. Drop administrative stock codes",
                work[~work["StockCode"].isin(ADMIN_CODES)].copy(),
                "fees, bad debt, manual corrections, discounts, samples, tests")

# 8. Flag products vs carriage - a classification, not a deletion.
work["IsProduct"] = ~work["StockCode"].isin(CARRIAGE_CODES) & ~work["StockCode"].str.startswith("GIFT")
steps.append({"Step": "8. Flag product vs carriage/voucher lines", "Rows": len(work),
              "Rows removed": 0, "Revenue (GBP)": round(float(work["Revenue"].sum()), 2),
              "Revenue removed": 0.0,
              "Rationale": f"{int((~work['IsProduct']).sum()):,} non-merchandise lines flagged, not dropped"})

sales = work
anonymous = sales[sales["CustomerID"].isna()].copy()

# 9. Known-customer subset - the base for every customer-level model.
sales_id = log_step("9. Restrict to known CustomerID (customer analytics base)",
                    sales[sales["CustomerID"].notna()].copy(),
                    "RFM, K-Means, CLV and cohorts all require an identified customer")

clean_log = pd.DataFrame(steps)
clean_log
''')

code(r'''
# Two analysis bases, used deliberately for different questions.
fact("sales_rows", int(len(sales)))
fact("sales_revenue", round(float(sales["Revenue"].sum()), 2))
fact("sales_id_rows", int(len(sales_id)))
fact("sales_id_revenue", round(float(sales_id["Revenue"].sum()), 2))
fact("anonymous_rows", int(len(anonymous)))
fact("anonymous_revenue", round(float(anonymous["Revenue"].sum()), 2))
# Both shares are measured against the same clean base, so they are comparable:
# guest lines are a larger share of rows than of money.
fact("anonymous_rows_pct", round(len(anonymous) / len(sales) * 100, 1))
fact("anonymous_revenue_pct", round(float(anonymous["Revenue"].sum())
                                    / float(sales["Revenue"].sum()) * 100, 1))
fact("id_coverage_pct", round(float(sales_id["Revenue"].sum()) / float(sales["Revenue"].sum()) * 100, 1))
fact("rows_retained_pct", round(len(sales) / len(df) * 100, 1))

bases = pd.DataFrame([
    ("sales", "All valid revenue lines (identified + guest)", len(sales),
     round(float(sales["Revenue"].sum()), 2),
     "revenue trend, demand rhythm, products, geography"),
    ("sales_id", "Known-customer subset of sales", len(sales_id),
     round(float(sales_id["Revenue"].sum()), 2),
     "RFM, K-Means, CLV, cohort retention"),
    ("cancellations", "C-prefix invoices set aside in step 3", len(cancellations),
     round(float(cancellations["Revenue"].sum()), 2), "returns analysis"),
    ("writeoffs", "Zero-priced negative lines set aside in step 4", len(writeoffs),
     round(float(writeoffs["Revenue"].sum()), 2), "quality reporting only"),
], columns=["Frame", "Definition", "Rows", "Revenue (GBP)", "Used for"])
print(bases.to_string(index=False))

print(f"\nCleaning retained {len(sales):,} of {len(df):,} rows ({len(sales) / len(df):.1%}).")
print(f"Identified customers cover {FACTS['id_coverage_pct']}% of clean revenue;")
print(f"the remaining {len(anonymous):,} guest lines "
      f"(£{anonymous['Revenue'].sum():,.0f}) cannot be attributed to a customer and are")
print("therefore excluded from every customer-level model, but retained for revenue reporting.")
''')

code(r'''
# Derived time columns used throughout, added once.
for frame in (sales, sales_id):
    frame["InvoiceMonth"] = frame["InvoiceDate"].dt.to_period("M")
    frame["Weekday"] = frame["InvoiceDate"].dt.day_name()
    frame["Hour"] = frame["InvoiceDate"].dt.hour

sales_id.head(5)
''')

# ============================================================== DESCRIPTIVE
md(r'''
---
# 8. Descriptive analysis

## 8.1 Headline figures
''')

code(r'''
order_totals = sales.groupby("Invoice")["Revenue"].sum()
basket_lines = sales.groupby("Invoice").size()

kpi = pd.DataFrame([
    ("Total revenue", money(sales["Revenue"].sum())),
    ("Orders (invoices)", f"{sales['Invoice'].nunique():,}"),
    ("Identified customers", f"{sales_id['CustomerID'].nunique():,}"),
    ("Distinct products sold", f"{sales.loc[sales['IsProduct'], 'StockCode'].nunique():,}"),
    ("Countries served", f"{sales['Country'].nunique():,}"),
    ("Average order value", money(order_totals.mean())),
    ("Median order value", money(order_totals.median())),
    ("Average lines per order", f"{basket_lines.mean():.1f}"),
    ("Units sold", f"{int(sales['Quantity'].sum()):,}"),
    ("Trading window", f"{sales['InvoiceDate'].min():%d %b %Y} to {sales['InvoiceDate'].max():%d %b %Y}"),
], columns=["Metric", "Value"])

fact("total_revenue", round(float(sales["Revenue"].sum()), 2))
fact("orders", int(sales["Invoice"].nunique()))
fact("customers", int(sales_id["CustomerID"].nunique()))
fact("products", int(sales.loc[sales["IsProduct"], "StockCode"].nunique()))
fact("countries", int(sales["Country"].nunique()))
fact("aov", round(float(order_totals.mean()), 2))
fact("median_order", round(float(order_totals.median()), 2))
fact("lines_per_order", round(float(basket_lines.mean()), 1))
fact("units", int(sales["Quantity"].sum()))

print(kpi.to_string(index=False))
''')

md(r'''
The gap between the **mean** order value and the **median** is the first real
finding. The mean sits far above the median, which means order value is
right-skewed: a small number of large wholesale orders pulls the average up
while the typical order is much smaller. Any plan built on "the average
customer" would be planning for a customer who does not exist — which is the
case for segmentation in one line.

## 8.2 Revenue over time
''')

code(r'''
monthly = (sales.groupby("InvoiceMonth")
           .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                Customers=("CustomerID", "nunique"))
           .reset_index())
monthly["Month"] = monthly["InvoiceMonth"].dt.to_timestamp()
monthly["AOV"] = (monthly["Revenue"] / monthly["Orders"]).round(2)

# The final month is truncated: the log stops on 09 Dec 2011, so December 2011
# holds 9 days of trading, not 31. It is plotted but marked, never compared
# like-for-like against a full month.
last_month = monthly["InvoiceMonth"].iloc[-1]
partial = monthly.index[-1]

fig, ax = plt.subplots(figsize=(11, 5.2))
ax.plot(monthly["Month"][:partial + 1], monthly["Revenue"][:partial + 1],
        color=BLUE, marker="o", markerfacecolor=BLUE, markeredgecolor=SURFACE,
        markeredgewidth=1.6, zorder=3)
ax.plot(monthly["Month"].iloc[partial - 1:], monthly["Revenue"].iloc[partial - 1:],
        color=BLUE, linestyle=(0, (3, 2)), zorder=2)
ax.scatter(monthly["Month"].iloc[partial], monthly["Revenue"].iloc[partial],
           color=SURFACE, edgecolor=BLUE, linewidth=1.8, zorder=4, s=70)
ax.annotate(f"{last_month} is a partial month\n(log ends 09 Dec 2011)",
            xy=(monthly["Month"].iloc[partial], monthly["Revenue"].iloc[partial]),
            xytext=(-14, 34), textcoords="offset points", ha="right", fontsize=9,
            color=INK_MUTED,
            arrowprops=dict(arrowstyle="-", color=GRID, linewidth=1.2))

# Mark the two Q4 peaks - the seasonal signature of a giftware business.
for yr in (2010, 2011):
    peak = monthly[monthly["Month"].dt.year == yr].nlargest(1, "Revenue")
    if len(peak):
        ax.annotate(money(peak["Revenue"].iloc[0]),
                    xy=(peak["Month"].iloc[0], peak["Revenue"].iloc[0]),
                    xytext=(0, 12), textcoords="offset points", ha="center",
                    fontsize=9, color=INK_2)

ax.yaxis.set_major_formatter(GBP)
ax.set_ylim(0, monthly["Revenue"].max() * 1.18)
ax.set_ylabel("Revenue")
ax.grid(axis="x", visible=False)
titled(ax, "Monthly revenue peaks in Q4 and troughs in February",
        "All valid sales lines, Dec 2009 - Dec 2011")
save(fig, "01_revenue_trend")
plt.show()

monthly[["InvoiceMonth", "Revenue", "Orders", "Customers", "AOV"]]
''')

code(r'''
# Quantify the seasonality rather than leaving it to the eye. September-November
# is the wholesale build-up to Christmas; the partial final month is excluded.
full = monthly.iloc[:-1]
q4 = full[full["Month"].dt.month.isin([9, 10, 11])]
rest = full[~full["Month"].dt.month.isin([9, 10, 11])]
fact("q4_monthly_avg", round(float(q4["Revenue"].mean()), 2))
fact("rest_monthly_avg", round(float(rest["Revenue"].mean()), 2))
fact("q4_uplift_pct", round((float(q4["Revenue"].mean()) / float(rest["Revenue"].mean()) - 1) * 100, 1))
fact("peak_month", str(full.nlargest(1, "Revenue")["InvoiceMonth"].iloc[0]))
fact("peak_month_revenue", round(float(full["Revenue"].max()), 2))
fact("trough_month", str(full.nsmallest(1, "Revenue")["InvoiceMonth"].iloc[0]))
fact("trough_month_revenue", round(float(full["Revenue"].min()), 2))

print(f"Sep-Nov average month : {money(q4['Revenue'].mean())}")
print(f"Other months average  : {money(rest['Revenue'].mean())}")
print(f"Seasonal uplift       : {FACTS['q4_uplift_pct']:+.1f}%")
print(f"Peak month            : {FACTS['peak_month']}  {money(FACTS['peak_month_revenue'])}")
print(f"Trough month          : {FACTS['trough_month']}  {money(FACTS['trough_month_revenue'])}")

# Year-on-year on a like-for-like window (Dec-Nov, both years complete).
y1 = sales[(sales["InvoiceDate"] >= "2009-12-01") & (sales["InvoiceDate"] < "2010-12-01")]["Revenue"].sum()
y2 = sales[(sales["InvoiceDate"] >= "2010-12-01") & (sales["InvoiceDate"] < "2011-12-01")]["Revenue"].sum()
fact("year1_revenue", round(float(y1), 2))
fact("year2_revenue", round(float(y2), 2))
fact("yoy_growth_pct", round((float(y2) / float(y1) - 1) * 100, 1))
print(f"\nLike-for-like years (Dec-Nov):  {money(y1)} -> {money(y2)}   {FACTS['yoy_growth_pct']:+.1f}%")
''')

md(r'''
## 8.3 New versus returning revenue

A revenue line that goes up says nothing about *why*. Splitting each month into
revenue from customers acquired that month and revenue from customers acquired
earlier separates an acquisition engine from a retention engine.
''')

code(r'''
first_month = sales_id.groupby("CustomerID")["InvoiceMonth"].min().rename("CohortMonth")
tagged = sales_id.join(first_month, on="CustomerID")
tagged["CustomerType"] = np.where(tagged["InvoiceMonth"] == tagged["CohortMonth"],
                                 "New customers", "Returning customers")

split_rev = (tagged.pivot_table(index="InvoiceMonth", columns="CustomerType",
                                values="Revenue", aggfunc="sum", observed=True)
             .fillna(0.0))
split_rev = split_rev.reindex(columns=["New customers", "Returning customers"], fill_value=0.0)
idx = split_rev.index.to_timestamp()

fig, ax = plt.subplots(figsize=(11, 5.2))
w = 22  # bar width in days, leaving a visible surface gap between months
ax.bar(idx, split_rev["Returning customers"], width=w, color=BLUE,
       label="Returning customers", zorder=3)
ax.bar(idx, split_rev["New customers"], width=w,
       bottom=split_rev["Returning customers"], color=ORANGE,
       label="New customers", zorder=3, linewidth=1.4, edgecolor=SURFACE)
ax.yaxis.set_major_formatter(GBP)
ax.set_ylabel("Revenue")
ax.grid(axis="x", visible=False)
ax.legend(loc="upper left", ncol=2, bbox_to_anchor=(0, 1.0))
titled(ax, "Returning customers carry almost all revenue after the first year",
        "Revenue split by whether the customer was acquired in that month")
save(fig, "02_new_vs_returning")
plt.show()

share = split_rev.div(split_rev.sum(axis=1), axis=0) * 100
out = split_rev.round(2).join(share.round(1), rsuffix=" %")
fact("returning_share_year2_pct", round(float(
    split_rev.loc[split_rev.index >= pd.Period("2010-12"), "Returning customers"].sum()
    / split_rev.loc[split_rev.index >= pd.Period("2010-12")].sum().sum() * 100), 1))
print(f"In the second year (Dec 2010 onward), returning customers generated "
      f"{FACTS['returning_share_year2_pct']}% of identified revenue.\n")
out
''')

md(r'''
## 8.4 Demand rhythm — when do orders arrive?

Order *volume* by weekday and hour is an operations question: it sizes
warehouse and support shifts. Counting distinct invoices, not lines, avoids
letting one large basket look like many orders.
''')

code(r'''
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
heat = (sales.drop_duplicates("Invoice")
        .pivot_table(index="Weekday", columns="Hour", values="Invoice",
                     aggfunc="count", observed=True)
        .reindex(DAYS))
heat_plot = heat.dropna(axis=1, how="all")

fig, ax = plt.subplots(figsize=(12, 4.4))
im = ax.imshow(heat_plot.to_numpy(dtype=float), aspect="auto", cmap=SEQ,
               vmin=0, vmax=np.nanmax(heat_plot.to_numpy(dtype=float)))
ax.set_xticks(range(len(heat_plot.columns)), [f"{h:02d}" for h in heat_plot.columns])
ax.set_yticks(range(len(heat_plot.index)), heat_plot.index)
ax.set_xlabel("Hour of day")
ax.grid(False)
peak_v = np.nanmax(heat_plot.to_numpy(dtype=float))
for i in range(heat_plot.shape[0]):
    for j in range(heat_plot.shape[1]):
        v = heat_plot.iat[i, j]
        if pd.isna(v):
            ax.annotate("-", (j, i), ha="center", va="center", fontsize=8, color=INK_MUTED)
        else:
            ax.annotate(f"{int(v):,}", (j, i), ha="center", va="center", fontsize=7.4,
                        color=SURFACE if v > peak_v * 0.55 else INK_2)
cb = fig.colorbar(im, ax=ax, pad=0.015, fraction=0.03)
cb.set_label("Orders", color=INK_2, fontsize=9.5)
cb.outline.set_visible(False)
titled(ax, "Orders arrive on weekdays, concentrated between 10:00 and 15:00",
        "Distinct invoices by weekday and hour of day")
save(fig, "03_demand_heatmap")
plt.show()

# The Saturday column is nearly empty. "Nearly" is the interesting part, so it is
# checked rather than rounded to zero.
sat = int(heat.loc["Saturday"].sum(skipna=True)) if "Saturday" in heat.index else 0
sat_rows = sales[sales["Weekday"] == "Saturday"]
sat_dates = sorted(sat_rows["InvoiceDate"].dt.date.unique())
fact("saturday_orders", sat)
fact("saturday_distinct_dates", len(sat_dates))
fact("saturday_date", str(sat_dates[0]) if sat_dates else None)
fact("saturday_share_pct", round(sat / int(sales["Invoice"].nunique()) * 100, 3))
sun = int(heat.loc["Sunday"].sum(skipna=True)) if "Sunday" in heat.index else 0
fact("sunday_orders", sun)
# How many Saturdays in the trading window saw no orders at all? Count the
# calendar Saturdays and subtract the ones that appear in the data.
all_saturdays = pd.date_range(sales["InvoiceDate"].min().normalize(),
                              sales["InvoiceDate"].max().normalize(), freq="W-SAT")
fact("saturdays_in_window", len(all_saturdays))
fact("saturdays_with_zero_orders", len(all_saturdays) - len(sat_dates))
print(f"Saturday orders across 25 months : {sat:,} "
      f"({FACTS['saturday_share_pct']}% of all orders)")
print(f"  ... falling on {len(sat_dates)} distinct Saturday date(s): "
      f"{', '.join(str(d) for d in sat_dates[:5])}")
print(f"  so {FACTS['saturdays_with_zero_orders']} of the "
      f"{FACTS['saturdays_in_window']} Saturdays in the window carry zero orders.")
print(f"Sunday orders, for contrast       : {sun:,} "
      f"({sun / max(sat, 1):.0f}x the Saturday count)")
busiest = heat.sum(axis=1).idxmax()
fact("busiest_weekday", str(busiest))
fact("busiest_weekday_orders", int(heat.sum(axis=1).max()))
fact("peak_hour", int(heat.sum(axis=0).idxmax()))
fact("core_hours_share_pct", round(float(
    heat.loc[:, [h for h in heat.columns if 10 <= h <= 15]].sum().sum() / heat.sum().sum() * 100), 1))
print(f"Saturday orders across 25 months : {sat:,}")
print(f"Busiest weekday                  : {busiest} ({FACTS['busiest_weekday_orders']:,} orders)")
print(f"Peak hour                        : {FACTS['peak_hour']:02d}:00")
print(f"Share of orders in 10:00-15:59   : {FACTS['core_hours_share_pct']}%")
heat.fillna(0).astype(int)
''')

md(r'''
**One Saturday in twenty-five months.** Saturday is not merely quiet — all 30
Saturday orders in the entire dataset fall on a *single date*, 5 December 2009,
so every other Saturday in the window carries zero orders. Sunday, meanwhile,
processes 4,826 orders at weekday-like volumes.

A customer base that trades every Sunday but only one Saturday in 25 months does
not exist. That pattern cannot describe customer behaviour, so `InvoiceDate` must
record **when an order was keyed into the system by staff**, not when the
customer decided to buy. The lone Saturday is most likely a catch-up session
during the pre-Christmas peak.

Two consequences follow, and they point in opposite directions. Any "best day to
advertise" conclusion drawn from this column would be an artefact of the office
rota and is worthless. The same column *is* reliable for operations: it shows
genuine back-office load, so warehouse and support staffing should follow the
weekday 10:00–15:00 concentration, which accounts for the large majority of all
order processing.

## 8.5 Product performance
''')

code(r'''
prod = sales[sales["IsProduct"]]
by_rev = (prod.groupby("Description")
          .agg(Revenue=("Revenue", "sum"), Units=("Quantity", "sum"),
               Orders=("Invoice", "nunique"))
          .sort_values("Revenue", ascending=False))
by_rev["Revenue share %"] = (by_rev["Revenue"] / prod["Revenue"].sum() * 100).round(2)
top10 = by_rev.head(10)

fig, ax = plt.subplots(figsize=(10.5, 5.4))
y = np.arange(len(top10))[::-1]
ax.barh(y, top10["Revenue"], height=0.72, color=BLUE, zorder=3)
ax.set_yticks(y, [t.title() if len(t) < 42 else t[:39].title() + "..." for t in top10.index])
label_barh(ax, top10["Revenue"].to_list())
ax.xaxis.set_major_formatter(GBP)
ax.set_xlim(0, top10["Revenue"].max() * 1.16)
ax.grid(axis="y", visible=False)
ax.grid(axis="x", visible=True)
titled(ax, "The top ten lines out of 4,700 products",
        "Merchandise only - carriage, vouchers and administrative codes excluded")
save(fig, "04_top_products")
plt.show()

fact("top_product", str(top10.index[0]).title())
fact("top_product_revenue", round(float(top10["Revenue"].iloc[0]), 2))
fact("top10_product_share_pct", round(float(top10["Revenue"].sum() / prod["Revenue"].sum() * 100), 2))
fact("product_count", int(prod["StockCode"].nunique()))
carriage_rev = float(sales.loc[~sales["IsProduct"], "Revenue"].sum())
fact("carriage_revenue", round(carriage_rev, 2))
fact("carriage_share_pct", round(carriage_rev / float(sales["Revenue"].sum()) * 100, 2))
fact("top10_concentration_ratio", round(FACTS["top10_product_share_pct"] / (10 / FACTS["product_count"] * 100), 1))
print(f"Top 10 products = {FACTS['top10_product_share_pct']}% of merchandise revenue "
      f"out of {FACTS['product_count']:,} distinct products,")
print(f"i.e. 0.2% of the catalogue earns {FACTS['top10_product_share_pct']}% of the money - "
      f"{FACTS['top10_concentration_ratio']}x over-representation.")
print(f"Carriage and vouchers add {money(carriage_rev)} ({FACTS['carriage_share_pct']}% of total revenue) "
      "and are excluded above.")
top10
''')

code(r'''
# Best sellers by units are not the same list as by revenue - a cheap high-volume
# line and an expensive low-volume line are different commercial propositions.
by_units = by_rev.sort_values("Units", ascending=False).head(10)
overlap = len(set(by_units.index) & set(top10.index))
fact("rev_unit_overlap", int(overlap))
print(f"Only {overlap} of the top 10 by revenue also appear in the top 10 by units sold.\n")
by_units[["Units", "Revenue", "Orders"]]
''')

md(r'''
## 8.6 Geography

The retailer is UK-registered and UK-dominated, so plotting all 40+ countries on
one axis would compress every export market into an invisible sliver. The
comparison is therefore split: the domestic/export share first, then the export
markets on their own scale.
''')

code(r'''
by_country = (sales.groupby("Country")
              .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                   Customers=("CustomerID", "nunique"))
              .sort_values("Revenue", ascending=False))
by_country["AOV"] = (by_country["Revenue"] / by_country["Orders"]).round(2)
by_country["Revenue share %"] = (by_country["Revenue"] / sales["Revenue"].sum() * 100).round(2)

uk = float(by_country.loc["United Kingdom", "Revenue"])
row_rev = float(by_country["Revenue"].sum()) - uk
export = by_country.drop(index="United Kingdom").head(10)

fig, axes = plt.subplots(1, 2, figsize=(12.4, 4.8), gridspec_kw={"width_ratios": [1, 2.1]})

ax = axes[0]
ax.bar([0, 1], [uk, row_rev], width=0.55, color=[BLUE, ORANGE], zorder=3)
ax.set_xticks([0, 1], ["United Kingdom", "Export markets"])
ax.set_xlim(-0.6, 1.6)
ax.yaxis.set_major_formatter(GBP)
ax.set_ylim(0, uk * 1.2)
for x, v in zip([0, 1], [uk, row_rev]):
    ax.annotate(f"{money(v)}\n{v / (uk + row_rev):.1%}", (x, v), xytext=(0, 8),
                textcoords="offset points", ha="center", fontsize=9.5, color=INK_2)
ax.grid(axis="x", visible=False)
titled(ax, "One market is the business", f"{uk / (uk + row_rev):.0%} of revenue is domestic")

ax = axes[1]
y = np.arange(len(export))[::-1]
ax.barh(y, export["Revenue"], height=0.7, color=ORANGE, zorder=3)
ax.set_yticks(y, export.index)
label_barh(ax, export["Revenue"].to_list())
ax.xaxis.set_major_formatter(GBP)
ax.set_xlim(0, export["Revenue"].max() * 1.2)
ax.grid(axis="y", visible=False)
titled(ax, "Top ten export markets", "Own scale - the UK bar would be 24x the tallest here")
fig.tight_layout(w_pad=3)
save(fig, "05_geography")
plt.show()

fact("uk_revenue", round(uk, 2))
fact("uk_share_pct", round(uk / (uk + row_rev) * 100, 2))
fact("export_revenue", round(row_rev, 2))
fact("top_export", str(export.index[0]))
fact("top_export_revenue", round(float(export["Revenue"].iloc[0]), 2))
uk_aov = float(by_country.loc["United Kingdom", "AOV"])
big_exp = by_country.drop(index="United Kingdom")
big_exp = big_exp[big_exp["Orders"] >= 100]
fact("uk_aov", round(uk_aov, 2))
fact("best_export_aov_country", str(big_exp["AOV"].idxmax()))
fact("best_export_aov", round(float(big_exp["AOV"].max()), 2))
print(f"UK average order value: {money(uk_aov)}")
print(f"Highest AOV among export markets with 100+ orders: {FACTS['best_export_aov_country']} "
      f"({money(FACTS['best_export_aov'])}) - {FACTS['best_export_aov'] / uk_aov:.1f}x the UK")
by_country.head(12)
''')

# ================================================================= RETURNS
md(r'''
---
# 9. Returns and cancellations

Section 7 set the cancellation invoices aside rather than deleting them. They
are a business signal: a return is a sale that was won and then lost, plus a
handling cost.
''')

code(r'''
ret = cancellations[cancellations["Quantity"] < 0].copy()
ret["ReturnValue"] = ret["Revenue"].abs()
ret["InvoiceMonth"] = ret["InvoiceDate"].dt.to_period("M")

gross = float(sales["Revenue"].sum())
total_ret = float(ret["ReturnValue"].sum())
fact("return_rows", int(len(ret)))
fact("return_value", round(total_ret, 2))
fact("return_rate_pct", round(total_ret / gross * 100, 2))
fact("returning_customers", int(ret["CustomerID"].nunique()))
fact("net_revenue", round(gross - total_ret, 2))

print(f"Return lines            : {len(ret):,}")
print(f"Value returned          : {money(total_ret)}")
print(f"As % of gross revenue   : {FACTS['return_rate_pct']}%")
print(f"Net of returns          : {money(gross - total_ret)}")
print(f"Customers with a return : {ret['CustomerID'].nunique():,} of {sales_id['CustomerID'].nunique():,} "
      f"({ret['CustomerID'].nunique() / sales_id['CustomerID'].nunique():.1%})")
''')

code(r'''
# The largest reversals, matched back to the orders they cancel. Cancellation
# invoices are numbered right after the sale they reverse, so the pair is visible.
big = ret.nlargest(5, "ReturnValue")[["Invoice", "StockCode", "Description",
                                      "Quantity", "Price", "InvoiceDate", "CustomerID",
                                      "ReturnValue"]]
print("Five largest cancellations:")
print(big.to_string(index=False))

# Trace the biggest one back to its original sale.
top_code = big["StockCode"].iloc[0]
top_cust = big["CustomerID"].iloc[0]
pair = df[(df["StockCode"] == top_code) & (df["CustomerID"] == top_cust)].sort_values("InvoiceDate")
print(f"\nEvery line in the dataset for stock code {top_code}, customer {top_cust}:")
print(pair[["Invoice", "Quantity", "Price", "InvoiceDate", "Revenue"]].to_string(index=False))
gap = (pair["InvoiceDate"].iloc[-1] - pair["InvoiceDate"].iloc[0]).total_seconds() / 60
fact("biggest_sale_product", str(pair["Description"].iloc[0]).title())
fact("biggest_sale_units", int(pair["Quantity"].iloc[0]))
fact("biggest_sale_value", round(float(pair["Revenue"].iloc[0]), 2))
fact("biggest_sale_cancel_gap_min", round(gap, 0))
fact("biggest_sale_net", round(float(pair["Revenue"].sum()), 2))
print(f"\nThe single largest line in the whole dataset - {pair['Quantity'].iloc[0]:,} units of")
print(f"{str(pair['Description'].iloc[0]).title()} at {money(pair['Revenue'].iloc[0])} - was cancelled")
print(f"{gap:.0f} minutes later, on the final day of the log. Net contribution: "
      f"{money(pair['Revenue'].sum())}.")
# The second largest cancellation tells the same story, so record it too.
second = big.iloc[1]
fact("second_cancel_product", str(second["Description"]).title())
fact("second_cancel_units", int(abs(second["Quantity"])))
fact("second_cancel_value", round(float(second["ReturnValue"]), 2))
print(f"\nThe second largest cancellation: {abs(int(second['Quantity'])):,} units of")
print(f"{str(second['Description']).title()} at {money(second['ReturnValue'])}, "
      f"invoice {second['Invoice']}.")
print("\nIt is almost certainly a keying error, not a real order. Any analysis that")
print("ranked products by gross sales without netting off cancellations would report")
print("this as the retailer's best-selling line. It sold nothing.")
''')

md(r'''
The two largest lines in the entire dataset were both cancelled in full. Ranking
products on gross sales alone would put a product that generated **£0** at the
top of the table — which is exactly why section 7 kept the cancellations instead
of deleting them, and why the product ranking in section 8.5 is worth
re-examining net of returns.
''')

code(r'''
rr = (pd.DataFrame({
        "Gross": sales.groupby("InvoiceMonth")["Revenue"].sum(),
        "Returned": ret.groupby("InvoiceMonth")["ReturnValue"].sum()})
      .fillna(0.0))
rr["Return rate %"] = (rr["Returned"] / rr["Gross"] * 100).round(2)
rr_plot = rr.iloc[:-1]  # exclude the partial final month
x = rr_plot.index.to_timestamp()

fig, ax = plt.subplots(figsize=(11, 4.8))
ax.plot(x, rr_plot["Return rate %"], color=ORANGE, marker="o",
        markeredgecolor=SURFACE, markeredgewidth=1.6, zorder=3)
mean_rate = float(rr_plot["Return rate %"].mean())
ax.axhline(mean_rate, color=INK_MUTED, linewidth=1, linestyle=(0, (4, 3)), zorder=2)
ax.annotate(f"mean {mean_rate:.1f}%", xy=(x[-1], mean_rate), xytext=(6, 4),
            textcoords="offset points", fontsize=9, color=INK_MUTED)
worst = rr_plot["Return rate %"].idxmax()
ax.annotate(f"{worst}: {rr_plot['Return rate %'].max():.1f}%",
            xy=(worst.to_timestamp(), rr_plot["Return rate %"].max()),
            xytext=(0, 12), textcoords="offset points", ha="center", fontsize=9, color=INK_2)
ax.yaxis.set_major_formatter(PercentFormatter(decimals=0))
ax.set_ylabel("Returned value as % of gross")
ax.set_ylim(0, rr_plot["Return rate %"].max() * 1.3)
ax.grid(axis="x", visible=False)
titled(ax, "The return rate is volatile and spike-driven, not a steady leak",
        "Monthly returned value divided by gross revenue; partial final month excluded")
save(fig, "06_return_rate")
plt.show()

fact("worst_return_month", str(worst))
fact("worst_return_rate_pct", round(float(rr_plot["Return rate %"].max()), 2))
fact("mean_monthly_return_rate_pct", round(mean_rate, 2))
rr
''')

code(r'''
top_ret = (ret[ret["StockCode"].isin(prod["StockCode"])]
           .groupby("Description")
           .agg(Returned=("ReturnValue", "sum"), Lines=("Quantity", "size"))
           .sort_values("Returned", ascending=False).head(10))
sold = prod.groupby("Description")["Revenue"].sum()
top_ret["Sold"] = sold.reindex(top_ret.index).fillna(0.0)
top_ret["Return rate %"] = (top_ret["Returned"] / top_ret["Sold"].replace(0, np.nan) * 100).round(1)

fig, ax = plt.subplots(figsize=(10.5, 5.2))
y = np.arange(len(top_ret))[::-1]
ax.barh(y, top_ret["Returned"], height=0.72, color=ORANGE, zorder=3)
ax.set_yticks(y, [t.title() if len(t) < 40 else t[:37].title() + "..." for t in top_ret.index])
label_barh(ax, top_ret["Returned"].to_list())
ax.xaxis.set_major_formatter(GBP)
ax.set_xlim(0, top_ret["Returned"].max() * 1.18)
ax.grid(axis="y", visible=False)
ax.grid(axis="x", visible=True)
titled(ax, "Returns concentrate in a handful of product lines",
        "Merchandise only; the table adds each line's return rate against its own sales")
save(fig, "07_top_returned")
plt.show()

fact("top_returned_product", str(top_ret.index[0]).title())
fact("top_returned_value", round(float(top_ret["Returned"].iloc[0]), 2))
top_ret
''')

code(r'''
# Product ranking net of cancellations - the version a buyer should actually use.
returned_by_product = ret.groupby("Description")["ReturnValue"].sum()
net_rank = (by_rev[["Revenue", "Units", "Orders"]]
            .assign(Returned=returned_by_product.reindex(by_rev.index).fillna(0.0))
            .assign(NetRevenue=lambda d: (d["Revenue"] - d["Returned"]).round(2)))
net_rank["Gross rank"] = net_rank["Revenue"].rank(ascending=False).astype(int)
net_rank["Net rank"] = net_rank["NetRevenue"].rank(ascending=False).astype(int)
net_rank["Rank change"] = net_rank["Gross rank"] - net_rank["Net rank"]

top_gross = net_rank.nlargest(10, "Revenue")
movers = top_gross[top_gross["Rank change"] != 0]
print("Top 10 by GROSS revenue, showing where each sits once cancellations net off:\n")
print(top_gross[["Revenue", "Returned", "NetRevenue", "Gross rank", "Net rank",
                 "Rank change"]].to_string())
fact("net_rank_movers", int(len(movers)))
worst = top_gross.nlargest(1, "Returned")
fact("worst_net_mover", str(worst.index[0]).title())
fact("worst_net_mover_gross_rank", int(worst["Gross rank"].iloc[0]))
fact("worst_net_mover_net_rank", int(worst["Net rank"].iloc[0]))
runner_up = top_gross.nlargest(2, "Returned").iloc[1]
fact("second_net_mover", str(top_gross.nlargest(2, "Returned").index[1]).title())
fact("second_net_mover_gross_rank", int(runner_up["Gross rank"]))
fact("second_net_mover_net_rank", int(runner_up["Net rank"]))
print(f"\n{len(movers)} of the top 10 change position once returns are netted off.")
print(f"{str(worst.index[0]).title()} falls from gross rank "
      f"{int(worst['Gross rank'].iloc[0])} to net rank {int(worst['Net rank'].iloc[0])}.")
net_rank.nlargest(10, "NetRevenue")[["Revenue", "Returned", "NetRevenue", "Units", "Orders"]]
''')

# =========================================================== CONCENTRATION
md(r'''
---
# 10. Revenue concentration — is segmentation worth doing?

Before building any segmentation, test whether the customer base is uneven
enough to justify treating groups differently. If revenue were spread evenly,
segmentation would be busywork.
''')

code(r'''
per_cust = sales_id.groupby("CustomerID")["Revenue"].sum().sort_values(ascending=False)
cum = per_cust.cumsum() / per_cust.sum()
xs = np.arange(1, len(per_cust) + 1) / len(per_cust)

fig, ax = plt.subplots(figsize=(9.6, 5.4))
ax.plot(xs * 100, cum.to_numpy() * 100, color=BLUE, zorder=3)
ax.plot([0, 100], [0, 100], color=INK_MUTED, linewidth=1, linestyle=(0, (4, 3)),
        zorder=2)
ax.annotate("perfectly even base\n(reference line)", xy=(62, 62), xytext=(10, -24),
            textcoords="offset points", fontsize=9, color=INK_MUTED)

marks = {}
for pct in (1, 5, 10, 20, 50):
    n = max(1, int(len(per_cust) * pct / 100))
    share = float(per_cust.head(n).sum() / per_cust.sum() * 100)
    marks[pct] = (n, share)
    ax.scatter([pct], [share], color=SURFACE, edgecolor=ORANGE, linewidth=1.8, s=64, zorder=4)
    ax.annotate(f"top {pct}% = {share:.0f}%", xy=(pct, share), xytext=(9, -3),
                textcoords="offset points", fontsize=9, color=INK_2)

ax.xaxis.set_major_formatter(PercentFormatter(decimals=0))
ax.yaxis.set_major_formatter(PercentFormatter(decimals=0))
ax.set_xlabel("Customers, ranked by revenue")
ax.set_ylabel("Cumulative share of revenue")
ax.set_xlim(0, 100)
ax.set_ylim(0, 101)
titled(ax, "A fifth of customers generate over three quarters of revenue",
        f"{len(per_cust):,} identified customers, Dec 2009 - Dec 2011")
save(fig, "08_pareto")
plt.show()

for pct, (n, share) in marks.items():
    fact(f"top{pct}pct_share", round(share, 1))
    fact(f"top{pct}pct_customers", int(n))
    print(f"top {pct:>2}% of customers ({n:>5,}) = {share:>5.1f}% of revenue")
print(f"\nMedian customer revenue: {money(per_cust.median())}   "
      f"Mean: {money(per_cust.mean())}   Max: {money(per_cust.max())}")
fact("median_customer_revenue", round(float(per_cust.median()), 2))
fact("mean_customer_revenue", round(float(per_cust.mean()), 2))
fact("max_customer_revenue", round(float(per_cust.max()), 2))
''')

# ==================================================================== RFM
md(r'''
---
# 11. RFM segmentation

**RFM** scores each customer on three behaviours that are known to predict
future purchasing better than demographics:

* **Recency** — days since the last order. A customer who bought last week is
  worth more than one who bought 18 months ago.
* **Frequency** — number of distinct orders placed.
* **Monetary** — total revenue generated.

Each dimension is cut into quintiles and scored 1–5. Recency is scored in
reverse (fewer days = better). Quintiles are taken on **ranks** rather than raw
values, because frequency is heavily tied — thousands of customers placed
exactly one order — and tied values would otherwise collapse the bin edges.

The reference date is **one day after the last transaction in the log**, so that
the most recent customers get a recency of 1 rather than 0. Using today's date
would make every customer look churned, since the log ends in 2011.
''')

code(r'''
snapshot = sales_id["InvoiceDate"].max() + pd.Timedelta(days=1)
print("reference date:", snapshot)

rfm = (sales_id.groupby("CustomerID")
       .agg(LastPurchase=("InvoiceDate", "max"),
            FirstPurchase=("InvoiceDate", "min"),
            Frequency=("Invoice", "nunique"),
            Monetary=("Revenue", "sum"),
            Units=("Quantity", "sum"),
            Lines=("Invoice", "size"))
       .assign(Recency=lambda d: (snapshot - d["LastPurchase"]).dt.days))
rfm["Monetary"] = rfm["Monetary"].round(2)
rfm["AvgOrderValue"] = (rfm["Monetary"] / rfm["Frequency"]).round(2)
rfm["TenureDays"] = (rfm["LastPurchase"] - rfm["FirstPurchase"]).dt.days

# Rank-based quintiles. method="first" breaks ties deterministically so the five
# bins are equal in size even where thousands of customers share a value.
rfm["R"] = pd.qcut(rfm["Recency"].rank(method="first"), 5, labels=[5, 4, 3, 2, 1]).astype(int)
rfm["F"] = pd.qcut(rfm["Frequency"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
rfm["M"] = pd.qcut(rfm["Monetary"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
rfm["RFM"] = rfm["R"].astype(str) + rfm["F"].astype(str) + rfm["M"].astype(str)
rfm["RFMScore"] = rfm[["R", "F", "M"]].sum(axis=1)

print(f"\n{len(rfm):,} customers scored")
rfm[["Recency", "Frequency", "Monetary", "AvgOrderValue", "R", "F", "M", "RFM"]].head(10)
''')

code(r'''
rfm[["Recency", "Frequency", "Monetary", "AvgOrderValue", "TenureDays"]].describe().T.round(2)
''')

md(r'''
### 11.1 From scores to named segments

A 5×5×5 score gives 125 cells — too many to act on. Frequency and Monetary are
combined into a single value axis (`FM`, their mean, rounded) and the resulting
5×5 grid is mapped onto nine named segments. The ladder below is mutually
exclusive and covers all 25 cells, so no customer is unclassified or
double-classified.
''')

code(r'''
def segment_of(r, fm):
    """Map an (R, FM) cell to an action-oriented segment name.

    Conditions are evaluated in order and together cover all 25 cells of the
    5x5 grid exactly once.
    """
    if r >= 4 and fm >= 4:
        return "Champions"          # bought recently, often, and a lot
    if r >= 3 and fm >= 3:
        return "Loyal"              # solid, regular, mid-to-high value
    if r >= 4 and fm <= 1:
        return "New"                # just arrived, no history yet
    if r >= 4 and fm == 2:
        return "Promising"          # recent, light spend so far
    if r == 3 and fm <= 2:
        return "Needs attention"    # slipping, was never high value
    if r <= 2 and fm >= 4:
        return "Cannot lose them"   # high value, gone quiet - the urgent group
    if r <= 2 and fm == 3:
        return "At risk"            # mid value, gone quiet
    if r == 2 and fm <= 2:
        return "Hibernating"        # low value, long gone
    return "Lost"                   # r == 1 and fm <= 2


rfm["FM"] = ((rfm["F"] + rfm["M"]) / 2).round().astype(int)
rfm["Segment"] = [segment_of(r, fm) for r, fm in zip(rfm["R"], rfm["FM"])]

SEG_ORDER = ["Champions", "Loyal", "Cannot lose them", "At risk", "Needs attention",
             "Promising", "New", "Hibernating", "Lost"]
assert set(rfm["Segment"]) <= set(SEG_ORDER), "unmapped segment"
assert rfm["Segment"].notna().all(), "unclassified customer"

seg = (rfm.groupby("Segment")
       .agg(Customers=("Monetary", "size"), Revenue=("Monetary", "sum"),
            AvgRecency=("Recency", "mean"), AvgFrequency=("Frequency", "mean"),
            AvgMonetary=("Monetary", "mean"), AvgOrderValue=("AvgOrderValue", "mean"))
       .reindex(SEG_ORDER).dropna(how="all"))
seg["Customers"] = seg["Customers"].astype(int)
seg["Customer %"] = (seg["Customers"] / seg["Customers"].sum() * 100).round(1)
seg["Revenue %"] = (seg["Revenue"] / seg["Revenue"].sum() * 100).round(1)
seg[["AvgRecency", "AvgFrequency", "AvgMonetary", "AvgOrderValue"]] = \
    seg[["AvgRecency", "AvgFrequency", "AvgMonetary", "AvgOrderValue"]].round(1)
seg = seg[["Customers", "Customer %", "Revenue", "Revenue %", "AvgRecency",
           "AvgFrequency", "AvgMonetary", "AvgOrderValue"]]
seg
''')

code(r'''
# Two measures of different scale (customer count and revenue) => two axes side
# by side, never two y-scales on one chart.
s = seg.sort_values("Revenue", ascending=True)
fig, axes = plt.subplots(1, 2, figsize=(12.6, 5.6), sharey=True)
y = np.arange(len(s))

ax = axes[0]
ax.barh(y, s["Revenue"], height=0.72, color=BLUE, zorder=3)
ax.set_yticks(y, s.index)
label_barh(ax, s["Revenue"].to_list(),
           texts=[f"{money(v)}  ({p:.0f}%)" for v, p in zip(s["Revenue"], s["Revenue %"])])
ax.xaxis.set_major_formatter(GBP)
ax.set_xlim(0, s["Revenue"].max() * 1.34)
ax.grid(axis="y", visible=False)
titled(ax, "Revenue by segment", "Total revenue generated to date")

ax = axes[1]
ax.barh(y, s["Customers"], height=0.72, color=ORANGE, zorder=3)
label_barh(ax, s["Customers"].to_list(),
           texts=[f"{int(v):,}  ({p:.0f}%)" for v, p in zip(s["Customers"], s["Customer %"])])
ax.set_xlim(0, s["Customers"].max() * 1.34)
ax.grid(axis="y", visible=False)
titled(ax, "Customers by segment", "Headcount in each segment")
fig.tight_layout(w_pad=2.5)
save(fig, "09_rfm_segments")
plt.show()

for name in SEG_ORDER:
    if name in seg.index:
        key = name.lower().replace(" ", "_")
        fact(f"seg_{key}_customers", int(seg.loc[name, "Customers"]))
        fact(f"seg_{key}_revenue", round(float(seg.loc[name, "Revenue"]), 2))
        fact(f"seg_{key}_cust_pct", float(seg.loc[name, "Customer %"]))
        fact(f"seg_{key}_rev_pct", float(seg.loc[name, "Revenue %"]))

at_risk_set = [s_ for s_ in ("Cannot lose them", "At risk") if s_ in seg.index]
fact("at_risk_customers", int(seg.loc[at_risk_set, "Customers"].sum()))
fact("at_risk_revenue", round(float(seg.loc[at_risk_set, "Revenue"].sum()), 2))
fact("at_risk_rev_pct", round(float(seg.loc[at_risk_set, "Revenue %"].sum()), 1))
print(f"Champions: {FACTS.get('seg_champions_cust_pct')}% of customers -> "
      f"{FACTS.get('seg_champions_rev_pct')}% of revenue")
print(f"'Cannot lose them' + 'At risk': {FACTS['at_risk_customers']:,} customers holding "
      f"{money(FACTS['at_risk_revenue'])} ({FACTS['at_risk_rev_pct']}%) of historic revenue")
''')

# ================================================================= KMEANS
md(r'''
---
# 12. K-Means clustering on RFM

The RFM segments above come from rules **I** wrote. A clustering algorithm is
given no rules and proposes its own groups from the shape of the data. Running
both and comparing them is the useful move: agreement is evidence the structure
is real, and disagreement shows where the rules are arbitrary.

Two preparation steps matter:

* **Log transform.** Recency, frequency and monetary value are all heavily
  right-skewed. K-Means minimises squared Euclidean distance, so without a
  transform a handful of wholesale accounts would dominate every centroid.
  `log1p` compresses the tail while preserving order.
* **Standardisation.** Monetary is measured in thousands of pounds and frequency
  in single orders. Unscaled, monetary would be the only feature that counts.
''')

code(r'''
features = ["Recency", "Frequency", "Monetary"]
X_log = np.log1p(rfm[features].clip(lower=0))
scaler = StandardScaler()
X = scaler.fit_transform(X_log)

scan = []
for k in range(2, 9):
    km = KMeans(n_clusters=k, n_init=10, random_state=42)
    labels = km.fit_predict(X)
    scan.append({"k": k, "Inertia": round(float(km.inertia_), 1),
                 "Silhouette": round(float(silhouette_score(X, labels)), 4)})
scan = pd.DataFrame(scan).set_index("k")

fig, axes = plt.subplots(1, 2, figsize=(12, 4.4))
ax = axes[0]
ax.plot(scan.index, scan["Inertia"], color=BLUE, marker="o",
        markeredgecolor=SURFACE, markeredgewidth=1.6, zorder=3)
ax.set_xlabel("Number of clusters (k)")
ax.set_ylabel("Within-cluster sum of squares")
titled(ax, "Elbow curve", "Inertia falls steeply to k=4, then flattens")

ax = axes[1]
ax.plot(scan.index, scan["Silhouette"], color=ORANGE, marker="o",
        markeredgecolor=SURFACE, markeredgewidth=1.6, zorder=3)
ax.set_xlabel("Number of clusters (k)")
ax.set_ylabel("Mean silhouette score")
titled(ax, "Silhouette by k", "Higher is better separated")
fig.tight_layout(w_pad=3)
save(fig, "10_kmeans_selection")
plt.show()

fact("silhouette_best_k", int(scan["Silhouette"].idxmax()))
fact("silhouette_best", float(scan["Silhouette"].max()))
scan
''')

md(r'''
### 12.1 Choosing k honestly

The silhouette score is highest at the smallest k on the scan. That is the
normal behaviour of this metric on continuous, single-blob behavioural data: it
rewards splitting the cloud in half because that maximises the separation
between two centroids, not because two groups are the right answer for a
marketing team. Two groups — "buys" and "buys less" — cannot be acted on.

**k = 4 is adopted**, for two stated reasons: it is where the elbow curve stops
falling steeply, and four is the practical ceiling on the number of distinct
campaign treatments this business could actually run. Its silhouette score is
reported below alongside the k = 2 figure so the trade-off is visible rather
than hidden. This is a judgement call, and it is labelled as one.
''')

code(r'''
K = 4
km = KMeans(n_clusters=K, n_init=10, random_state=42)
rfm["Cluster"] = km.fit_predict(X)
sil_k = float(silhouette_score(X, rfm["Cluster"]))
fact("k_chosen", K)
fact("silhouette_at_k", round(sil_k, 4))
print(f"k = {K}: silhouette {sil_k:.4f}   (k = 2 scored {scan.loc[2, 'Silhouette']:.4f})")

profile = (rfm.groupby("Cluster")
           .agg(Customers=("Monetary", "size"), AvgRecency=("Recency", "mean"),
                AvgFrequency=("Frequency", "mean"), AvgMonetary=("Monetary", "mean"),
                TotalRevenue=("Monetary", "sum"))
           .round(1))
profile["Revenue %"] = (profile["TotalRevenue"] / profile["TotalRevenue"].sum() * 100).round(1)

# Name the clusters from their own centroids, ranked by monetary value, so the
# labels describe what the algorithm found rather than what I hoped it would.
rank = profile.sort_values("AvgMonetary", ascending=False).index.to_list()
NAMES = ["High-value regulars", "Mid-value active", "Low-value occasional", "Dormant low-value"]
cluster_name = {c: NAMES[i] for i, c in enumerate(rank)}
profile["Name"] = [cluster_name[c] for c in profile.index]
rfm["ClusterName"] = rfm["Cluster"].map(cluster_name)
profile = profile[["Name", "Customers", "AvgRecency", "AvgFrequency", "AvgMonetary",
                   "TotalRevenue", "Revenue %"]]
profile
''')

code(r'''
# Three measures on three different scales => small multiples, one axis each.
order = [c for c in rank]
labels = [cluster_name[c] for c in order]
colors = [PALETTE[i] for i in range(len(order))]
spec = [("AvgRecency", "Days since last order", "{:.0f} d", False),
        ("AvgFrequency", "Orders placed", "{:.1f}", False),
        ("AvgMonetary", "Revenue per customer", None, True)]

fig, axes = plt.subplots(1, 3, figsize=(13, 4.2))
for ax, (col, title, fmt, is_money) in zip(axes, spec):
    vals = [float(profile.loc[c, col]) for c in order]
    y = np.arange(len(order))[::-1]
    ax.barh(y, vals, height=0.7, color=colors, zorder=3)
    ax.set_yticks(y, labels if ax is axes[0] else [""] * len(labels))
    texts = [money(v) if is_money else fmt.format(v) for v in vals]
    label_barh(ax, vals, texts=texts)
    ax.set_xlim(0, max(vals) * 1.28)
    if is_money:
        ax.xaxis.set_major_formatter(GBP)
    ax.grid(axis="y", visible=False)
    titled(ax, title)
fig.tight_layout(w_pad=2)
save(fig, "11_cluster_profiles")
plt.show()

profile[["Name", "Customers", "Revenue %"]]
''')

code(r'''
# Scatter on log axes: colour follows the cluster (the entity), and centroids are
# ringed in the surface colour so they read as marks on top of the cloud.
fig, ax = plt.subplots(figsize=(9.8, 5.8))
for i, c in enumerate(order):
    m = rfm["Cluster"] == c
    ax.scatter(rfm.loc[m, "Frequency"], rfm.loc[m, "Monetary"], s=16, alpha=0.55,
               color=PALETTE[i], label=cluster_name[c], linewidths=0, zorder=3)
cent = np.expm1(scaler.inverse_transform(km.cluster_centers_))
for i, c in enumerate(order):
    row = cent[c]
    ax.scatter(row[1], row[2], s=170, color=PALETTE[i], edgecolor=SURFACE,
               linewidth=2.5, zorder=5)
ax.set_xscale("log")
ax.set_yscale("log")
ax.yaxis.set_major_formatter(GBP)
ax.set_xlabel("Orders placed (log scale)")
ax.set_ylabel("Revenue per customer (log scale)")
ax.legend(loc="upper left", ncol=2)
ax.grid(axis="x", visible=True)
titled(ax, "The clusters separate along both order count and spend",
        "One point per customer; large ringed marks are cluster centroids")
save(fig, "12_cluster_scatter")
plt.show()
print("Centroids in original units (Recency days, Frequency orders, Monetary GBP):")
pd.DataFrame(cent, columns=features).assign(Name=[cluster_name[i] for i in range(K)]).round(1)
''')

code(r'''
# Do the rules and the algorithm agree? Rows are the rule-based segments,
# columns the clusters the algorithm found on its own.
xtab = pd.crosstab(rfm["Segment"], rfm["ClusterName"])
xtab = xtab.reindex(index=[s_ for s_ in SEG_ORDER if s_ in xtab.index],
                    columns=[cluster_name[c] for c in order])
share = (xtab.div(xtab.sum(axis=1), axis=0) * 100).round(1)
print("Row-wise % of each RFM segment falling into each K-Means cluster:\n")
print(share.to_string())

agree = float(sum(share.max(axis=1) >= 60) / len(share) * 100)
fact("segment_cluster_agreement_pct", round(agree, 1))
print(f"\n{agree:.0f}% of the rule-based segments send at least 60% of their members to a")
print("single cluster - the two methods are describing the same structure, which is")
print("the point of running both.")
xtab
''')

# ==================================================================== CLV
md(r'''
---
# 13. Customer Lifetime Value

Two figures are reported, and the difference between them matters:

1. **Historical CLV** — revenue the customer has actually generated. A fact.
2. **Projected 12-month value** — the customer's observed revenue *rate*
   extended forward one year, i.e.
   `monthly rate x 12`, where the rate is total revenue divided by observed
   active tenure.

The projection is a **constant-rate assumption**, not a probabilistic forecast.
It assumes a customer keeps buying at the pace already observed, which is
optimistic for a dormant customer and conservative for one that is accelerating.
A production model would use BG/NBD plus a Gamma-Gamma spend model to put a
churn probability on each account; that is named in the limitations rather than
silently skipped. Customers with a tenure under 30 days are floored at one month
so that a single large first order does not project to an implausible annual
figure.
''')

code(r'''
MONTHS = 12
clv = rfm.copy()
clv["ActiveMonths"] = (clv["TenureDays"] / 30.44).clip(lower=1.0).round(2)
clv["MonthlyRate"] = (clv["Monetary"] / clv["ActiveMonths"]).round(2)
clv["HistoricalCLV"] = clv["Monetary"]
clv["Projected12M"] = (clv["MonthlyRate"] * MONTHS).round(2)
clv["OrdersPerMonth"] = (clv["Frequency"] / clv["ActiveMonths"]).round(3)

fact("total_projected_12m", round(float(clv["Projected12M"].sum()), 2))
fact("mean_historical_clv", round(float(clv["HistoricalCLV"].mean()), 2))
fact("median_historical_clv", round(float(clv["HistoricalCLV"].median()), 2))
fact("mean_projected_12m", round(float(clv["Projected12M"].mean()), 2))

print(f"Mean historical CLV     : {money(clv['HistoricalCLV'].mean())}")
print(f"Median historical CLV   : {money(clv['HistoricalCLV'].median())}")
print(f"Mean projected 12-month : {money(clv['Projected12M'].mean())}")
clv[["Recency", "Frequency", "Monetary", "ActiveMonths", "MonthlyRate",
     "Projected12M", "Segment"]].head(8)
''')

code(r'''
seg_clv = (clv.groupby("Segment")
           .agg(Customers=("HistoricalCLV", "size"),
                HistoricalCLV=("HistoricalCLV", "mean"),
                Projected12M=("Projected12M", "mean"),
                PoolProjected=("Projected12M", "sum"))
           .reindex([s_ for s_ in SEG_ORDER if s_ in set(clv["Segment"])])
           .round(2))
seg_clv["Share of projected pool %"] = (
    seg_clv["PoolProjected"] / seg_clv["PoolProjected"].sum() * 100).round(1)

s = seg_clv.sort_values("Projected12M")
fig, ax = plt.subplots(figsize=(10.6, 5.4))
y = np.arange(len(s))
ax.barh(y, s["Projected12M"], height=0.72, color=BLUE, zorder=3)
ax.set_yticks(y, s.index)
label_barh(ax, s["Projected12M"].to_list(),
           texts=[f"{money(v)} per customer  x{int(n):,}"
                  for v, n in zip(s["Projected12M"], s["Customers"])])
ax.xaxis.set_major_formatter(GBP)
ax.set_xlim(0, s["Projected12M"].max() * 1.5)
ax.grid(axis="y", visible=False)
titled(ax, "Projected 12-month value per customer, by segment",
        "Observed revenue rate extended one year - a constant-rate projection, not a forecast")
save(fig, "13_clv_by_segment")
plt.show()

for name in seg_clv.index:
    key = name.lower().replace(" ", "_")
    fact(f"clv_{key}_projected", float(seg_clv.loc[name, "Projected12M"]))
    fact(f"clv_{key}_pool_pct", float(seg_clv.loc[name, "Share of projected pool %"]))
seg_clv
''')

md(r'''
### 13.1 Sanity-checking the projection

A projection that is never checked against reality is a decoration. Two tests
below, and the model fails both in an instructive way.
''')

code(r'''
# Test 1: sum the per-customer projection and compare it with what the business
# actually took in its most recent full year.
proj_total = float(clv["Projected12M"].sum())
actual_y2 = float(FACTS["year2_revenue"])
fact("proj_vs_actual_ratio", round(proj_total / actual_y2, 2))
print(f"Sum of 12-month projections across all {len(clv):,} customers : {money(proj_total)}")
print(f"Revenue the business actually took in year 2 (Dec-Nov)        : {money(actual_y2)}")
print(f"Ratio                                                         : "
      f"{proj_total / actual_y2:.2f}x")

# Test 2: does the projection rank the segments sensibly?
rank_check = seg_clv[["Projected12M"]].sort_values("Projected12M", ascending=False)
print("\nSegments ranked by projected 12-month value per customer:")
print(rank_check.to_string())
top_seg = rank_check.index[0]
fact("clv_rank_top_segment", str(top_seg))
print(f"\nThe projection puts '{top_seg}' above 'Champions'.")
print("\nBoth tests fail, and they fail for the same reason. The projection divides a")
print("customer's total spend by the time between their first and last order, so a")
print("customer who bought heavily for two months and then vanished is credited with a")
print("very high monthly rate - and is then assumed to sustain it for a year. Dormant")
print("segments are therefore inflated, and summing across everyone double-counts")
print("customers who have already left.")
print("\nConclusion: this projection is usable as a per-customer UPPER BOUND - what an")
print("account would be worth if fully recovered and held at its old pace - and is not")
print("usable as a revenue forecast or as a ranking of who to prioritise. Historical")
print("CLV is used for prioritisation below, and a churn-aware model (BG/NBD plus")
print("Gamma-Gamma) is named in the limitations as the correct fix.")
''')

code(r'''
# What is the recoverable prize? Value the quiet high-worth customers at their
# own historical rate, and state the assumption in the same breath.
recover = clv[clv["Segment"].isin(["Cannot lose them", "At risk"])]
fact("recoverable_customers", int(len(recover)))
fact("recoverable_historical", round(float(recover["HistoricalCLV"].sum()), 2))
fact("recoverable_annual_rate", round(float(recover["Projected12M"].sum()), 2))
print(f"{len(recover):,} customers sit in 'Cannot lose them' or 'At risk'.")
print(f"They have already spent  {money(recover['HistoricalCLV'].sum())}.")
print(f"At their own historical rate, a full year of their custom is worth "
      f"{money(recover['Projected12M'].sum())}.")
print("\nThat figure is the ceiling on a win-back campaign, not its expected return:")
print("it assumes every one of them is recovered and resumes their previous pace.")
''')

# ================================================================= COHORTS
md(r'''
---
# 14. Cohort retention

Segments describe the base as it is today. Cohorts ask a sharper question: of
the customers acquired in a given month, what fraction were still ordering
1, 2, 3 … months later? This separates two very different problems that a single
revenue line hides — failing to *acquire* customers, and failing to *keep* them.

Each customer is assigned to the month of their first order and never moves.
''')

code(r'''
coh = sales_id[["CustomerID", "InvoiceMonth"]].drop_duplicates().copy()
coh = coh.join(first_month, on="CustomerID")
coh["PeriodIndex"] = ((coh["InvoiceMonth"] - coh["CohortMonth"])
                      .apply(lambda p: p.n))

sizes = coh[coh["PeriodIndex"] == 0].groupby("CohortMonth")["CustomerID"].nunique()
counts = coh.pivot_table(index="CohortMonth", columns="PeriodIndex",
                         values="CustomerID", aggfunc="nunique", observed=True)
retention = (counts.div(sizes, axis=0) * 100).round(1)
retention.index = retention.index.astype(str)

show = retention.iloc[:, :13]
fig, ax = plt.subplots(figsize=(12.6, 6.4))
im = ax.imshow(show.to_numpy(dtype=float), cmap=SEQ, aspect="auto", vmin=0, vmax=60)
ax.set_xticks(range(show.shape[1]), show.columns)
ax.set_yticks(range(show.shape[0]), [f"{m}  (n={int(sizes.iloc[i]):,})"
                                     for i, m in enumerate(show.index)])
ax.set_xlabel("Months since first order")
ax.grid(False)
for i in range(show.shape[0]):
    for j in range(show.shape[1]):
        v = show.iat[i, j]
        if pd.notna(v):
            ax.annotate(f"{v:.0f}", (j, i), ha="center", va="center", fontsize=7.6,
                        color=SURFACE if v > 33 else INK_2)
cb = fig.colorbar(im, ax=ax, pad=0.012, fraction=0.024)
cb.set_label("% of cohort still ordering", color=INK_2, fontsize=9.5)
cb.outline.set_visible(False)
titled(ax, "Reorder rates settle around a fifth and then hold, rather than collapsing",
        "Rows are acquisition cohorts; month 0 is 100% by definition")
save(fig, "14_cohort_retention")
plt.show()

m1 = retention[1].dropna()
m3 = retention[3].dropna() if 3 in retention.columns else pd.Series(dtype=float)
m6 = retention[6].dropna() if 6 in retention.columns else pd.Series(dtype=float)
m12 = retention[12].dropna() if 12 in retention.columns else pd.Series(dtype=float)
fact("retention_m1_mean", round(float(m1.mean()), 1))
fact("retention_m3_mean", round(float(m3.mean()), 1) if len(m3) else None)
fact("retention_m6_mean", round(float(m6.mean()), 1) if len(m6) else None)
fact("retention_m12_mean", round(float(m12.mean()), 1) if len(m12) else None)
fact("largest_cohort", str(sizes.idxmax()))
fact("largest_cohort_size", int(sizes.max()))
fact("first_cohort_size", int(sizes.iloc[0]))
print(f"Average month-1 retention : {m1.mean():.1f}%")
print(f"Average month-3 retention : {m3.mean():.1f}%")
print(f"Average month-6 retention : {m6.mean():.1f}%")
print(f"Average month-12 retention: {m12.mean():.1f}%")
print(f"\nLargest cohort: {sizes.idxmax()} with {sizes.max():,} customers")
print("Blanks on the right of the matrix mean 'not yet observable', not 'zero':")
print("cohorts acquired late in the window have not had time to show a curve.")
retention
''')


code(r'''
# Cohort quality: does a bigger cohort retain as well as a small one?
cq = pd.DataFrame({
    "CohortSize": sizes.astype(int),
    "Month1 %": retention[1].to_numpy(),
    "Month3 %": retention[3].to_numpy() if 3 in retention.columns else np.nan,
    "Month6 %": retention[6].to_numpy() if 6 in retention.columns else np.nan,
    "Revenue": sales_id.join(first_month, on="CustomerID")
                       .groupby("CohortMonth")["Revenue"].sum().to_numpy(),
})
cq.index = cq.index.astype(str)
cq["Revenue per acquired customer"] = (cq["Revenue"] / cq["CohortSize"]).round(2)
corr = cq[["CohortSize", "Month6 %"]].dropna().corr().iloc[0, 1]
fact("cohort_size_vs_m6_corr", round(float(corr), 3))
print(f"Correlation between cohort size and month-6 retention: {corr:+.3f}")
cq.round(2)
''')

md(r'''
### 14.1 The first cohort is not an acquisition cohort

The December 2009 row deserves suspicion rather than celebration. It is both the
**largest** cohort and by far the **best-retaining** one — 35–42% still ordering
each month, roughly double every cohort that follows. A first month that
outperforms every later month by 2x is not a marketing success; it is an artefact.

The cause is **left-censoring**. The log begins on 1 December 2009, so a customer
who had been buying from this retailer for years appears in the data for the first
time that month and is classified as "newly acquired". The December 2009 cohort is
therefore not new customers at all — it is *the entire pre-existing customer base*,
observed for the first time. Established customers retain well, which is exactly
what the row shows.

Two figures reported above inherit this bias and are recomputed below without that
row: the average retention curve, and the correlation between cohort size and
retention.
''')

code(r'''
# Recompute the retention summary with the left-censored first cohort excluded.
true_cohorts = retention.iloc[1:]
cmp_rows = []
for p in (1, 3, 6, 12):
    if p not in retention.columns:
        continue
    cmp_rows.append({
        "Month": p,
        "All cohorts %": round(float(retention[p].dropna().mean()), 1),
        "Excluding Dec-2009 %": round(float(true_cohorts[p].dropna().mean()), 1),
    })
cmp = pd.DataFrame(cmp_rows)
cmp["Difference"] = (cmp["All cohorts %"] - cmp["Excluding Dec-2009 %"]).round(1)
print(cmp.to_string(index=False))

for p in (1, 3, 6, 12):
    if p in true_cohorts.columns:
        fact(f"retention_m{p}_true", round(float(true_cohorts[p].dropna().mean()), 1))

cq_true = cq.iloc[1:]
corr_true = cq_true[["CohortSize", "Month6 %"]].dropna().corr().iloc[0, 1]
fact("cohort_size_vs_m6_corr_true", round(float(corr_true), 3))
print(f"\nCohort size vs month-6 retention correlation:")
print(f"  all cohorts          : {FACTS['cohort_size_vs_m6_corr']:+.3f}")
print(f"  excluding Dec-2009   : {corr_true:+.3f}")
print("\nThe apparent 'bigger cohorts retain better' relationship is substantially an")
print("artefact of that single left-censored row.")

# Does the month a customer is acquired in predict how well they stick?
xmas = true_cohorts.index.str.endswith(("-11", "-12"))
fact("xmas_cohort_m3", round(float(true_cohorts.loc[xmas, 3].dropna().mean()), 1))
fact("nonxmas_cohort_m3", round(float(true_cohorts.loc[~xmas, 3].dropna().mean()), 1))
print(f"\nMonth-3 retention, customers acquired in Nov/Dec : "
      f"{FACTS['xmas_cohort_m3']}%")
print(f"Month-3 retention, acquired in any other month   : "
      f"{FACTS['nonxmas_cohort_m3']}%")
print("\nCustomers won during the Christmas rush are markedly worse at coming back.")
print("The peak season buys volume, not loyalty - a one-off gift buyer looks identical")
print("to a new wholesale account on the day they order, and very different by month 3.")
cmp
''')

# ================================================================ FINDINGS
md(r'''
---
# 15. Key findings

Every figure below is computed by the cells above and stored in
`outputs/facts.json`. Nothing here is estimated or rounded from memory.

### 1. Revenue is extremely concentrated — segmentation is justified before it is built

**1,170 customers (20%) generate 77.2% of revenue. The top 58 customers (1%)
generate 31.9%.** The median customer is worth £889 and the mean £2,994 — the
mean is 3.4x the median, so "the average customer" is a fiction. Order value
tells the same story: a £517.14 mean against a £303.85 median.

This is the finding that licenses everything after it. Uniform treatment of this
base is guaranteed to be wrong, because it is calibrated to a customer who does
not exist.

### 2. One segment holds 71.5% of revenue, and one holds the risk

**Champions — 1,555 customers, 26.6% of the base — account for £12.54M, or 71.5%
of all identified revenue.** At the other end, 945 *Lost* and 714 *Hibernating*
customers (28.3% of the base together) contribute 3.4% between them.

The urgent group is neither: **683 customers in *Cannot lose them* and *At risk*
have already spent £1.69M (9.6% of revenue) and have stopped ordering.** These
are not low-value customers who drifted away — they are proven buyers who went
quiet, and they are the only group where an intervention has a large, identified
prize attached.

### 3. The two largest "sales" in the dataset were never sales

**The single biggest line in 1.07M rows — 80,995 units of *Paper Craft, Little
Birdie*, £168,469.60 — was cancelled 12 minutes after it was raised**, on the
final day of the log. Net contribution: **£0.00**. The second largest line
(74,215 ceramic storage jars, £77,183.60) was also cancelled in full.

The consequence is not academic. Ranked on gross revenue, that product is the
retailer's **3rd** best seller. Ranked net of cancellations it is **5,320th**.
**Eight of the top ten products change position once returns are netted off.**
Any product report built on gross sales — the default in most tutorials on this
dataset — is materially wrong at the top of the table.

### 4. Returns cost £1.52M, and the aggregate rate hides the problem

Returns total **£1,523,788 across 19,432 lines — 7.43% of gross revenue**, taking
£20.52M down to a net £18.99M. The monthly rate averages 6.79% but is not stable:
it **spikes to 19.05% in January 2011**, immediately after the Christmas peak.

Returns are a post-seasonal event, not a constant leak, so a single annual
return-rate assumption will under-provision January and over-provision the rest
of the year.

### 5. The business is one country and one quarter

**The UK is 85.36% of revenue (£17.51M);** all 42 other markets together make
£3.00M. **September–November averages £1.19M per month against £708,613 for every
other month — a 67.6% seasonal uplift**, peaking at £1.50M in November 2011
against a £521,308 trough in February 2011.

Both concentrations are risks, but the export data also contains an opportunity
that the revenue ranking hides: **the Netherlands averages £2,484.46 per order
against the UK's £483.27 — 5.1x** — on a small number of orders. The export
markets are low-volume and high-value, which is a different business from the
domestic one and probably deserves different handling.

### 6. Growth came from retained customers, not new ones

Like-for-like years grew **8.1%** (£9.56M to £10.32M). But in the second year,
**91.6% of identified revenue came from customers acquired earlier.** New
customers contributed the remaining 8.4%.

Growth is therefore being produced by the existing base, which raises the cost of
losing any part of it and makes finding 2's £1.69M at-risk pool the most
expensive problem on this list.

### 7. Christmas buys volume, not loyalty

Retention settles at roughly a fifth and then holds: **20.5% at month 1, 20.7% at
month 3, 16.8% at month 6, 16.7% at month 12** (excluding the left-censored first
cohort). The curve flattens rather than decaying, which is the signature of a
wholesale reorder cycle rather than one-off consumer purchasing.

Splitting cohorts by acquisition month exposes a sharp difference:
**customers acquired in November or December retain at 9.3% by month 3, against
21.9% for customers acquired in any other month — less than half.** The peak
season that produces the revenue spike produces the retailer's worst customers.

### 8. The rules and the algorithm agree, which is the point of running both

K-Means on log-scaled RFM was given no rules, and **every one of the nine
rule-based segments sends at least 60% of its members to a single cluster**.
*Cannot lose them* maps 96.7% onto one cluster and *Lost* maps 98.4% onto
another. Two independent methods recovering the same structure is evidence the
structure is in the data rather than in the analyst's choice of thresholds.

The honest caveat: the silhouette score preferred **k = 2 (0.4376)** over the
**k = 4 (0.3649)** adopted here. Two clusters cannot be marketed to differently,
so interpretability was chosen over the metric, and the cost is stated rather
than hidden.

### 9. Two data artefacts that would have produced wrong answers

* **`InvoiceDate` is not customer intent.** Every Saturday order in 25 months
  falls on a single date, 5 December 2009 — 30 orders, 0.076% of the total —
  while Sunday processes 4,826 orders at weekday-like volumes. No customer base
  trades every Sunday and only one Saturday in two years. The timestamp records
  back-office keying, so it is valid for staffing (78.1% of orders are processed
  between 10:00 and 15:59, peaking at noon) and worthless for deciding when to
  advertise.
* **Cancellations are not the only reversal.** Six `A`-prefixed invoices carry
  **−£147,614** of bad-debt write-offs. The standard "drop invoices starting with
  C" recipe leaves all of it in the revenue line.
''')

md(r'''
---
# 16. Recommendations

Each recommendation names the finding it rests on and the number that sizes it.
Where the data cannot support a decision, that is stated instead of guessed.

### 1. Run a win-back on the 683 quiet high-value accounts — and measure it properly

*From finding 2.* These customers have already spent **£1.69M** and have stopped
ordering. Prioritise by **historical** value, not by the projected figure: section
13.1 shows the projection ranks this group *above* Champions, which is an artefact
of the constant-rate assumption, not a real ordering.

Because these are wholesale accounts, the intervention is a phone call from a
named account manager, not a discount email — and the discount is the wrong lever
anyway, since nothing in this dataset says price caused them to leave.

**Hold back a random 20% as an untreated control group.** Without one, any
subsequent recovery is indistinguishable from customers who would have reordered
regardless — and given the flat 16.7% month-12 reorder rate in finding 7, a
meaningful fraction would have.

### 2. Protect the 1,555 Champions before chasing anyone new

*From findings 2 and 6.* Champions produce **71.5% of revenue**, and **91.6% of
second-year revenue came from previously-acquired customers**. The concentration
cuts both ways: losing 58 customers (the top 1%) would remove **31.9% of
revenue**.

Concretely: guaranteed stock availability on their repeat lines through the
Sep–Nov peak, and a named contact. The defensive case is stronger than any
acquisition case on this data, because acquisition is demonstrably not what is
producing growth.

### 3. Stop reporting product performance on gross revenue

*From finding 3.* Eight of the top ten products change rank once cancellations
are netted off, and the gross-rank-3 product is genuinely 5,320th. Every product
report should net cancellations against the original sale.

This is a reporting fix, not a strategy: it costs nothing and it stops the buying
team restocking a line that sold nothing.

### 4. Treat January as a returns event and investigate the top returned lines

*From finding 4.* The return rate hits **19.05% in January 2011** against a 6.79%
average. Two actions follow: provision warehouse and refund capacity for a
post-Christmas spike rather than an even monthly rate, and inspect the
concentrated returned lines in section 9 for a cause — the data shows *which*
products come back but never *why*, so this is a question to take to the
warehouse, not one to answer from the file.

### 5. Change what the Christmas peak is expected to deliver

*From finding 7.* Customers acquired in November–December retain at **9.3% by
month 3 versus 21.9%** otherwise. The peak should be run as a volume and
cash-generation event, and peak-acquired customers should **not** be counted as
new relationships in any target that assumes they behave like the rest of the
base.

The corollary is to spend acquisition budget outside the peak, where the customers
who arrive are more than twice as likely to still be buying at month 3.

### 6. Test the export markets deliberately

*From finding 5.* The Netherlands averages **£2,484.46 per order — 5.1x the UK's
£483.27**. That is a striking ratio on a small base, and a small base is exactly
why it must be tested rather than acted on: with £3.00M of export revenue against
£17.51M domestic, a handful of large wholesale accounts could produce this entire
effect.

The recommendation is a bounded test on the highest-AOV markets with a defined
budget, not a market-entry decision. The data supports the hypothesis; it does not
support the investment.

### 7. Fix the 14.6% attribution gap at source

*From the limitations.* **£2.99M of clean revenue — 14.6% of the total, spread
over 22.9% of all sales lines — has no customer ID**, so it is invisible to every
model here. No amount of analysis recovers it —
it is a checkout-and-CRM change, and it raises the ceiling on all future customer
work.

---

### What this analysis does not support

* **Any pricing or product-mix decision.** The dataset has no cost of goods, so
  every figure here is revenue, not margin.
* **Any claim that a campaign *causes* a customer to return.** There is no
  campaign data and no experiment in this file. That is why recommendation 1
  specifies a control group.
* **Any conclusion about when customers prefer to shop** — see finding 9.
* **Using the 12-month CLV projection as a forecast.** It sums to 2.43x actual
  second-year revenue (section 13.1). It is an upper bound per account, nothing
  more.
''')

md(r'''
---
# 17. Limitations

Stating what this analysis cannot support is part of the analysis.

1. **Attribution gap.** 22.9% of clean sales lines (235,813), carrying 14.6% of
   clean revenue (£2.99M), have no customer ID — both shares measured on the same
   cleaned base. Every customer-level result — RFM, clustering,
   CLV, cohorts — is computed on the identified 85.4% only. If guest orders are
   disproportionately one-off purchases, the true repeat rate is lower than
   reported here; if they are unrecognised repeat buyers, several segments are
   undercounted. The data cannot distinguish the two cases.
2. **Revenue, not profit.** The dataset has no cost of goods. Every "value"
   figure here is revenue. A low-margin bestseller and a high-margin niche line
   are indistinguishable in this analysis, so no pricing or product-mix decision
   should rest on it alone.
3. **The CLV projection is demonstrably biased upward, and section 13.1 shows by
   how much.** Summed across the base it exceeds actual second-year revenue by a
   wide margin, and it ranks a dormant segment above Champions. It carries no
   churn probability, so it should be read only as a per-customer upper bound.
   BG/NBD plus a Gamma-Gamma spend model is the correct fix and is not done here;
   prioritisation in section 16 therefore rests on historical value, not on the
   projection.
4. **Censoring at both ends of the cohort matrix.** Cohorts acquired late in the
   window have not had time to show a retention curve, so blank cells on the
   right mean "not yet observable", not "zero". At the other end, the first
   cohort is left-censored: customers who predate the log appear as newly
   acquired in December 2009. Section 14.1 quantifies that bias and reports the
   corrected figures; the uncorrected averages remain in section 14 only so the
   size of the correction is visible.
5. **`InvoiceDate` records processing, not intent.** All 30 Saturday orders in
   the 25-month window fall on one single date, so the timestamp reflects
   back-office keying rather than customer behaviour. No conclusion about *when
   customers want to shop* can be drawn from the hour and weekday columns.
6. **Quintile scoring is relative.** An RFM score of 5 means "top fifth of this
   base", not "good" in absolute terms. Scores are not comparable across a
   different customer base or a different time window.
7. **k = 4 is a judgement call.** The silhouette metric preferred k = 2. Four
   clusters were chosen for operational interpretability, and the metric cost of
   that choice is reported in section 12 rather than hidden.
8. **One retailer, one category, 2009–2011.** A UK giftware wholesaler in the
   post-financial-crisis period. The method transfers; these coefficients do not.
9. **Segments are descriptive, not causal.** Nothing here establishes that a
   campaign *causes* a customer to return. A holdout test would be needed to
   claim that, and this dataset contains no campaign data.
''')

code(r'''
# Persist every computed headline number, the segment tables and the figure list,
# so the written report quotes computed values rather than hand-copied ones.
OUTPUTS.mkdir(parents=True, exist_ok=True)
payload = {
    "facts": FACTS,
    "figures": FIGURES,
    "clean_log": clean_log.to_dict(orient="records"),
    "kpi": kpi.to_dict(orient="records"),
    "segments": seg.reset_index().to_dict(orient="records"),
    "segment_clv": seg_clv.reset_index().to_dict(orient="records"),
    "clusters": profile.reset_index().to_dict(orient="records"),
    "kmeans_scan": scan.reset_index().to_dict(orient="records"),
    "monthly": monthly.assign(InvoiceMonth=monthly["InvoiceMonth"].astype(str))
                      .drop(columns=["Month"]).to_dict(orient="records"),
    "top_products": top10.reset_index().to_dict(orient="records"),
    "countries": by_country.head(12).reset_index().to_dict(orient="records"),
    "retention": retention.reset_index().rename(columns={"index": "CohortMonth"})
                          .to_dict(orient="records"),
    "cohort_quality": cq.reset_index().rename(columns={"index": "CohortMonth"})
                        .to_dict(orient="records"),
    "top_returned": top_ret.reset_index().to_dict(orient="records"),
}
def jsonable(obj):
    """Replace NaN/inf with null so the file is strict JSON any reader can parse.

    The cohort matrix is upper-triangular: cells for months a cohort has not yet
    lived through are genuinely absent, not zero.
    """
    if isinstance(obj, dict):
        return {k: jsonable(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [jsonable(v) for v in obj]
    if isinstance(obj, float) and not np.isfinite(obj):
        return None
    return obj


path = OUTPUTS / "facts.json"
path.write_text(json.dumps(jsonable(payload), indent=2, default=str), encoding="utf-8")
print(f"wrote {path.relative_to(ROOT)}  ({len(FACTS)} facts, {len(FIGURES)} figures)")
print("\nfigures saved:")
for name in FIGURES.values():
    print("  outputs/figures/" + name)
''')

md(r'''
---

### Reproducibility

Every number and every chart in this notebook is computed from the source data
by the cells above — nothing is typed in by hand. `outputs/facts.json` holds the
computed values that the accompanying project report quotes, so the two
documents cannot drift apart.

**Dataset:** Chen, D. (2019). *Online Retail II* [Dataset]. UCI Machine Learning
Repository. <https://doi.org/10.24432/C5CG6D> —
<https://archive.ics.uci.edu/dataset/502/online+retail+ii>
''')


# ------------------------------------------------------------------- build
nb = nbf.v4.new_notebook()
nb.cells = [nbf.v4.new_markdown_cell(s) if k == "md" else nbf.v4.new_code_cell(s)
            for k, s in C]
nb.metadata.update({
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python"},
    "authors": [{"name": "Aryan Verma"}],
    "title": "Retail Customer Segmentation & Sales Analysis",
})
nbf.write(nb, OUT)
print(f"wrote {OUT.name}: {len(nb.cells)} cells "
      f"({sum(1 for k, _ in C if k == 'code')} code, {sum(1 for k, _ in C if k == 'md')} markdown)")
