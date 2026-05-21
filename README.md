# Credit Risk Portfolio Analysis — Lending Club 2007–2018

A end-to-end credit risk analysis project built on a 5% random sample 
of Lending Club loan data (~113,000 loans). The project replicates the 
core analytical workflows used in bank credit portfolio management teams 
including delinquency analysis, vintage tracking, roll rate matrices, 
portfolio segmentation, and automated MIS reporting.

Built as preparation for a Data Analyst role in credit portfolio risk 
at a Singapore bank.

---

## Project Structure

credit-risk-lending-club/
├── data/
│   ├── raw/                        # Original Lending Club CSV
│   └── processed/
│       ├── accepted_cleaned.csv    # Cleaned dataset from EDA
│       └── accepted_features.csv  # Feature-engineered dataset
├── notebooks/
│   ├── 01_eda.ipynb                # Exploratory data analysis and cleaning
│   ├── 02_transform.ipynb         # Feature engineering
│   ├── 03_vintage_analysis.ipynb  # Vintage default rate analysis
│   ├── 04_roll_rate.ipynb         # Roll rate matrix and heatmap
│   ├── 05_portfolio_segmentation.ipynb  # Portfolio segmentation dashboard
│   └── 06_performance_trends.ipynb     # Performance trend analysis
├── scripts/
│   └── generate_report.py         # Automated MIS Excel report generator
├── outputs/                       # All saved charts and reports
└── README.md

---

## Dataset

**Source:** Lending Club Loan Data (2007–2018), via Kaggle  
**Sample:** 5% random sample — 113,244 loans, 111 columns after cleaning  
**Target variable:** `is_default` — 1 if loan was Charged Off or Defaulted, 0 otherwise

---

## Notebooks

### 01 — Exploratory Data Analysis
- Loaded raw dataset and profiled 150+ columns
- Dropped high-null columns, irrelevant identifiers, and post-origination leakage columns
- Final cleaned dataset: 113,244 rows × 95 columns

### 02 — Feature Engineering
- Created `is_default` binary flag from `loan_status` column
- Created `vintage` column in `YYYY-MM` format from `issue_d`
- Output saved as `accepted_features.csv` for all downstream notebooks

### 03 — Vintage Analysis
- Tracked default rates by loan origination cohort (vintage year)
- Built two charts:
  - Default rate by calendar month grouped by vintage year
  - Cumulative default rate by loan age per vintage year
- Key finding: 2015 vintage shows highest cumulative default rate,
  consistent with loose underwriting during rapid platform growth

### 04 — Roll Rate Matrix
- Mapped `loan_status` to DPD (Days Past Due) buckets:
  Current → Fully Paid → 1-15 DPD → 16-30 DPD → 31-120 DPD → 120+ DPD → Charged Off
- Built static roll rate matrix grouped by loan grade
- Visualised as annotated heatmap using matplotlib
- Key finding: Charge-off rate increases monotonically from 3.1% (Grade A)
  to 37.5% (Grade G)

### 05 — Portfolio Segmentation Dashboard
- 2×2 dashboard segmenting the portfolio by grade, loan purpose,
  and home ownership status
- Key findings:
  - Grade is the strongest predictor of default
  - Renters default at 15.4% vs mortgage holders at 11.5%
  - Credit card refinancing loans have the lowest default rate (10.6%)

### 06 — Performance Trend Analysis
- Yearly trends from 2008–2018 across four metrics:
  loan volume, default rate, average interest rate, average loan size
- Key findings:
  - Portfolio grew from 110 loans (2008) to 25,067 (2018)
  - Default rate peaked at 18.45% in 2015
  - 2017–2018 rates are understated due to incomplete loan maturity
  - Average loan size grew from $9,177 to $15,977 over the period

---

## Automated MIS Report

`scripts/generate_report.py` generates a dated Excel report automatically:

```bash
cd scripts
python generate_report.py
```

Output: `outputs/credit_portfolio_report_YYYY-MM-DD.xlsx`

**Four sheets:**
- **Portfolio Summary** — high-level KPIs: total loans, default rate,
  average interest rate, average DTI
- **Grade Analysis** — default rate, interest rate, loan size by grade
- **Yearly Trends** — year-over-year portfolio performance metrics
- **Purpose Analysis** — default rate by loan purpose

---

## Tools and Libraries

- **Python 3.14.3**
- **Pandas** — data manipulation and aggregation
- **Matplotlib** — all visualisations
- **NumPy** — numerical operations
- **OpenPyXL** — Excel report generation

---

## How to Run

```bash
# 1. Clone the repo
git clone https://github.com/irsyadudinsutiro/credit-risk-lending-club.git

# 2. Activate your virtual environment
source /path/to/your/venv/bin/activate

# 3. Install dependencies
pip install pandas matplotlib numpy openpyxl jupyter

# 4. Run notebooks in order (01 through 06) in VS Code or Jupyter

# 5. Generate the MIS report
cd scripts
python generate_report.py
```

---

## Context

This project was built as part of preparation for a Data Analyst 
role in credit portfolio risk management at a Singapore bank. The analytical 
workflows — vintage analysis, roll rate matrices, portfolio segmentation, 
and MIS reporting — are representative of day-to-day work in a bank credit 
risk team.