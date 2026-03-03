# SAAM PROJECT 2026 - PART I
## Group Assignment: Pacific Area | Scope 1

### Project Overview
This project implements minimum-variance portfolio allocation for Pacific region firms using Scope 1 carbon data. The analysis runs from 2014-2025 with annual rebalancing.

---

## 📁 Files Provided

### Python Scripts (Run in order)
1. **`saam_part1_complete.py`** - Main analysis script
   - Data loading and cleaning
   - Investment set construction
   - Portfolio optimization
   - Performance calculation
   - Results export

2. **`saam_part1_visualize.py`** - Visualization and reporting
   - Cumulative return plots
   - Performance analysis
   - Portfolio composition analysis
   - Template filling (if available)

### Required Input Files
Place these files in the same directory as the Python scripts:
- `Static_2025.xlsx`
- `DS_RI_T_USD_Y_2025-3.xlsx` (Annual Return Index)
- `DS_RI_T_USD_M_2025.xlsx` (Monthly Return Index)
- `DS_MV_T_USD_M_2025-2.xlsx` (Monthly Market Cap)

---

## 🚀 How to Run

### Step 1: Setup Environment
```bash
pip install pandas numpy scipy matplotlib seaborn openpyxl
```

### Step 2: Run Main Analysis
```bash
python saam_part1_complete.py
```

**Expected output files:**
- `part1_results.csv` - Monthly returns for both portfolios
- `part1_summary_statistics.csv` - Annualized performance metrics
- `part1_portfolio_compositions.csv` - Yearly portfolio weights

**Expected runtime:** 5-10 minutes (depending on system)

### Step 3: Create Visualizations
```bash
python saam_part1_visualize.py
```

**Expected output files:**
- `cumulative_returns_plot.png` - Main plot for report
- `additional_analysis_plots.png` - Supplementary analysis
- `summary_table_formatted.csv` - Formatted statistics table

---

## 📊 Output Files Explained

### 1. `part1_results.csv`
Monthly returns from Jan 2014 to Dec 2025

| Column | Description |
|--------|-------------|
| Date | Month-end date |
| MV_Return | Minimum Variance portfolio monthly return |
| VW_Return | Value-Weighted portfolio monthly return |
| MV_CumReturn | MV cumulative return (base = 1) |
| VW_CumReturn | VW cumulative return (base = 1) |

### 2. `part1_summary_statistics.csv`
Annualized performance metrics

| Metric | Description |
|--------|-------------|
| Annualized Return | Geometric mean annual return |
| Annualized Volatility | Standard deviation × √12 |
| Sharpe Ratio | Return / Volatility (risk-free rate = 0) |
| Min Monthly Return | Worst single-month return |
| Max Monthly Return | Best single-month return |
| Cumulative Return | Total return over entire period |

### 3. `part1_portfolio_compositions.csv`
Portfolio weights by year

| Column | Description |
|--------|-------------|
| Year | Investment year (2014-2025) |
| ISIN | Firm identifier |
| Weight | Portfolio weight (0-1) |

---

## 📈 Key Implementation Details

### Investment Set Construction
For each rebalancing year Y (end of Dec):
1. **Region filter**: Pacific firms only
2. **Price filter**: Annual RI ≥ 0.5 at year-end
3. **Data availability**: ≥36 months of valid returns in 10-year window
4. **Stale price filter**: ≤50% zero-return months in window
5. **Carbon data**: Required for Part II consistency

### Minimum Variance Optimization
**Objective:**
```
min α'Σα
subject to:
  Σ αᵢ = 1
  αᵢ ≥ 0 for all i
```

**Solver:** `scipy.optimize.minimize` with SLSQP method

**Rolling window:** 10 years (120 months) of data to estimate Σ

### Portfolio Returns Calculation
**Rebalancing:** Annual (December → next year)

**Monthly drift adjustment:**
```
wᵢ,ₜ = wᵢ,ₜ₋₁ × (1 + Rᵢ,ₜ) / (1 + Rₚ,ₜ)
```

### Value-Weighted Benchmark
**Weights:** Market cap at t-1 / Total market cap at t-1

**Rebalancing:** Monthly (uses lagged market caps)

---

## 🧪 Data Cleaning Applied

### Price Data (RI)
- RI < 0.5 → NaN (illiquid/rounding issues)
- Delisted firms: -100% return in delist month, NaN after
- Missing values: Filled with 0 in return calculations (no trade)

### Delist Detection
Pattern matched from NAME column:
- `DEAD - DELIST.DD/MM/YY`
- `DEAD - DEAD.DD/MM/YY`

### Stale Prices
Firms with >50% zero-return months in 10-year window excluded

---

## 📝 Template Filling Instructions

Once you have `Template for Part I-SAAM.xlsx`:

### Sheet Structure (Expected)
- **Column A:** Date
- **Column B:** MV Portfolio Monthly Returns
- **Column C:** VW Portfolio Monthly Returns
- **Section 2:** Summary Statistics

### Manual Filling
If auto-fill doesn't work:

1. **Monthly Returns:**
   - Copy `Date`, `MV_Return`, `VW_Return` from `part1_results.csv`
   - Paste into template columns A, B, C

2. **Summary Statistics:**
   - Copy from `part1_summary_statistics.csv`
   - Place MV stats in designated column
   - Place VW stats in adjacent column

---

## 🎯 Deliverable Checklist

For April 12 midnight submission:

- [ ] Filled `Template for Part I-SAAM.xlsx`
  - [ ] Monthly returns (Jan 2014 - Dec 2025)
  - [ ] Summary statistics for MV portfolio
  - [ ] Summary statistics for VW portfolio

For April 14 presentation:

- [ ] Cumulative return plot (`cumulative_returns_plot.png`)
- [ ] Summary statistics table
- [ ] Brief methodology explanation
- [ ] Key findings (which portfolio performed better?)

---

## 🔍 Validation Checks

### Sanity Checks to Run:

1. **Investment set size**
   - Should have 200-400 firms per year
   - Declining over time is normal (delistings)

2. **Portfolio weights**
   - All weights ≥ 0
   - Sum of weights = 1.0 per year
   - Max weight typically < 10% (diversified)

3. **Returns**
   - No NaN or infinite values
   - Monthly returns typically between -30% and +30%
   - Cumulative returns > 0 by 2025 expected

4. **Performance metrics**
   - Annualized volatility: 10-25% range typical
   - Sharpe ratio: 0.3-1.0 range typical for equity
   - MV portfolio should have lower volatility than VW

---

## 🐛 Troubleshooting

### Issue: "Optimization not converging"
**Solution:** Normal for some years with high correlation. Check that:
- Investment set > 10 firms
- Covariance matrix not singular
- Weights still sum to 1.0

### Issue: "Missing data for year X"
**Solution:** Check:
- Sufficient firms pass all filters?
- Are there gaps in monthly RI data?
- Try reducing `min_obs_months` to 30 if needed

### Issue: "Returns seem too high/low"
**Solution:** Verify:
- Monthly returns (not annualized) in output
- Cumulative returns start from 1.0 in Jan 2014
- Check for delisted firms with -100% returns

---

## 📧 Questions?

For methodology questions, refer to:
- `SAAM_Project_2026.pdf` (project guidelines)
- Lecture notes on minimum variance portfolios
- Section 2 of project document (Part I details)

For code issues:
- Check that all input files are in same directory
- Verify Python packages are installed
- Review console output for specific error messages

---

## 🎓 Academic Integrity Note

This code was developed with LLM assistance for:
- Debugging Datastream file format parsing
- Optimizing pandas operations for large datasets
- Structuring code organization and documentation

All methodological choices, portfolio construction logic, and results interpretation are the group's own work. See project submission for full LLM disclosure statement.

---

**Good luck with Part I! 🚀**

Questions or issues? Review the console output carefully - it provides detailed progress tracking and warnings for any data quality issues.
