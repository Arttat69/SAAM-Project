# SAAM Project — Code Audit & Revision Plan

## A. Main Issues Found

### Critical Errors

1. **Covariance estimation is non-standard and biased.**  
   `estimate_moments()` (line 274) fills missing demeaned returns with **zero** and divides by `tau = valid_counts.min()`. This is *not* the pairwise covariance described in the project brief (§2.1) or the reference report (§3.1). The project requires pairwise covariance where each (i,j) entry uses only the months where *both* assets have valid returns. The current implementation biases off-diagonal entries toward zero for assets with missing data, which distorts the minimum-variance solution. **This must be fixed.**

2. **`validate_oos_series` expects exactly 144 months (2014–2025), but data may only run through 2024.**  
   The code sets `END_YEAR_OOS = 2025`, expecting returns through Dec 2025. If the data only covers through end-2024 (as implied by the project: "carbon data at best covers 2002 to 2024"), this will crash. The project says rebalancing runs Dec 2013 to Dec 2024, giving OOS returns Jan 2014 – Dec 2025 = 144 months. This is correct *if* price data runs through Dec 2025. Flag for manual verification.

3. **Double forward-fill of CO2 panel creates no additional bug but is confusing.**  
   Lines 887 and 891 both do `co2_panel.T.ffill().T`. The first overwrites `co2_panel`; the second assigns to `co2_ff`. Not a bug, but sloppy.

### Methodological Concerns

4. **`MIN_OBS_MONTHS = 36` (3 years)** is the project's suggested minimum, but the reference report uses L=72 months. Either is defensible, but the student should document and justify the choice. If the goal is consistency with the reference report style, 72 months is stronger.

5. **VW benchmark weights in Part II use annual cap (`capA_ff[Y]`)** for the optimizer reference weights. This is correct for the annual rebalancing decision point, but differs from the monthly-rebalanced VW benchmark used for performance. This is fine methodologically since the TE optimization is annual.

6. **Only Scope 1 CO2 loaded (`CO2_FILE = "DS_CO2_SCOPE_1_Y_2025.xlsx"`).** The scope assignment depends on the group. If the group is assigned "Scope 1+2", a second file must be loaded and summed. If "Scope 1" only, this is correct. **Flag for manual confirmation.**

7. **Revenue forward-fill includes beginning-of-sample gaps.** `revM_ff = rev_m.T.ffill().T` — `.ffill()` only propagates forward, so beginning NaN stays NaN. This is correct per the project rules. OK.

8. **Carbon data timing convention**: The project (§3.6 of the reference report) notes that the optimizer at end of year Y only knows CI_{Y-1}. The current code uses `get_carbon_vectors(Y, isins)` in the constraint, meaning it uses *year Y* carbon data when deciding weights at end of year Y. The project brief says "The allocation is based on information available at the end of the year" — this is ambiguous about whether that means the just-completed year Y or the previous year Y-1. The reference report explicitly uses CI_{Y-1}. **Potential look-ahead bias if carbon data for year Y is not yet available at end of Y.**

### Notebook / Reproducibility Issues

9. **No markdown cells.** The script is a flat `.py` file with section comments but no explanatory text for a notebook.

10. **Hard-coded template path.** `export_part1_excel_template()` requires `"Template for Part I-SAAM.xlsx"` which may not be provided. Should be wrapped in try/except.

11. **`warnings.filterwarnings("ignore")`** hides potential numerical issues (near-singular matrices, convergence warnings).

12. **No explicit dependency check.** `cvxpy` import is optional but not clearly communicated. The SLSQP fallback works but convergence is less reliable for large N.

13. **Region code `PAC` is hard-coded.** Should be a clearly documented parameter at the top.

### Missing Plots / Tables

14. **No drawdown subplot.** The reference report pairs every cumulative return plot with a drawdown panel (Figures 2, 3, 6, 8, 10). The current code only produces simple cumulative wealth lines.

15. **No bar charts for CF or WACI.** The reference report uses grouped bar charts with annotations, constraint lines, and firm counts (Figures 4, 5, 7, 9, 11, 12, 13). The code produces simple line plots.

16. **No summary statistics table formatted for display.** Tables 3–9, 13 in the reference report show side-by-side portfolio metrics. The code exports CSVs but doesn't format them as display tables.

17. **No correlation matrix (Table 14).** Not computed.

18. **No robustness analysis (Table 15, Section 4.6).** The reference report re-runs the entire pipeline with L ∈ {48, 60, 72, 84}. Not present in the code.

19. **No data processing summary table (Table 1).** Annual sample sizes at each filtering stage.

20. **No top-10 carbon contributors table with proper formatting** for the report (project §3.1 asks for firm names + ISIN codes).

### Style / Readability Issues

21. **Methodological notes are verbose.** Multi-line comments above functions explain choices but read as AI-generated disclaimers rather than natural student notes.

22. **Some duplicated logic.** `compute_mv_oos_returns` and `compute_oos_returns` are nearly identical; only one is needed.

23. **Variable naming is generally good** but some names like `df_31`, `port_32`, `ret_33` are opaque without context. Adding brief inline comments would help.

24. **Output file naming** mixes `RESULTS_PART1` and `RESULTS_PART2` with string concatenation — would be cleaner with `os.path.join`.

---

## B. Recommended Fixes (Priority Order)

### Must Fix (Critical)

| # | Fix | Effort |
|---|-----|--------|
| 1 | **Replace `estimate_moments` with proper pairwise covariance** | Medium |
| 2 | **Add drawdown + cumulative return dual-panel plots** | Medium |
| 3 | **Add bar-chart CF/WACI plots** matching reference style | Medium |
| 4 | **Compute and display correlation matrix** | Low |
| 5 | **Add formatted summary statistics tables** (display + CSV) | Low |
| 6 | **Add top-10 carbon contributors table** with names | Low |

### Should Fix (Methodological / Presentation)

| # | Fix | Effort |
|---|-----|--------|
| 7 | Add data-processing summary table (Table 1 style) | Low |
| 8 | Guard template export in try/except | Low |
| 9 | Remove duplicate `compute_mv_oos_returns` (use `compute_oos_returns` only) | Low |
| 10 | Clean up double ffill of CO2 panel | Low |
| 11 | Trim verbose methodological notes to concise comments | Low |
| 12 | Add `os.path.join` consistently for output paths | Low |

### Nice to Have (Completeness)

| # | Fix | Effort |
|---|-----|--------|
| 13 | Add robustness analysis (L=48,60,72,84) as optional section | High |
| 14 | Verify carbon data timing (Y vs Y-1) against project brief | Manual |
| 15 | Verify Scope assignment (1 only vs 1+2) | Manual |

---

## C. Revised Code

See `saam_revised.py` — key changes:

1. **Pairwise covariance estimator**: New `estimate_pairwise_covariance()` computes each (i,j) entry using only overlapping valid months, as specified in §2.1 of the project and §3.1 of the reference report.

2. **Enhanced plotting functions**: `plot_cumret_drawdown()` for dual-panel cumulative + drawdown. `plot_cf_bars()` and `plot_waci_bars()` for grouped bar charts with constraint lines and annotations.

3. **Correlation matrix computation** added after all portfolio returns are computed.

4. **Formatted summary tables** using `pd.DataFrame.style` for display and clean CSV export.

5. **Removed duplicate OOS return function**; unified on `compute_oos_returns`.

6. **Cleaned comments** to sound natural rather than AI-generated.

---

## D. Notebook Structure

```
1. Setup & Configuration
   - Markdown: project overview, group info, region, scope
   - Code: imports, configuration constants

2. Data Loading
   - Markdown: brief description of datasets
   - Code: load static, RI monthly, MV monthly, CO2, revenue, annual cap, RF

3. Data Cleaning
   - Markdown: cleaning strategy (missing prices, low prices, stale prices, delistings)
   - Code: price cleaning, return computation, delisting handling
   - Output: data-processing summary table (Table 1 style)

4. Part I — Standard Portfolio Allocation
   4.1 Investment Set Construction
       - Markdown: explain criteria
       - Code: build_investment_set loop, print eligible counts
   4.2 Minimum-Variance Portfolio
       - Markdown: optimization problem, covariance estimation
       - Code: estimate moments, solve min-variance, compute OOS returns
   4.3 Value-Weighted Benchmark
       - Code: compute VW returns
   4.4 Results — Part I
       - Code: summary statistics table
       - Code: cumulative return + drawdown plot (MV vs VW)
       - Code: export Part I template

5. Part II — Carbon Objectives
   5.1 Carbon Metrics (Section 3.1)
       - Markdown: define WACI, CF
       - Code: compute carbon metrics for MV and VW
       - Output: CF and WACI comparison plots (bar charts)
       - Output: top-10 carbon contributors table
   5.2 Min-Variance with 50% Carbon Cap (Section 3.2)
       - Markdown: optimization problem
       - Code: solve constrained MV, compute returns
       - Output: cumulative return plot, CF trajectory, stats table
   5.3 Tracking-Error Minimization (Section 3.3)
       - Markdown: optimization problem
       - Code: solve TE portfolio, compute returns
       - Output: cumulative return plot, CF trajectory, stats table
   5.4 Net-Zero Portfolio (Section 4.1)
       - Markdown: dynamic constraint
       - Code: solve NZ portfolio, compute returns
       - Output: cumulative return plot, NZ path plot, stats table

6. Comparison & Synthesis
   - Code: 5-portfolio comparison table
   - Code: all-strategies cumulative return plot
   - Code: correlation matrix
   - Code: all-strategies CF and WACI evolution

7. Export
   - Code: save all CSVs, weight files
```

---

## E. Missing Plots

### Added in revised code:
1. **Cumulative return + drawdown** dual-panel plots for each portfolio comparison
2. **Grouped bar charts** for CF comparison (MV vs VW, with constraint lines)
3. **Grouped bar charts** for WACI comparison
4. **NZ carbon path plot** with 10% decline trajectory
5. **5-strategy overlay** cumulative return plot
6. **Correlation heatmap**

### Still required (need data or manual work):
7. Data-processing summary table (Table 1) — needs intermediate filtering counts
8. Robustness table (Table 15) — needs re-running pipeline with different L values
9. Weight composition analysis plots (sector tilts, concentration)

---

## F. Assumptions and Open Points

1. **Region = PAC (Pacific)** is assumed correct per group assignment. Must verify.
2. **Scope = Scope 1 only** is assumed correct. If the assignment is Scope 1+2, the CO2 loading must add both files.
3. **Carbon data timing**: The code uses year Y carbon data for the optimization at end of year Y. The reference report uses Y-1 (look-ahead safe). Verify against project brief — the brief says "information available at the end of the year", which could mean either. Using Y-1 is the safer choice.
4. **Data files**: The code expects specific filenames (`DS_CO2_SCOPE_1_Y_2025.xlsx`, etc.). These must match what the student has.
5. **Risk-free rate**: Loaded from `Risk_Free_Rate_2025.xlsx`. If unavailable, rf=0 is used. The Sharpe ratio formula uses excess returns, so this matters.
6. **`cvxpy` availability**: The SLSQP fallback works but may give less accurate solutions for large problems. Installing cvxpy is recommended.
7. **Revenue division by 1000** is done (`rev_m = rev_raw / 1000.0`), matching the project requirement.
8. **Number of OOS months = 144** assumes price data runs through Dec 2025. If data stops at Dec 2024, this needs adjustment (132 months).
9. **The reference report PDF (SAAM_ENG_34_amended.pdf) is from a different group** (North America, 2014–2024). Plot styles should be emulated but numbers will differ.
10. **Robustness analysis** (Section 4.6 of reference) is not implemented. This is a significant gap for completeness but requires re-running the entire pipeline 3 extra times.
