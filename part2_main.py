"""
SAAM Project — Part 2: Portfolio Allocation with Carbon Emission Reduction
Sections 3.1 | 3.2 | 3.3 | 3.4 | 4.1 | 4.2

Group: Pacific Area | Scope 1

How this script works
─────────────────────
• All helper functions (data loading, cleaning, moments, OOS returns) are
  re-used directly from saam_part1_complete_v2.py via importlib — zero
  code duplication.
• The raw-data pipeline is re-run (same files, same cleaning logic) so that
  ri_returns, ri_prices, and covariance matrices are available in memory.
  (Part 1 did not serialise sigma matrices or ri_returns to disk.)
• Part 1 portfolio compositions and monthly returns are loaded from
  resultsPart1/.
• Carbon data (CO2, Revenue, Annual MV) is loaded and merged with the Part 1
  investment set; firms without carbon data at year-end Y are excluded for
  that year's optimisation (per project spec §2.1).
• All optimisation problems (§3.2, §3.3, §4.1) are solved with CVXPY
  (OSQP solver, SCS fallback).
• Outputs are saved to ResultsPart2/.

Dependencies: cvxpy, pandas, numpy, openpyxl
Run: python part2_main.py
"""

import os
import sys
import importlib.util
import warnings

import cvxpy as cp
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION  (must match Part 1)
# ─────────────────────────────────────────────────────────────────────────────
RESULTS_PART1  = "resultsPart1"       # exact folder name used in Part 1
RESULTS_PART2  = "ResultsPart2"
os.makedirs(RESULTS_PART2, exist_ok=True)

REGION_CODE    = "PAC"
START_YEAR_OOS = 2014
END_YEAR_OOS   = 2025
DECISION_YEARS = list(range(2013, 2025))   # year_end values (decision at end-Y, active Y+1)

Y0    = 2013
THETA = 0.10    # net-zero annual reduction rate

# Part 1 numerical settings (keep identical for moment estimation)
WINDOW_YEARS   = 10
MIN_OBS_MONTHS = 36
STALE_THRESHOLD = 0.50
LOW_PRICE_THRESHOLD = 0.5
RIDGE = 1e-8

# ─────────────────────────────────────────────────────────────────────────────
# IMPORT ALL HELPER FUNCTIONS FROM PART 1
# ─────────────────────────────────────────────────────────────────────────────
_spec = importlib.util.spec_from_file_location("part1", "saam_part1_complete_v2.py")
_p1   = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_p1)

load_datastream_wide      = _p1.load_datastream_wide
parse_monthly_columns     = _p1.parse_monthly_columns
extract_delist_date       = _p1.extract_delist_date
clean_monthly_ri_prices   = _p1.clean_monthly_ri_prices
apply_delisting_to_returns = _p1.apply_delisting_to_returns
year_end_col              = _p1.year_end_col
window_cols               = _p1.window_cols
build_investment_set      = _p1.build_investment_set
estimate_moments          = _p1.estimate_moments
perf_stats                = _p1.perf_stats   # geometric: (1+mu_m)^12 - 1

print("Part 1 functions imported successfully.")

# ─────────────────────────────────────────────────────────────────────────────
# 1.  RE-RUN PART 1 DATA PIPELINE  (to recover ri_returns, ri_prices, dates)
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("Step 1 — Loading and cleaning monthly RI / MV data (Part 1 pipeline)…")

static = pd.read_excel("Static_2025.xlsx", engine="openpyxl")
static.columns = ["ISIN", "NAME", "Country", "Region"]
static["ISIN"] = static["ISIN"].astype(str).str.strip()
pac = static[static["Region"] == REGION_CODE].set_index("ISIN")
pac_isins = set(pac.index)
print(f"  Pacific firms in Static: {len(pac_isins)}")

# Annual RI — for delist dates only
ri_y = load_datastream_wide("DS_RI_T_USD_Y_2025.xlsx")
ri_y = ri_y[ri_y.index.isin(pac_isins)].copy()
delist_dates = {isin: extract_delist_date(ri_y.at[isin, "NAME"]) for isin in ri_y.index}

# Monthly RI
ri_m_raw = load_datastream_wide("DS_RI_T_USD_M_2025.xlsx")
keep_cols, parsed_dates = parse_monthly_columns(list(ri_m_raw.columns))
ri_m = ri_m_raw[["NAME"] + keep_cols].copy()
ri_m.columns = ["NAME"] + parsed_dates
ri_m = ri_m[ri_m.index.isin(pac_isins)].copy()
parsed_dates = [d for d in parsed_dates if d <= pd.Timestamp("2025-12-31")]
ri_m = ri_m[["NAME"] + parsed_dates]

# Monthly MV
mv_m_raw = load_datastream_wide("DS_MV_T_USD_M_2025.xlsx")
keep_cols_mv, _ = parse_monthly_columns(list(mv_m_raw.columns))
mv_m = mv_m_raw[["NAME"] + keep_cols_mv].copy()
mv_m.columns = ["NAME"] + _
mv_m = mv_m[mv_m.index.isin(pac_isins)].copy()
mv_m = mv_m[["NAME"] + parsed_dates].copy()
mv_m[parsed_dates] = mv_m[parsed_dates].apply(pd.to_numeric, errors="coerce")

# Common ISINs
common_isins = pac_isins & set(ri_m.index) & set(mv_m.index) & set(ri_y.index)
ri_m  = ri_m.loc[list(common_isins)].copy()
mv_m  = mv_m.loc[list(common_isins)].copy()
print(f"  Common ISINs: {len(common_isins)}")

# Clean prices & compute returns
ri_prices = clean_monthly_ri_prices(ri_m[parsed_dates], parsed_dates,
                                    LOW_PRICE_THRESHOLD, True)
all_missing = ri_prices.isna().all(axis=1)
if all_missing.any():
    # Drop firms with all-missing RI prices from both price and MV panels
    ri_prices = ri_prices.loc[~all_missing].copy()
    mv_m      = mv_m.loc[~all_missing].copy()

ri_returns = ri_prices.pct_change(axis=1)
delist_common = {k: v for k, v in delist_dates.items() if k in ri_returns.index}
ri_returns = apply_delisting_to_returns(ri_returns, delist_common, parsed_dates)
print("  Monthly returns ready.")

# ─────────────────────────────────────────────────────────────────────────────
# 2.  LOAD ANNUAL CARBON DATA
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("Step 2 — Loading annual carbon / fundamental data…")

def load_annual_panel(filepath: str) -> pd.DataFrame:
    """Read Datastream annual panel → ISIN-indexed DataFrame with int year columns."""
    df = pd.read_excel(filepath, engine="openpyxl")
    df.columns = ["NAME", "ISIN"] + list(df.columns[2:])
    df = df[~df["NAME"].astype(str).str.startswith("$$ER", na=False)]
    df = df.dropna(subset=["ISIN"])
    df["ISIN"] = df["ISIN"].astype(str).str.strip()
    df = df[df["ISIN"] != ""].set_index("ISIN").drop(columns=["NAME"], errors="ignore")
    df.columns = df.columns.astype(int)
    df = df.apply(pd.to_numeric, errors="coerce")
    return df

def ffill_annual(df: pd.DataFrame) -> pd.DataFrame:
    """Forward-fill annual gaps row-by-row (middle/end only; leading NaN stays NaN)."""
    return df.T.ffill().T

co2_raw  = load_annual_panel("DS_CO2_SCOPE_1_Y_2025.xlsx")   # tonnes
rev_raw  = load_annual_panel("DS_REV_Y_2025.xlsx")          # thousands USD → /1000 = M USD
cap_ann_raw = load_annual_panel("DS_MV_T_USD_Y_2025.xlsx")  # million USD (year-end)

rev_m    = rev_raw / 1_000.0     # convert to million USD

co2_ff   = ffill_annual(co2_raw)
revM_ff  = ffill_annual(rev_m)
capA_ff  = ffill_annual(cap_ann_raw)

# Firm name lookup for top-10 table
_tmp = pd.read_excel("DS_CO2_SCOPE_1_Y_2025.xlsx", engine="openpyxl")
_tmp.columns = ["NAME", "ISIN"] + list(_tmp.columns[2:])
_tmp = _tmp[~_tmp["NAME"].astype(str).str.startswith("$$ER", na=False)].dropna(subset=["ISIN"])
_tmp["ISIN"] = _tmp["ISIN"].astype(str).str.strip()
firm_names = _tmp.set_index("ISIN")["NAME"]

print("  Annual carbon data ready.")

# ─────────────────────────────────────────────────────────────────────────────
# 3.  LOAD PART 1 PORTFOLIO RESULTS
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("Step 3 — Loading Part 1 portfolio compositions and returns…")

# Compositions: Year = year_end + 1 (= investment year)
comp_df = pd.read_csv(f"{RESULTS_PART1}/part1_portfolio_compositions.csv")
comp_df["Year"] = comp_df["Year"].astype(int)
comp_df["ISIN"] = comp_df["ISIN"].astype(str).str.strip()

# Build dict: year_end → Series(weight, index=ISIN)
mv_weights_p1 = {}
for yr, grp in comp_df.groupby("Year"):
    year_end = yr - 1   # year_end = investment_year - 1
    mv_weights_p1[year_end] = grp.set_index("ISIN")["Weight"]

# Monthly returns
ret_df = pd.read_csv(f"{RESULTS_PART1}/part1_results.csv", parse_dates=["Date"])
ret_df = ret_df.set_index("Date").sort_index()
mv_port_ret = ret_df["MV_Return"]
vw_port_ret = ret_df["VW_Return"]
print(f"  Part 1 returns loaded: {len(mv_port_ret)} monthly obs.")

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS  (carbon metrics + portfolio return computation)
# ─────────────────────────────────────────────────────────────────────────────

def get_carbon_vectors(Y: int, isins: list):
    """Return (co2, rev_m, cap_m) numpy arrays for year Y and given ISINs."""
    e = co2_ff[Y].reindex(isins).fillna(0.0).values  if Y in co2_ff.columns  else np.zeros(len(isins))
    r = revM_ff[Y].reindex(isins).fillna(np.nan).values if Y in revM_ff.columns else np.full(len(isins), np.nan)
    c = capA_ff[Y].reindex(isins).fillna(0.0).values if Y in capA_ff.columns else np.zeros(len(isins))
    return e, r, c

def carbon_intensity_vec(e, r):
    with np.errstate(invalid="ignore", divide="ignore"):
        return np.where(r > 0, e / r, np.nan)

def waci_metric(w, ci):
    valid = ~np.isnan(ci)
    if not valid.any(): return np.nan
    w2 = np.where(valid, w, 0.0); s = w2.sum()
    return float((w2 / s) @ np.where(valid, ci, 0.0)) if s > 0 else np.nan

def cf_metric(w, e, c):
    """CF = sum_i alpha_i * E_i / Cap_i  [tCO2 / M USD invested]"""
    with np.errstate(invalid="ignore", divide="ignore"):
        ratio = np.where(c > 0, e / c, 0.0)
    return float(w @ ratio)

def cf_vw_metric(e, c):
    """CF_vw = sum(E_i) / sum(Cap_i)"""
    total = c.sum()
    return float(e.sum() / total) if total > 0 else np.nan

def carbon_eligible_isins(base_isins: list, Y: int) -> list:
    """Subset of base_isins that have both CO2 and Revenue data at year Y."""
    e_ok  = co2_ff[Y].reindex(base_isins).notna() if Y in co2_ff.columns  else pd.Series(False, index=base_isins)
    r_ok  = revM_ff[Y].reindex(base_isins).notna() if Y in revM_ff.columns else pd.Series(False, index=base_isins)
    c_ok  = capA_ff[Y].reindex(base_isins).notna() if Y in capA_ff.columns else pd.Series(False, index=base_isins)
    mask  = e_ok & r_ok & c_ok
    return [i for i in base_isins if mask.get(i, False)]

def e_over_c_vec(e, c):
    with np.errstate(invalid="ignore", divide="ignore"):
        return np.where(c > 0, e / c, 0.0)

def compute_oos_returns(portfolios_dict: dict, ri_ret: pd.DataFrame,
                        dates: list, start_yr: int, end_yr: int) -> pd.Series:
    """
    OOS portfolio returns with intra-year weight drift.
    portfolios_dict: {year_end: {"isins": list, "weights": np.ndarray}}
    Mirrors Part 1 compute_mv_oos_returns exactly.
    """
    out_r, out_d = [], []
    for year_end in sorted(portfolios_dict.keys()):
        invest_year = year_end + 1
        if invest_year < start_yr or invest_year > end_yr:
            continue
        year_months = [d for d in dates if d.year == invest_year]
        if not year_months:
            continue
        isins = portfolios_dict[year_end]["isins"]
        w     = portfolios_dict[year_end]["weights"].copy()
        for d in year_months:
            r_i = ri_ret.loc[isins, d].fillna(0.0).to_numpy() if d in ri_ret.columns else np.zeros(len(isins))
            r_p = float(w @ r_i)
            out_r.append(r_p); out_d.append(d)
            denom = 1.0 + r_p
            if denom != 0:
                w = w * (1.0 + r_i) / denom
    return pd.Series(out_r, index=pd.to_datetime(out_d)).sort_index()

def solve_cvxpy(Sigma_arr, e_c, cf_target, w_ref=None, mode="mv"):
    """
    mode='mv' : minimise alpha'Σalpha  (Section 3.2)
    mode='te' : minimise (alpha-w_ref)'Σ(alpha-w_ref)  (Sections 3.3 / 4.1)
    Returns weight vector (length N), normalised and clipped ≥ 0.
    """
    N = Sigma_arr.shape[0]
    Sp = cp.psd_wrap(Sigma_arr)
    alpha = cp.Variable(N)
    if mode == "mv":
        obj = cp.Minimize(cp.quad_form(alpha, Sp))
    else:
        diff = alpha - w_ref
        obj  = cp.Minimize(cp.quad_form(diff, Sp))
    constraints = [cp.sum(alpha) == 1, alpha >= 0,
                   alpha @ e_c <= cf_target]
    prob = cp.Problem(obj, constraints)
    for solver in (cp.OSQP, cp.SCS):
        try:
            prob.solve(solver=solver,
                       **({"eps_abs": 1e-8, "eps_rel": 1e-8, "max_iter": 20_000}
                          if solver == cp.OSQP else {}))
        except Exception:
            continue
        if alpha.value is not None and prob.status in ("optimal", "optimal_inaccurate"):
            w = np.maximum(alpha.value, 0.0)
            s = w.sum()
            return w / s if s > 0 else np.ones(N) / N
    return None   # both solvers failed

# ─────────────────────────────────────────────────────────────────────────────
# 4.  BUILD CARBON-ELIGIBLE INVESTMENT SETS + SIGMA MATRICES
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("Step 4 — Building carbon-eligible investment sets and Σ matrices…")

inv_sets_p2   = {}   # year_end → list of carbon-eligible ISINs
sigma_p2      = {}   # year_end → np.ndarray
vw_weights_p2 = {}   # year_end → np.ndarray (VW weights in the eligible set)

for Y in DECISION_YEARS:
    # Part 1 investment set (from compositions CSV)
    if Y not in mv_weights_p1:
        print(f"  Y={Y}: no Part 1 weights found — skipping.")
        continue
    base_isins = list(mv_weights_p1[Y].index)

    # Intersect with carbon data availability
    eligible = carbon_eligible_isins(base_isins, Y)
    if len(eligible) < 2:
        print(f"  Y={Y}: only {len(eligible)} carbon-eligible firms — skipping.")
        continue

    # Re-estimate Sigma on the carbon-eligible subset (same 10-yr window)
    cols = window_cols(parsed_dates, Y, WINDOW_YEARS)
    _, Sigma = estimate_moments(ri_returns, eligible, cols)   # reuses Part 1 function
    sigma_p2[Y] = Sigma

    # VW weights at year-end Y (using annual market caps)
    cap_y = capA_ff[Y].reindex(eligible).fillna(0.0).values if Y in capA_ff.columns else np.zeros(len(eligible))
    total_cap = cap_y.sum()
    w_vw = cap_y / total_cap if total_cap > 0 else np.ones(len(eligible)) / len(eligible)
    vw_weights_p2[Y] = w_vw

    inv_sets_p2[Y] = eligible
    print(f"  Y={Y}: eligible (carbon) = {len(eligible)} / {len(base_isins)}")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3.1 — CARBON METRICS FOR EXISTING PORTFOLIOS
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("Section 3.1 — Carbon intensity, WACI, Carbon Footprint")

rows_31 = {}
for Y in DECISION_YEARS:
    if Y not in inv_sets_p2:
        continue
    isins = inv_sets_p2[Y]
    e, r, c = get_carbon_vectors(Y, isins)
    ci = carbon_intensity_vec(e, r)

    # MV weights (from Part 1, restricted to carbon-eligible firms)
    w_mv = mv_weights_p1[Y].reindex(isins).fillna(0.0).values
    s = w_mv.sum(); w_mv = w_mv / s if s > 0 else np.ones(len(isins)) / len(isins)

    w_vw = vw_weights_p2[Y]

    rows_31[Y] = {
        "WACI_MV": waci_metric(w_mv, ci),
        "WACI_VW": waci_metric(w_vw, ci),
        "CF_MV":   cf_metric(w_mv, e, c),
        "CF_VW":   cf_vw_metric(e, c),
    }

df_31 = pd.DataFrame(rows_31).T
df_31.index.name = "Year"
df_31.to_csv(f"{RESULTS_PART2}/carbon_metrics_mv_vw.csv")
print(df_31.round(4).to_string())

# Top-10 firms by average CI
ci_acc = {}
for Y in DECISION_YEARS:
    if Y not in inv_sets_p2:
        continue
    isins = inv_sets_p2[Y]
    e, r, _ = get_carbon_vectors(Y, isins)
    for isin, ci_val in zip(isins, carbon_intensity_vec(e, r)):
        if not np.isnan(ci_val):
            ci_acc.setdefault(isin, []).append(ci_val)

mean_ci = pd.Series({k: np.mean(v) for k, v in ci_acc.items()}).sort_values(ascending=False).head(10)
top10 = pd.DataFrame({"ISIN": mean_ci.index,
                      "Name": mean_ci.index.map(firm_names),
                      "Avg_CI": mean_ci.values})
top10.to_csv(f"{RESULTS_PART2}/top10_carbon_intensity.csv", index=False)
print("\nTop 10 firms by average carbon intensity:")
print(top10.to_string(index=False))

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3.2 — LONG-ONLY MV WITH 50 % CF REDUCTION
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("Section 3.2 — MV portfolio with 50 % CF constraint")

port_32 = {}   # year_end → {"isins": [...], "weights": np.array}
cf_32, waci_32 = {}, {}

for Y in DECISION_YEARS:
    if Y not in inv_sets_p2:
        continue
    isins   = inv_sets_p2[Y]
    Sigma   = sigma_p2[Y]
    e, r, c = get_carbon_vectors(Y, isins)
    ec      = e_over_c_vec(e, c)
    cf_mv_Y = df_31.loc[Y, "CF_MV"]

    if np.isnan(cf_mv_Y) or cf_mv_Y <= 0:
        print(f"  Y={Y}: CF_MV invalid ({cf_mv_Y:.4f}), skipping 3.2 for this year.")
        continue

    cf_target = 0.5 * cf_mv_Y
    w = solve_cvxpy(Sigma, ec, cf_target, mode="mv")

    if w is None:
        print(f"  Y={Y}: solver failed — falling back to Part 1 MV weights.")
        w = mv_weights_p1[Y].reindex(isins).fillna(0.0).values
        w = np.maximum(w, 0.0); w /= w.sum() if w.sum() > 0 else 1.0

    port_32[Y]  = {"isins": isins, "weights": w}
    cf_32[Y]    = cf_metric(w, e, c)
    waci_32[Y]  = waci_metric(w, carbon_intensity_vec(e, r))
    print(f"  Y={Y}  CF_target={cf_target:.3f}  achieved={cf_32[Y]:.3f}  "
          f"non-zero={int(np.sum(w > 1e-4))}/{len(isins)}")

pd.concat([pd.DataFrame({"Year": Y, "ISIN": v["isins"], "Weight": v["weights"]})
           for Y, v in port_32.items()]).to_csv(
    f"{RESULTS_PART2}/weights_32_mv_carbon05.csv", index=False)

ret_32 = compute_oos_returns(port_32, ri_returns, parsed_dates, START_YEAR_OOS, END_YEAR_OOS)
ret_32.to_csv(f"{RESULTS_PART2}/returns_32_mv_carbon05.csv", header=["Return"])

stats_32 = pd.DataFrame({
    "P_mv_oos":       perf_stats(mv_port_ret),
    "P_mv_oos(0.5)":  perf_stats(ret_32),
})
stats_32.to_csv(f"{RESULTS_PART2}/stats_32_comparison.csv")
print("\n  Section 3.2 performance comparison:")
print(stats_32.round(4).to_string())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3.3 — TRACKING ERROR MIN WITH 50 % CF REDUCTION
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("Section 3.3 — Tracking-error minimisation with 50 % CF constraint")

port_33 = {}
cf_33, waci_33 = {}, {}

for Y in DECISION_YEARS:
    if Y not in inv_sets_p2:
        continue
    isins   = inv_sets_p2[Y]
    Sigma   = sigma_p2[Y]
    e, r, c = get_carbon_vectors(Y, isins)
    ec      = e_over_c_vec(e, c)
    w_vw    = vw_weights_p2[Y]
    cf_vw_Y = df_31.loc[Y, "CF_VW"]

    if np.isnan(cf_vw_Y) or cf_vw_Y <= 0:
        print(f"  Y={Y}: CF_VW invalid, skipping 3.3 for this year.")
        continue

    cf_target = 0.5 * cf_vw_Y
    w = solve_cvxpy(Sigma, ec, cf_target, w_ref=w_vw, mode="te")

    if w is None:
        print(f"  Y={Y}: solver failed — falling back to VW weights.")
        w = w_vw.copy()

    port_33[Y]  = {"isins": isins, "weights": w}
    cf_33[Y]    = cf_metric(w, e, c)
    waci_33[Y]  = waci_metric(w, carbon_intensity_vec(e, r))
    print(f"  Y={Y}  CF_target={cf_target:.3f}  achieved={cf_33[Y]:.3f}  "
          f"non-zero={int(np.sum(w > 1e-4))}/{len(isins)}")

pd.concat([pd.DataFrame({"Year": Y, "ISIN": v["isins"], "Weight": v["weights"]})
           for Y, v in port_33.items()]).to_csv(
    f"{RESULTS_PART2}/weights_33_te_carbon05.csv", index=False)

ret_33 = compute_oos_returns(port_33, ri_returns, parsed_dates, START_YEAR_OOS, END_YEAR_OOS)
ret_33.to_csv(f"{RESULTS_PART2}/returns_33_te_carbon05.csv", header=["Return"])

stats_33 = pd.DataFrame({
    "P_vw_oos":       perf_stats(vw_port_ret),
    "P_vw_oos(0.5)":  perf_stats(ret_33),
})
stats_33.to_csv(f"{RESULTS_PART2}/stats_33_comparison.csv")
print("\n  Section 3.3 performance comparison:")
print(stats_33.round(4).to_string())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4.1 — NET ZERO PORTFOLIO
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print(f"Section 4.1 — Net-zero portfolio (θ={THETA:.0%}/yr from Y0={Y0})")

# Reference CF_vw at Y0=2013
isins_y0 = inv_sets_p2.get(Y0, [])
if isins_y0:
    e0, _, c0 = get_carbon_vectors(Y0, isins_y0)
    cf_vw_y0  = cf_vw_metric(e0, c0)
else:
    cf_vw_y0 = df_31.loc[Y0, "CF_VW"] if Y0 in df_31.index else np.nan
print(f"  CF_vw reference (Y0={Y0}) = {cf_vw_y0:.4f} tCO2/M USD")

port_41 = {}
cf_41, waci_41 = {}, {}

for Y in DECISION_YEARS:
    if Y not in inv_sets_p2:
        continue
    isins   = inv_sets_p2[Y]
    Sigma   = sigma_p2[Y]
    e, r, c = get_carbon_vectors(Y, isins)
    ec      = e_over_c_vec(e, c)
    w_vw    = vw_weights_p2[Y]

    # Tightening constraint: (1-θ)^(Y-Y0+1) * CF_vw(Y0)
    cf_target = ((1.0 - THETA) ** (Y - Y0 + 1)) * cf_vw_y0
    w = solve_cvxpy(Sigma, ec, cf_target, w_ref=w_vw, mode="te")

    if w is None:
        print(f"  Y={Y}: solver failed — falling back to VW weights.")
        w = w_vw.copy()

    port_41[Y]  = {"isins": isins, "weights": w}
    cf_41[Y]    = cf_metric(w, e, c)
    waci_41[Y]  = waci_metric(w, carbon_intensity_vec(e, r))
    print(f"  Y={Y}  CF_target={cf_target:.4f}  achieved={cf_41[Y]:.4f}  "
          f"non-zero={int(np.sum(w > 1e-4))}/{len(isins)}")

pd.concat([pd.DataFrame({"Year": Y, "ISIN": v["isins"], "Weight": v["weights"]})
           for Y, v in port_41.items()]).to_csv(
    f"{RESULTS_PART2}/weights_41_netzero.csv", index=False)

ret_41 = compute_oos_returns(port_41, ri_returns, parsed_dates, START_YEAR_OOS, END_YEAR_OOS)
ret_41.to_csv(f"{RESULTS_PART2}/returns_41_netzero.csv", header=["Return"])

stats_41 = pd.DataFrame({
    "P_vw_oos":       perf_stats(vw_port_ret),
    "P_vw_oos(0.5)":  perf_stats(ret_33),
    "P_vw_oos(NZ)":   perf_stats(ret_41),
})
stats_41.to_csv(f"{RESULTS_PART2}/stats_41_comparison.csv")
print("\n  Section 4.2 three-way performance comparison:")
print(stats_41.round(4).to_string())

# ─────────────────────────────────────────────────────────────────────────────
# MASTER SUMMARY FILES
# ─────────────────────────────────────────────────────────────────────────────
all_stats = pd.DataFrame({
    "P_mv_oos":       perf_stats(mv_port_ret),
    "P_mv_oos(0.5)":  perf_stats(ret_32),
    "P_vw_oos":       perf_stats(vw_port_ret),
    "P_vw_oos(0.5)":  perf_stats(ret_33),
    "P_vw_oos(NZ)":   perf_stats(ret_41),
})
all_stats.to_csv(f"{RESULTS_PART2}/all_portfolio_stats.csv")

all_carbon = pd.DataFrame({
    "CF_MV":      df_31["CF_MV"],
    "CF_VW":      df_31["CF_VW"],
    "CF_MV_05":   pd.Series(cf_32),
    "CF_TE_05":   pd.Series(cf_33),
    "CF_NZ":      pd.Series(cf_41),
    "WACI_MV":    df_31["WACI_MV"],
    "WACI_VW":    df_31["WACI_VW"],
    "WACI_MV_05": pd.Series(waci_32),
    "WACI_TE_05": pd.Series(waci_33),
    "WACI_NZ":    pd.Series(waci_41),
})
all_carbon.index.name = "Year"
all_carbon.to_csv(f"{RESULTS_PART2}/all_carbon_metrics.csv")

# NZ target path (for plotting)
nz_path = pd.Series(
    {Y: ((1.0 - THETA) ** (Y - Y0 + 1)) * cf_vw_y0 for Y in DECISION_YEARS},
    name="NZ_target"
)
nz_path.index.name = "Year"
nz_path.to_csv(f"{RESULTS_PART2}/nz_target_path.csv")

print("=" * 60)
print(f"All Part 2 results saved to '{RESULTS_PART2}/'")
print("Files generated:")
for fn in sorted(os.listdir(RESULTS_PART2)):
    print(f"  {fn}")
print("=" * 60)
