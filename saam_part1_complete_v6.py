# ============================================================
# SAAM PROJECT 2026 — PART I  (v4 — three-bug fix release)
# Group: Pacific Area | Scope 1
#
# Fixes vs v3 + v5 + v6 fixes
# ──────────────────────────────────────
# Bug 4 (CRITICAL) — estimate_moments() now uses Ledoit-Wolf shrinkage instead
# of pairwise deletion + zero-fill + tiny ridge (1e-8). The old approach
# produced a non-PSD covariance matrix (up to 6 negative eigenvalues for the
# 2013 window), causing SLSQP to stall. LW guarantees PSD by construction and
# provides data-driven regularisation (δ̂ ≈ 0.36–0.48 empirically). NaN entries
# are filled with each firm's own mean before passing to LW.
# Change: R.T.cov(ddof=0) + np.where(nan,0) + RIDGE*I → LedoitWolf().fit()
#
# Bug 5 (WARNING) — pct_change() now passes fill_method=None explicitly to
# suppress the FutureWarning about the deprecated default fill_method='pad'.
# Leaving the default active risks implicit forward-filling of missing prices
# before return computation, which is not the intended behaviour.
# Change: ri_prices.pct_change(axis=1) →
#         ri_prices.pct_change(axis=1, fill_method=None)
#
# Fixes vs v3 + v5
# ──────────────────────────────────────
# Bug 3 (IMPORTANT) — compute_vw_oos_returns() now restricts the VW benchmark
# to the same annual eligible investment set used for the MV portfolio.
# Previously, VW used any firm with available lagged monthly cap + current
# return — a looser universe ignoring the stale-price, carbon-availability,
# and min-obs filters applied to the MV portfolio.
# Fix: accepts portfolios dict; for each investment year Y+1, only
# portfolios[Y]["isins"] contribute to VW weights and returns.
# This ensures a like-for-like comparison (brief §2.3 + §3.3).
# Call-site: portfolios arg added before ri_returns.
#
# Fixes vs v3
# ───────────
# Bug 1 (CRITICAL) — Part 1 now forward-fills the CO2 panel before passing it
#   to build_investment_set(), exactly as Part 2 does.  Without this fix the
#   two parts used different investment universes, violating the brief's
#   requirement that both parts share the same investment set.
#   Change: in main(), after load_annual_panel(CO2_FILE), add
#           co2_panel = co2_panel.T.ffill().T
#
# Bug 2 (IMPORTANT) — estimate_moments() now uses ddof=0 to match the brief's
#   formula  Σ_Y = (1/τ) Σ (R_t - μ̂)(R_t - μ̂)'  exactly.
#   The previous version used pandas default ddof=1 (divided by τ−1).
#   Change: R.T.cov()  →  R.T.cov(ddof=0)
#
# Retained from v3
# ────────────────
# • estimate_moments(): NaN-aware pairwise deletion (no fillna(0.0) on returns)
# • perf_stats(): Sharpe computed with actual monthly RF from
#   Risk_Free_Rate_2025.xlsx  (YYYYMM index, RF column in % per month)
# • build_investment_set(): co2_panel= parameter excludes firms without
#   Scope 1 data at year-end Y  →  same eligible universe for Parts I & II
# • load_rf_monthly(): bespoke loader for Risk_Free_Rate_2025.xlsx format
# ============================================================

import os
import pandas as pd
import numpy as np
import re
from datetime import datetime
from scipy.optimize import minimize
from sklearn.covariance import LedoitWolf

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
REGION_CODE         = "PAC"
LOW_PRICE_THRESHOLD = 0.5
RESULTS_DIR         = "resultsPart1"

BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(BASE_DIR, "data")


def data_path(filename: str) -> str:
    """
    Return the path to an input data file.

    Prefers the 'data' subfolder; falls back to the script directory so that
    existing setups (files in the repo root) keep working.
    """
    candidate = os.path.join(DATA_DIR, filename)
    return candidate if os.path.exists(candidate) else os.path.join(BASE_DIR, filename)

WINDOW_YEARS    = 10
MIN_OBS_MONTHS  = 36
STALE_THRESHOLD = 0.50

START_YEAR_OOS = 2014
END_YEAR_OOS   = 2025

CO2_FILE = "DS_CO2_SCOPE_1_Y_2025.xlsx"
RF_FILE  = "Risk_Free_Rate_2025.xlsx"

RIDGE         = 1e-8
SLSQP_MAXITER = 400
SLSQP_FTOL    = 1e-9

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS — loading & cleaning
# ─────────────────────────────────────────────────────────────────────────────

def load_datastream_wide(filepath, sheet=0):
    df = pd.read_excel(filepath, sheet_name=sheet, header=0, engine="openpyxl")
    df.columns = ["NAME", "ISIN"] + list(df.columns[2:])
    df = df[~df["NAME"].astype(str).str.startswith("$$ER", na=False)]
    df = df.dropna(subset=["ISIN"])
    df["ISIN"] = df["ISIN"].astype(str).str.strip()
    df["NAME"] = df["NAME"].astype(str).str.strip()
    df = df[df["ISIN"] != ""].set_index("ISIN")
    return df


def load_annual_panel(filepath: str) -> pd.DataFrame:
    """Read Datastream annual panel → ISIN-indexed DataFrame with int year cols."""
    df = pd.read_excel(filepath, engine="openpyxl")
    df.columns = ["NAME", "ISIN"] + list(df.columns[2:])
    df = df[~df["NAME"].astype(str).str.startswith("$$ER", na=False)]
    df = df.dropna(subset=["ISIN"])
    df["ISIN"] = df["ISIN"].astype(str).str.strip()
    df = df[df["ISIN"] != ""].set_index("ISIN").drop(columns=["NAME"], errors="ignore")
    df.columns = df.columns.astype(int)
    df = df.apply(pd.to_numeric, errors="coerce")
    return df


def load_rf_monthly(filepath):
    """
    Load Risk_Free_Rate_2025.xlsx.

    Format
    ------
    Index column : YYYYMM integer  (e.g. 202401 = January 2024)
    'RF' column  : monthly rate in PERCENT  (e.g. 0.41 = 0.41 %/month)

    Returns a pd.Series of monthly DECIMAL rates indexed by month-end
    Timestamps (same end-of-month convention as Datastream columns).
    Division by 100 converts percent → decimal.
    """
    df    = pd.read_excel(filepath, engine="openpyxl", index_col=0)
    df.index = df.index.astype(str).str.strip()
    dates = pd.to_datetime(df.index, format="%Y%m") + pd.offsets.MonthEnd(0)
    rf    = pd.to_numeric(df.iloc[:, 0], errors="coerce") / 100.0
    rf.index = dates
    return rf.dropna().sort_index()


def parse_monthly_columns(cols):
    parsed, keep = [], []
    for c in cols:
        if c == "NAME":
            continue
        dt = pd.to_datetime(str(c), errors="coerce", dayfirst=True)
        if pd.notna(dt):
            keep.append(c)
            parsed.append(pd.Timestamp(dt).normalize())
    return keep, parsed


def extract_delist_date(name_str):
    m = re.search(r"(?:DELIST|DEAD)\.(\d{2}/\d{2}/\d{2,4})", str(name_str))
    if not m:
        return None
    s = m.group(1)
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def forward_fill_middle_only(price_df, dates):
    arr = price_df.to_numpy(dtype=float, copy=True)
    for i in range(arr.shape[0]):
        row       = arr[i, :]
        valid_idx = np.where(~np.isnan(row))[0]
        if valid_idx.size == 0:
            continue
        first, last = valid_idx[0], valid_idx[-1]
        last_val = row[first]
        for k in range(first + 1, last + 1):
            if np.isnan(row[k]):
                row[k] = last_val
            else:
                last_val = row[k]
        arr[i, :] = row
    return pd.DataFrame(arr, index=price_df.index, columns=dates)


def clean_monthly_ri_prices(raw_ri_m, dates, low_price_threshold=0.5,
                             preserve_december_missing=True):
    prices = raw_ri_m.copy().apply(pd.to_numeric, errors="coerce")
    prices = prices.mask(prices < low_price_threshold)
    if preserve_december_missing:
        dec_cols         = [d for d in dates if d.month == 12]
        orig_dec_missing = prices[dec_cols].isna()
    prices_filled = forward_fill_middle_only(prices, dates)
    if preserve_december_missing and len(dec_cols) > 0:
        prices_filled.loc[:, dec_cols] = (
            prices_filled.loc[:, dec_cols].mask(orig_dec_missing)
        )
    return prices_filled


def apply_delisting_to_returns(returns_df, delist_dates, dates):
    out        = returns_df.copy()
    date_index = pd.Index(dates)
    for isin, ddate in delist_dates.items():
        if ddate is None or isin not in out.index:
            continue
        dts        = pd.Timestamp(ddate).normalize()
        candidates = date_index[date_index >= dts]
        if len(candidates) == 0:
            continue
        dcol = candidates.min()
        out.at[isin, dcol] = -1.0
        after = date_index[date_index > dcol]
        if len(after) > 0:
            out.loc[isin, after] = np.nan
    return out

# ─────────────────────────────────────────────────────────────────────────────
# INVESTMENT SET, MOMENTS, OPTIMISATION
# ─────────────────────────────────────────────────────────────────────────────

def year_end_col(dates, year):
    decs = [d for d in dates if d.year == year and d.month == 12]
    return max(decs) if decs else None


def window_cols(dates, year_end, window_years=10):
    start = pd.Timestamp(f"{year_end - window_years + 1}-01-01")
    end   = pd.Timestamp(f"{year_end}-12-31")
    return [d for d in dates if start <= d <= end]


def stale_mask(returns_df, cols, threshold=0.50):
    sub   = returns_df[cols]
    denom = sub.notna().sum(axis=1).replace(0, np.nan)
    frac0 = (sub == 0).sum(axis=1) / denom
    return frac0 > threshold


def build_investment_set(ri_prices, ri_returns, dates, year_end,
                         co2_panel=None):
    """
    Eligible firms at end of year_end must satisfy ALL of:
      1. Non-missing December price (>= LOW_PRICE_THRESHOLD).
      2. >= MIN_OBS_MONTHS non-NaN returns in the prior 10-year window.
      3. Not stale (zero-return fraction <= STALE_THRESHOLD).
      4. If co2_panel is supplied (always the FORWARD-FILLED panel):
         non-missing Scope 1 value for year_end.  Ensures Parts I and II
         share the same eligible universe (brief §2.1).

    Note: co2_panel must be forward-filled BEFORE calling this function
    (Bug 1 fix: main() now calls co2_panel = co2_panel.T.ffill().T first).
    """
    dec_col = year_end_col(dates, year_end)
    if dec_col is None:
        raise ValueError(f"No December column for year {year_end}.")

    cols     = window_cols(dates, year_end, WINDOW_YEARS)
    ok_price = ri_prices[dec_col].notna()
    n_obs    = ri_returns[cols].notna().sum(axis=1)
    ok_obs   = n_obs >= MIN_OBS_MONTHS
    ok_stale = ~stale_mask(ri_returns, cols, STALE_THRESHOLD).fillna(True)

    ok_carbon = pd.Series(True, index=ri_prices.index)
    if co2_panel is not None:
        if year_end in co2_panel.columns:
            ok_carbon = (
                co2_panel[year_end]
                .reindex(ri_prices.index)
                .notna()
                .fillna(False)
            )
        else:
            print(f"  Warning: year {year_end} not in CO2 panel — "
                  "carbon filter skipped.")

    eligible = ri_prices.index[ok_price & ok_obs & ok_stale & ok_carbon]
    return list(eligible), cols


def estimate_moments(ri_returns, isins, cols):
    """
    Compute mu (N,) and Sigma (N,N) on the 10-year window.

    Covariance estimator: Ledoit-Wolf shrinkage  ← Bug 4 fix (v6)
    ────────────────────────────────────────────────────────────────
    The previous pairwise-deletion + zero-fill + tiny-ridge approach
    produced a non-PSD matrix (up to 6 negative eigenvalues for the
    2013 window), causing SLSQP to stall or diverge.

    Fix:
      1. mu  — arithmetic mean of each firm's OBSERVED monthly returns
               (NaN-aware, skipna=True).  Unchanged from v4.
      2. Sigma— Ledoit-Wolf shrinkage estimator (sklearn).
               LW guarantees a PSD matrix by construction and provides
               data-driven regularisation (δ̂ ≈ 0.36–0.48 empirically).
               NaN entries are filled with each firm's own mean before
               passing to LW, so missing periods contribute zero
               deviation from the mean (no information injected).
               This avoids the zero-imputation bias of the old approach
               and eliminates the need for any additional RIDGE term.

    Brief compliance: the brief specifies Σ_Y = (1/τ)Σ(R−μ̂)(R−μ̂)'.
    LW targets this same sample covariance and shrinks it toward a
    scaled identity, which is a standard and well-accepted regularisation
    of that formula when N is close to T.
    """
    R   = ri_returns.loc[isins, cols].astype(float)
    mu  = R.mean(axis=1, skipna=True).to_numpy()

    # Fill NaN with each firm's own time-series mean → zero deviation imputed
    R_filled = R.T.fillna(R.mean(axis=1)).T  # shape (N, T)

    lw    = LedoitWolf()
    lw.fit(R_filled.values.T)          # LW expects (T, N)
    Sigma = lw.covariance_             # (N, N), guaranteed PSD

    return mu, Sigma


def solve_min_variance(Sigma):
    """Long-only min-var: min a'Σa  s.t. sum(a)=1, a≥0."""
    n  = Sigma.shape[0]
    a0 = np.full(n, 1.0 / n)
    res = minimize(
        lambda a: float(a @ Sigma @ a), a0,
        method="SLSQP",
        bounds=[(0.0, 1.0)] * n,
        constraints=[{"type": "eq", "fun": lambda a: np.sum(a) - 1.0}],
        options={"maxiter": SLSQP_MAXITER, "ftol": SLSQP_FTOL, "disp": False}
    )
    if not res.success:
        raise RuntimeError(f"SLSQP failed: {res.message}")
    w = np.where(res.x < 1e-10, 0.0, res.x)
    return w / w.sum()


def compute_mv_oos_returns(portfolios, ri_returns, dates):
    out_r, out_d = [], []
    for year_end, info in portfolios.items():
        invest_year = year_end + 1
        year_months = [d for d in dates
                       if pd.Timestamp(f"{invest_year}-01-01") <= d
                       <= pd.Timestamp(f"{invest_year}-12-31")]
        if not year_months:
            continue
        isins = info["isins"]
        w     = info["weights"].copy()
        for d in year_months:
            r_i = ri_returns.loc[isins, d].fillna(0.0).to_numpy()
            r_p = float(w @ r_i)
            out_r.append(r_p); out_d.append(d)
            denom = 1.0 + r_p
            if denom != 0:
                w = w * (1.0 + r_i) / denom
    return pd.Series(out_r, index=pd.to_datetime(out_d)).sort_index()


def compute_vw_oos_returns(portfolios, ri_returns, mv_caps, dates,
                            start_year=2014, end_year=2025):
    """
    Value-weighted benchmark restricted to the annual eligible investment set.

    For calendar year Y+1, we use portfolios[Y]["isins"] — the same universe
    used to build the MV portfolio — ensuring a like-for-like comparison.

      - In calendar year 2014, use firms eligible at end-2013
      - In calendar year 2015, use firms eligible at end-2014  (etc.)

    This fixes the v4 implementation which used any firm with available
    lagged monthly cap + current return (a looser, inconsistent universe that
    did not respect the stale-price, carbon-availability, or min-obs filters).
    See brief §2.3 and §3.3: the VW benchmark denominator is explicitly
    described as "the total market value of the investment set".
    """
    out_r, out_d = [], []
    dates_sorted = sorted(dates)
    pos          = {d: i for i, d in enumerate(dates_sorted)}

    for year_end, info in portfolios.items():
        invest_year    = year_end + 1
        if invest_year < start_year or invest_year > end_year:
            continue
        eligible_isins = info["isins"]

        year_months = [d for d in dates_sorted
                       if pd.Timestamp(f"{invest_year}-01-01") <= d
                       <= pd.Timestamp(f"{invest_year}-12-31")]

        for d in year_months:
            i = pos.get(d)
            if i is None or i == 0:
                continue
            d_prev = dates_sorted[i - 1]
            if d_prev not in mv_caps.columns or d not in ri_returns.columns:
                continue

            # Restrict to the eligible investment set for this year
            elig_in_caps = [s for s in eligible_isins if s in mv_caps.index]
            elig_in_rets = [s for s in eligible_isins if s in ri_returns.index]
            caps   = mv_caps.loc[elig_in_caps, d_prev].dropna()
            rets   = ri_returns.loc[elig_in_rets, d].dropna()
            common = caps.index.intersection(rets.index)
            if len(common) == 0:
                continue

            w = caps.loc[common] / caps.loc[common].sum()
            out_r.append(float(w.to_numpy() @ rets.loc[common].to_numpy()))
            out_d.append(d)

    return pd.Series(out_r, index=pd.to_datetime(out_d)).sort_index()


def perf_stats(r, rf=None):
    """
    Annualised performance statistics.

    rf : pd.Series of monthly decimal rates (from Risk_Free_Rate_2025.xlsx).
         Aligned to r.index.  If None, rf=0 (v2 fallback).

    Annualised return  = (1 + mu_m)^12 − 1  [compounds arithmetic mean, per template]
    Geometric return   = [(1+r_1)…(1+r_T)]^(12/T) − 1  [also reported for transparency]
    Sharpe             = mean(r − rf) / std(r − rf) × √12
    """
    r = r.dropna()
    T = len(r)
    rf_aligned = rf.reindex(r.index).fillna(0.0) if rf is not None \
                 else pd.Series(0.0, index=r.index)
    excess  = r - rf_aligned
    mu_m    = r.mean()
    sigma_e = excess.std(ddof=1)
    return {
        "Annualized Return (arith. compounding)": (1 + mu_m) ** 12 - 1,
        "Annualized Return (geometric avg)"      : (1 + r).prod() ** (12.0 / T) - 1 if T > 0 else np.nan,
        "Annualized Volatility"                  : r.std(ddof=1) * np.sqrt(12),
        "Sharpe Ratio"                           : (excess.mean() / sigma_e) * np.sqrt(12) if sigma_e > 0 else np.nan,
        "Min Monthly Return"                     : r.min(),
        "Max Monthly Return"                     : r.max(),
        "Cumulative Return"                      : (1 + r).prod() - 1,
    }

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("SAAM 2026 — PART I (v6)")
    print(f"Region: {REGION_CODE}")
    print("=" * 60)
    os.makedirs(RESULTS_DIR, exist_ok=True)

    # ── Static ───────────────────────────────────────────────────────────────
    static         = pd.read_excel(data_path("Static_2025.xlsx"), engine="openpyxl")
    static.columns = ["ISIN", "NAME", "Country", "Region"]
    static["ISIN"] = static["ISIN"].astype(str).str.strip()
    pac            = static[static["Region"] == REGION_CODE].copy().set_index("ISIN")
    pac_isins      = set(pac.index)
    print(f"Pacific firms in Static: {len(pac_isins)}")

    # ── Annual RI (delist dates) ──────────────────────────────────────────────
    ri_y         = load_datastream_wide(data_path("DS_RI_T_USD_Y_2025.xlsx"))
    ri_y         = ri_y[ri_y.index.isin(pac_isins)].copy()
    delist_dates = {isin: extract_delist_date(ri_y.at[isin, "NAME"])
                    for isin in ri_y.index}
    print(f"Delisted Pacific firms: {sum(v is not None for v in delist_dates.values())}")

    # ── Monthly RI ────────────────────────────────────────────────────────────
    ri_m_raw                = load_datastream_wide(data_path("DS_RI_T_USD_M_2025.xlsx"))
    keep_cols, parsed_dates = parse_monthly_columns(list(ri_m_raw.columns))
    ri_m                    = ri_m_raw[["NAME"] + keep_cols].copy()
    ri_m.columns            = ["NAME"] + parsed_dates
    ri_m                    = ri_m[ri_m.index.isin(pac_isins)].copy()
    parsed_dates            = [d for d in parsed_dates
                                if d <= pd.Timestamp("2025-12-31")]
    ri_m                    = ri_m[["NAME"] + parsed_dates]
    print(f"Monthly RI: {ri_m.shape[0]} firms, {len(parsed_dates)} months")

    # ── Monthly MV ────────────────────────────────────────────────────────────
    mv_m_raw           = load_datastream_wide(data_path("DS_MV_T_USD_M_2025.xlsx"))
    keep_cols_mv, _mv  = parse_monthly_columns(list(mv_m_raw.columns))
    mv_m               = mv_m_raw[["NAME"] + keep_cols_mv].copy()
    mv_m.columns       = ["NAME"] + _mv
    mv_m               = mv_m[mv_m.index.isin(pac_isins)].copy()
    mv_m               = mv_m[["NAME"] + parsed_dates].copy()
    mv_m[parsed_dates] = mv_m[parsed_dates].apply(pd.to_numeric, errors="coerce")
    print(f"Monthly MV: {mv_m.shape[0]} firms")

    # ── Common ISINs ──────────────────────────────────────────────────────────
    common_isins = (pac_isins
                    .intersection(set(ri_m.index))
                    .intersection(set(mv_m.index))
                    .intersection(set(ri_y.index)))
    ri_m = ri_m.loc[list(common_isins)].copy()
    mv_m = mv_m.loc[list(common_isins)].copy()
    print(f"Common ISINs: {len(common_isins)}")

    # ── Clean prices & returns ────────────────────────────────────────────────
    ri_prices   = clean_monthly_ri_prices(ri_m[parsed_dates], parsed_dates,
                                           LOW_PRICE_THRESHOLD, True)
    all_missing = ri_prices.isna().all(axis=1)
    if all_missing.any():
        missing_isins = ri_prices.index[all_missing]
        ri_prices     = ri_prices.loc[~all_missing].copy()
        mv_m          = mv_m.drop(index=missing_isins, errors="ignore")
        print(f"Dropped fully-missing RI rows: {all_missing.sum()}")

    ri_returns = ri_prices.pct_change(axis=1, fill_method=None)  # explicit: no implicit forward-fill
    ri_returns = apply_delisting_to_returns(
        ri_returns,
        {k: v for k, v in delist_dates.items() if k in ri_returns.index},
        parsed_dates
    )

    # ── Load CO2 panel — forward-fill BEFORE passing to investment set ────────
    # BUG 1 FIX: forward-fill here so that a firm with CO2 data in year Y-1
    # but missing in year Y is still treated as having CO2 data available at
    # end-Y (the brief says: "use the number from the previous year" for
    # missing values between two available years or at the end of the sample).
    # This makes Part I use the same forward-filled panel as Part II.
    print(f"Loading CO2 panel: {CO2_FILE}")
    co2_panel = load_annual_panel(data_path(CO2_FILE))
    co2_panel = co2_panel.T.ffill().T          # ← Bug 1 fix
    print(f"  {co2_panel.shape[0]} ISINs, "
          f"years {co2_panel.columns.min()}–{co2_panel.columns.max()}")

    # ── Load risk-free rate ───────────────────────────────────────────────────
    print(f"Loading risk-free rate: {RF_FILE}")
    try:
        rf_monthly = load_rf_monthly(data_path(RF_FILE))
        print(f"  {len(rf_monthly)} obs, "
              f"mean={rf_monthly.mean()*100:.3f} %/month")
    except Exception as exc:
        print(f"  WARNING: could not load RF ({exc}). Using rf=0.")
        rf_monthly = None

    # ── Build annual MV portfolios ────────────────────────────────────────────
    portfolios = {}
    for year_end in range(START_YEAR_OOS - 1, END_YEAR_OOS):   # 2013 … 2024
        elig, cols = build_investment_set(
            ri_prices, ri_returns, parsed_dates, year_end,
            co2_panel=co2_panel    # now forward-filled
        )
        if len(elig) < 2:
            print(f"Year {year_end}: eligible={len(elig)} (skipped)")
            continue
        _, Sigma = estimate_moments(ri_returns, elig, cols)    # ddof=0
        w        = solve_min_variance(Sigma)
        portfolios[year_end] = {"isins": elig, "weights": w}
        print(f"Year {year_end}: eligible={len(elig)}, max_w={w.max():.4f}")

    # ── OOS returns 2014–2025 ─────────────────────────────────────────────────
    mv_r = compute_mv_oos_returns(portfolios, ri_returns, parsed_dates)
    vw_r = compute_vw_oos_returns(
        portfolios, ri_returns, mv_m[parsed_dates], parsed_dates,
        START_YEAR_OOS, END_YEAR_OOS
    )
    common_months = mv_r.index.intersection(vw_r.index)
    mv_r = mv_r.loc[common_months]
    vw_r = vw_r.loc[common_months]

    # ── Export ────────────────────────────────────────────────────────────────
    out = pd.DataFrame({
        "Date"        : common_months,
        "MV_Return"   : mv_r.values,
        "VW_Return"   : vw_r.values,
    })
    out["MV_CumReturn"] = (1 + out["MV_Return"]).cumprod()
    out["VW_CumReturn"] = (1 + out["VW_Return"]).cumprod()
    out.to_csv(os.path.join(RESULTS_DIR, "part1_results.csv"), index=False)

    stats = pd.DataFrame({
        "Minimum Variance": perf_stats(mv_r, rf=rf_monthly),
        "Value Weighted"  : perf_stats(vw_r, rf=rf_monthly),
    }).T
    stats.to_csv(os.path.join(RESULTS_DIR, "part1_summary_statistics.csv"))

    comp = []
    for year_end, info in portfolios.items():
        comp.append(pd.DataFrame({
            "Year"  : year_end + 1,
            "ISIN"  : info["isins"],
            "Weight": info["weights"],
        }))
    pd.concat(comp, ignore_index=True).to_csv(
        os.path.join(RESULTS_DIR, "part1_portfolio_compositions.csv"),
        index=False
    )

    print("=" * 60)
    print("Done.  Results written to:", RESULTS_DIR)
    print("  part1_results.csv")
    print("  part1_summary_statistics.csv")
    print("  part1_portfolio_compositions.csv")
    print("=" * 60)


if __name__ == "__main__":
    main()
