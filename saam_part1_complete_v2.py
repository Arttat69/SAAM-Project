# ============================================================
# SAAM PROJECT 2026 — PART I (v2, with full data-cleaning logic)
# Group: Pacific Area | Scope 1
# ============================================================

import os
import pandas as pd
import numpy as np
import re
from datetime import datetime
from scipy.optimize import minimize

# -----------------------------
# Config
# -----------------------------
REGION_CODE = "PAC"
LOW_PRICE_THRESHOLD = 0.5
RESULTS_DIR = "resultsPart1"

WINDOW_YEARS = 10
MIN_OBS_MONTHS = 36          # at least 3 years of monthly returns in window
STALE_THRESHOLD = 0.50       # >50% zero-return months => exclude

START_YEAR_OOS = 2014
END_YEAR_OOS   = 2025

# SLSQP / numerical stability
RIDGE = 1e-8                 # add ridge to covariance diag
SLSQP_MAXITER = 400
SLSQP_FTOL = 1e-9


# ============================================================
# Helpers: loading & cleaning
# ============================================================

def load_datastream_wide(filepath, sheet=0):
    """
    Load a Datastream wide-format Excel export:
      columns: NAME | ISIN | date1 | date2 | ...
    Drops $$ER rows and rows with missing ISIN.
    """
    df = pd.read_excel(filepath, sheet_name=sheet, header=0, engine="openpyxl")

    df.columns = ["NAME", "ISIN"] + list(df.columns[2:])
    df = df[~df["NAME"].astype(str).str.startswith("$$ER", na=False)]
    df = df.dropna(subset=["ISIN"])
    df["ISIN"] = df["ISIN"].astype(str).str.strip()
    df["NAME"] = df["NAME"].astype(str).str.strip()
    df = df[df["ISIN"] != ""]
    df = df.set_index("ISIN")
    return df


def parse_monthly_columns(cols):
    """
    Convert Datastream column headers to datetime.
    Keeps only columns that successfully parse.
    """
    parsed = []
    keep = []
    for c in cols:
        if c == "NAME":
            continue
        dt = pd.to_datetime(str(c), errors="coerce", dayfirst=True)
        if pd.notna(dt):
            keep.append(c)
            parsed.append(pd.Timestamp(dt).normalize())
    return keep, parsed


def extract_delist_date(name_str):
    """
    Extract delist date from strings like:
      '... DELIST.02/07/24' or '... DEAD.10/07/24'
    """
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
    """
    Forward-fill ONLY between first and last valid observation per row.
    Leading NaNs (before listing) stay NaN.
    Trailing NaNs (after delisting / end) stay NaN.
    """
    arr = price_df.to_numpy(dtype=float, copy=True)
    n, t = arr.shape

    for i in range(n):
        row = arr[i, :]
        valid_idx = np.where(~np.isnan(row))[0]
        if valid_idx.size == 0:
            continue
        first, last = valid_idx[0], valid_idx[-1]

        # forward fill within [first, last]
        last_val = row[first]
        for k in range(first + 1, last + 1):
            if np.isnan(row[k]):
                row[k] = last_val
            else:
                last_val = row[k]
        arr[i, :] = row

    return pd.DataFrame(arr, index=price_df.index, columns=dates)


def clean_monthly_ri_prices(raw_ri_m, dates, low_price_threshold=0.5, preserve_december_missing=True):
    """
    Implements:
    - Low prices (<0.5) treated as missing (NaN).
    - Missing values in the middle are forward-filled (previous value) within each firm’s listed span.
    - If December (year-end) is missing originally, keep it missing (so the firm is excluded for Y+1).
    """
    prices = raw_ri_m.copy()
    prices = prices.apply(pd.to_numeric, errors="coerce")

    # low price => missing
    prices = prices.mask(prices < low_price_threshold)

    # record which December months were missing originally
    if preserve_december_missing:
        dec_cols = [d for d in dates if d.month == 12]
        orig_dec_missing = prices[dec_cols].isna()

    # fill middle gaps only
    prices_filled = forward_fill_middle_only(prices, dates)

    # restore year-end (December) missingness if originally missing
    if preserve_december_missing and len(dec_cols) > 0:
        prices_filled.loc[:, dec_cols] = prices_filled.loc[:, dec_cols].mask(orig_dec_missing)

    return prices_filled


def apply_delisting_to_returns(returns_df, delist_dates, dates):
    """
    For delisted firms:
      - set return = -1.0 at first month >= delist_date (acknowledge price goes to 0)
      - set subsequent returns to NaN
    """
    out = returns_df.copy()
    date_index = pd.Index(dates)

    for isin, ddate in delist_dates.items():
        if ddate is None or isin not in out.index:
            continue
        dts = pd.Timestamp(ddate).normalize()
        # first monthly column on/after delist date
        candidates = date_index[date_index >= dts]
        if len(candidates) == 0:
            continue
        dcol = candidates.min()

        out.at[isin, dcol] = -1.0
        after = date_index[date_index > dcol]
        if len(after) > 0:
            out.loc[isin, after] = np.nan

    return out


# ============================================================
# Investment set, stale filter, moments, optimization
# ============================================================

def year_end_col(dates, year):
    """
    Return the month-end column for December of given year.
    Datastream exports month-ends; take the max date with month=12.
    """
    decs = [d for d in dates if d.year == year and d.month == 12]
    return max(decs) if decs else None


def window_cols(dates, year_end, window_years=10):
    """
    10-year window ending at Dec of year_end.
    Example: year_end=2013 => Jan 2004 ... Dec 2013 (120 months).
    """
    start = pd.Timestamp(f"{year_end - window_years + 1}-01-01")
    end   = pd.Timestamp(f"{year_end}-12-31")
    return [d for d in dates if start <= d <= end]


def stale_mask(returns_df, cols, threshold=0.50):
    """
    fraction of months with return == 0 among non-missing returns.
    """
    sub = returns_df[cols]
    denom = sub.notna().sum(axis=1).replace(0, np.nan)
    frac0 = (sub == 0).sum(axis=1) / denom
    return frac0 > threshold


def build_investment_set(ri_prices, ri_returns, dates, year_end):
    """
    Criteria (Part I):
      - year-end (Dec) price is available and >= 0.5 (implemented as not-NaN after cleaning)
      - enough return observations in the last 10 years (>= MIN_OBS_MONTHS non-NaN)
      - not stale (>50% zero-return months)
    """
    dec_col = year_end_col(dates, year_end)
    if dec_col is None:
        raise ValueError(f"No December column found for year {year_end}.")

    cols = window_cols(dates, year_end, WINDOW_YEARS)

    # year-end price must be available (do not forward-fill December if it was missing)
    ok_price = ri_prices[dec_col].notna()

    # enough data in window
    n_obs = ri_returns[cols].notna().sum(axis=1)
    ok_obs = n_obs >= MIN_OBS_MONTHS

    # stale filter
    stale = stale_mask(ri_returns, cols, STALE_THRESHOLD).fillna(True)  # if denom=0 => stale/exclude
    ok_stale = ~stale

    eligible = ri_prices.index[ok_price & ok_obs & ok_stale]
    return list(eligible), cols


def estimate_moments(ri_returns, isins, cols):
    """
    Compute mu and Sigma on the 10-year window.
    We fill remaining NaNs with 0 ONLY for matrix operations (keeps window length fixed).
    """
    R = ri_returns.loc[isins, cols].astype(float)
    R = R.fillna(0.0)

    mu = R.mean(axis=1).to_numpy()
    X = R.to_numpy()
    X = X - X.mean(axis=1, keepdims=True)
    T = X.shape[1]
    Sigma = (X @ X.T) / T

    # ridge for numerical stability
    Sigma = Sigma + RIDGE * np.eye(Sigma.shape[0])
    return mu, Sigma


def solve_min_variance(Sigma):
    """
    Long-only min-variance:
      min a' Σ a  s.t. sum(a)=1, a>=0
    """
    n = Sigma.shape[0]

    def obj(a):
        return float(a @ Sigma @ a)

    cons = [{"type": "eq", "fun": lambda a: np.sum(a) - 1.0}]
    bnds = [(0.0, 1.0)] * n
    a0 = np.full(n, 1.0 / n)

    res = minimize(
        obj, a0, method="SLSQP",
        bounds=bnds, constraints=cons,
        options={"maxiter": SLSQP_MAXITER, "ftol": SLSQP_FTOL, "disp": False}
    )
    if not res.success:
        raise RuntimeError(f"SLSQP failed: {res.message}")

    w = res.x
    w = np.where(w < 1e-10, 0.0, w)
    w = w / w.sum()  # re-normalize

    return w


def compute_mv_oos_returns(portfolios, ri_returns, dates):
    """
    Compute out-of-sample monthly returns for MV strategy with weight drift inside each year.
    """
    out_r = []
    out_d = []

    for year_end, info in portfolios.items():
        invest_year = year_end + 1
        year_months = [d for d in dates if pd.Timestamp(f"{invest_year}-01-01") <= d <= pd.Timestamp(f"{invest_year}-12-31")]
        if len(year_months) == 0:
            continue

        isins = info["isins"]
        w = info["weights"].copy()

        for d in year_months:
            r_i = ri_returns.loc[isins, d].fillna(0.0).to_numpy()
            r_p = float(w @ r_i)

            out_r.append(r_p)
            out_d.append(d)

            # drift weights
            denom = (1.0 + r_p)
            if denom != 0:
                w = w * (1.0 + r_i) / denom
            # if denom==0 (portfolio wiped), keep weights unchanged (edge case)

    return pd.Series(out_r, index=pd.to_datetime(out_d)).sort_index()


def compute_vw_oos_returns(ri_returns, mv_caps, dates, start_year=2014, end_year=2025):
    """
    Monthly-rebalanced value-weighted benchmark:
      weights at t-1 from market caps at t-1, applied to returns at t.
    """
    out_r = []
    out_d = []

    months = [d for d in dates if pd.Timestamp(f"{start_year}-01-01") <= d <= pd.Timestamp(f"{end_year}-12-31")]
    dates_sorted = sorted(dates)

    pos = {d: i for i, d in enumerate(dates_sorted)}

    for d in months:
        i = pos.get(d, None)
        if i is None or i == 0:
            continue
        d_prev = dates_sorted[i - 1]

        if d_prev not in mv_caps.columns or d not in ri_returns.columns:
            continue

        caps = mv_caps[d_prev].dropna()
        rets = ri_returns[d].dropna()

        common = caps.index.intersection(rets.index)
        if len(common) == 0:
            continue

        w = caps.loc[common] / caps.loc[common].sum()
        r = float(w.to_numpy() @ rets.loc[common].to_numpy())

        out_r.append(r)
        out_d.append(d)

    return pd.Series(out_r, index=pd.to_datetime(out_d)).sort_index()


def perf_stats(r):
    """
    Annualized mean, vol, Sharpe (rf=0), min, max.
    Uses arithmetic mean monthly -> annualized geometric approximation per template convention.
    """
    r = r.dropna()
    mu_m = r.mean()
    vol_m = r.std(ddof=1)
    ann_return = (1 + mu_m) ** 12 - 1
    ann_vol = vol_m * np.sqrt(12)
    sharpe = ann_return / ann_vol if ann_vol > 0 else np.nan
    return {
        "Annualized Return": ann_return,
        "Annualized Volatility": ann_vol,
        "Sharpe Ratio": sharpe,
        "Min Monthly Return": r.min(),
        "Max Monthly Return": r.max(),
        "Cumulative Return": (1 + r).prod() - 1
    }


# ============================================================
# Main
# ============================================================

def main():
    print("=" * 60)
    print("SAAM 2026 — PART I (v2)")
    print(f"Region: {REGION_CODE}")
    print("=" * 60)

    # ensure results directory exists
    os.makedirs(RESULTS_DIR, exist_ok=True)

    # --- Static (Pacific universe)
    static = pd.read_excel("Static_2025.xlsx", engine="openpyxl")
    static.columns = ["ISIN", "NAME", "Country", "Region"]
    static["ISIN"] = static["ISIN"].astype(str).str.strip()
    static["NAME"] = static["NAME"].astype(str).str.strip()

    pac = static[static["Region"] == REGION_CODE].copy().set_index("ISIN")
    pac_isins = set(pac.index)
    print(f"Pacific firms in Static: {len(pac_isins)}")

    # --- Annual RI (for delist date in NAME)
    ri_y = load_datastream_wide("DS_RI_T_USD_Y_2025.xlsx")
    ri_y = ri_y[ri_y.index.isin(pac_isins)].copy()
    delist_dates = {isin: extract_delist_date(ri_y.at[isin, "NAME"]) for isin in ri_y.index}
    n_del = sum(v is not None for v in delist_dates.values())
    print(f"Delisted Pacific firms (from annual NAME parsing): {n_del}")

    # --- Monthly RI
    ri_m_raw = load_datastream_wide("DS_RI_T_USD_M_2025.xlsx")
    keep_cols, parsed_dates = parse_monthly_columns(list(ri_m_raw.columns))
    ri_m = ri_m_raw[["NAME"] + keep_cols].copy()
    ri_m.columns = ["NAME"] + parsed_dates
    ri_m = ri_m[ri_m.index.isin(pac_isins)].copy()

    # Keep only up to end-2025 for Part I
    parsed_dates = [d for d in parsed_dates if d <= pd.Timestamp("2025-12-31")]
    ri_m = ri_m[["NAME"] + parsed_dates]
    print(f"Monthly RI Pacific firms: {ri_m.shape[0]}, months: {len(parsed_dates)}")

    # --- Monthly MV (market caps)
    mv_m_raw = load_datastream_wide("DS_MV_T_USD_M_2025.xlsx")
    keep_cols_mv, parsed_dates_mv = parse_monthly_columns(list(mv_m_raw.columns))

    mv_m = mv_m_raw[["NAME"] + keep_cols_mv].copy()
    mv_m.columns = ["NAME"] + parsed_dates_mv
    mv_m = mv_m[mv_m.index.isin(pac_isins)].copy()

    # align MV months to RI months (same date columns)
    mv_m = mv_m[["NAME"] + parsed_dates].copy()
    mv_m[parsed_dates] = mv_m[parsed_dates].apply(pd.to_numeric, errors="coerce")
    print(f"Monthly MV Pacific firms: {mv_m.shape[0]}, months: {len(parsed_dates)}")

    # --- Enforce common ISINs across required tables (drop missing rows consistently)
    common_isins = pac_isins.intersection(set(ri_m.index)).intersection(set(mv_m.index)).intersection(set(ri_y.index))
    ri_m = ri_m.loc[list(common_isins)].copy()
    mv_m = mv_m.loc[list(common_isins)].copy()
    print(f"Common ISINs across Static/RI/MV/Annual RI: {len(common_isins)}")

    # --- Clean RI monthly prices with middle-gap ffill and December-missing preservation
    ri_prices = clean_monthly_ri_prices(
        raw_ri_m=ri_m[parsed_dates],
        dates=parsed_dates,
        low_price_threshold=LOW_PRICE_THRESHOLD,
        preserve_december_missing=True
    )

    # Drop rows that are fully missing after cleaning (missing prices / no match)
    all_missing = ri_prices.isna().all(axis=1)
    if all_missing.any():
        drop_isins = ri_prices.index[all_missing]
        ri_prices = ri_prices.loc[~all_missing].copy()
        mv_m = mv_m.drop(index=drop_isins, errors="ignore")
        print(f"Dropped fully-missing RI rows (no usable RI data): {len(drop_isins)}")

    # --- Compute returns from cleaned prices
    ri_returns = ri_prices.pct_change(axis=1)

    # --- Apply delisting rule to returns (-100% in delist month, NaN after)
    delist_dates_common = {k: v for k, v in delist_dates.items() if k in ri_returns.index}
    ri_returns = apply_delisting_to_returns(ri_returns, delist_dates_common, parsed_dates)

    # --- Build annual MV portfolios (decisions at end of Y for Y+1)
    portfolios = {}
    for year_end in range(START_YEAR_OOS - 1, END_YEAR_OOS):  # 2013..2024
        elig, cols = build_investment_set(ri_prices, ri_returns, parsed_dates, year_end)
        if len(elig) < 2:
            print(f"Year {year_end}: eligible={len(elig)} (skipped)")
            continue

        mu, Sigma = estimate_moments(ri_returns, elig, cols)
        w = solve_min_variance(Sigma)

        portfolios[year_end] = {"isins": elig, "weights": w}
        print(f"Year {year_end}: eligible={len(elig)}, max_w={w.max():.4f}")

    # --- Compute OOS returns 2014..2025
    mv_r = compute_mv_oos_returns(portfolios, ri_returns, parsed_dates)
    vw_r = compute_vw_oos_returns(ri_returns, mv_m[parsed_dates], parsed_dates, START_YEAR_OOS, END_YEAR_OOS)

    # Align to common months
    common_months = mv_r.index.intersection(vw_r.index)
    mv_r = mv_r.loc[common_months]
    vw_r = vw_r.loc[common_months]

    # --- Export (to results directory)
    out = pd.DataFrame({
        "Date": common_months,
        "MV_Return": mv_r.values,
        "VW_Return": vw_r.values
    })
    out["MV_CumReturn"] = (1 + out["MV_Return"]).cumprod()
    out["VW_CumReturn"] = (1 + out["VW_Return"]).cumprod()
    out.to_csv(os.path.join(RESULTS_DIR, "part1_results.csv"), index=False)

    stats = pd.DataFrame({
        "Minimum Variance": perf_stats(mv_r),
        "Value Weighted": perf_stats(vw_r)
    }).T
    stats.to_csv(os.path.join(RESULTS_DIR, "part1_summary_statistics.csv"))

    comp = []
    for year_end, info in portfolios.items():
        comp.append(pd.DataFrame({
            "Year": year_end + 1,
            "ISIN": info["isins"],
            "Weight": info["weights"]
        }))
    pd.concat(comp, ignore_index=True).to_csv(
        os.path.join(RESULTS_DIR, "part1_portfolio_compositions.csv"),
        index=False
    )

    print("=" * 60)
    print("Done.")
    print("Wrote results to 'results/' directory:")
    print(" - part1_results.csv")
    print(" - part1_summary_statistics.csv")
    print(" - part1_portfolio_compositions.csv")
    print("=" * 60)


if __name__ == "__main__":
    main()
