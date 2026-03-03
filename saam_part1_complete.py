# ============================================================
# SAAM PROJECT 2026 - PART I: COMPLETE CODE
# Group: Pacific Area | Scope 1
# ============================================================

import pandas as pd
import numpy as np
import re
from datetime import datetime
from scipy.optimize import minimize
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def load_datastream_wide(filepath, sheet=0):
    """Load Datastream wide-format Excel file."""
    df = pd.read_excel(filepath, sheet_name=sheet, header=0)
    df.columns = ['NAME', 'ISIN'] + list(df.columns[2:])
    df = df[~df['NAME'].astype(str).str.startswith('$$ER')]
    df = df.dropna(subset=['ISIN'])
    df = df[df['ISIN'].astype(str).str.strip() != '']
    df['ISIN'] = df['ISIN'].astype(str).str.strip()
    df['NAME'] = df['NAME'].astype(str).str.strip()
    df = df.set_index('ISIN')
    return df

def extract_delist_date(name_str):
    """Extract delist date from firm name."""
    match = re.search(r'(?:DELIST|DEAD)\.(\d{2}/\d{2}/\d{2,4})', str(name_str))
    if match:
        date_str = match.group(1)
        for fmt in ('%d/%m/%y', '%d/%m/%Y'):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
    return None

# ============================================================
# 1. LOAD STATIC FILE - PACIFIC FIRMS
# ============================================================
print("=" * 60)
print("SAAM PROJECT 2026 - PART I")
print("Group: Pacific Area | Scope 1")
print("=" * 60)

print("\n[1/8] Loading Static file...")
static_raw = pd.read_excel('Static_2025.xlsx', header=0)
static_raw.columns = ['ISIN', 'NAME', 'Country', 'Region']
static_raw['ISIN'] = static_raw['ISIN'].astype(str).str.strip()
static_raw['NAME'] = static_raw['NAME'].astype(str).str.strip()

pacific = static_raw[static_raw['Region'] == 'PAC'].copy()
pacific = pacific.set_index('ISIN')
pac_isins = set(pacific.index)

print(f"   Total firms: {len(static_raw)}")
print(f"   Pacific firms: {len(pacific)}")

# ============================================================
# 2. LOAD ANNUAL RI FILE
# ============================================================
print("\n[2/8] Loading Annual RI file...")
ri_y_raw = load_datastream_wide('DS_RI_T_USD_Y_2025-3.xlsx')

year_cols = [c for c in ri_y_raw.columns if str(c).isdigit() and 1999 <= int(str(c)) <= 2025]
ri_y_raw = ri_y_raw[['NAME'] + year_cols]
ri_y = ri_y_raw[ri_y_raw.index.isin(pac_isins)].copy()
ri_y[year_cols] = ri_y[year_cols].apply(pd.to_numeric, errors='coerce')

# Extract delist dates
ri_y['delist_date'] = ri_y['NAME'].apply(extract_delist_date)
ri_y['is_delisted'] = ri_y['delist_date'].notna()

print(f"   Pacific firms in annual RI: {len(ri_y)}")
print(f"   Delisted firms: {ri_y['is_delisted'].sum()}")

# ============================================================
# 3. LOAD MONTHLY RI FILE
# ============================================================
print("\n[3/8] Loading Monthly RI file...")
ri_m_raw = load_datastream_wide('DS_RI_T_USD_M_2025.xlsx')

# Parse date columns
date_cols = [c for c in ri_m_raw.columns if c != 'NAME']
parsed_dates = []
for c in date_cols:
    try:
        parsed_dates.append(pd.to_datetime(c))
    except:
        parsed_dates.append(None)

valid_pairs = [(orig, parsed) for orig, parsed in zip(date_cols, parsed_dates) if parsed is not None]
valid_orig = [v[0] for v in valid_pairs]
valid_parsed = [v[1] for v in valid_pairs]

ri_m = ri_m_raw[['NAME'] + valid_orig].copy()
ri_m.columns = ['NAME'] + valid_parsed
ri_m = ri_m[ri_m.index.isin(pac_isins)].copy()
ri_m[valid_parsed] = ri_m[valid_parsed].apply(pd.to_numeric, errors='coerce')

print(f"   Pacific firms: {len(ri_m)}")
print(f"   Date range: {min(valid_parsed)} to {max(valid_parsed)}")

# Apply low price rule: RI < 0.5 → NaN
for col in valid_parsed:
    ri_m.loc[ri_m[col] < 0.5, col] = np.nan

# Compute monthly returns
ri_m_returns = ri_m[valid_parsed].pct_change(axis=1)

# Handle delisted firms: -100% return at delist month
for isin in ri_m.index:
    if isin in ri_y.index and ri_y.at[isin, 'is_delisted']:
        delist_date = ri_y.at[isin, 'delist_date']
        if delist_date is not None:
            delist_cols = [col for col in valid_parsed if col >= pd.Timestamp(delist_date)]
            if delist_cols:
                delist_col = min(delist_cols)
                ri_m_returns.at[isin, delist_col] = -1.0
                after_cols = [col for col in valid_parsed if col > delist_col]
                ri_m_returns.loc[isin, after_cols] = np.nan

print("   Monthly returns computed.")

# ============================================================
# 4. LOAD MONTHLY MARKET CAP FILE
# ============================================================
print("\n[4/8] Loading Monthly Market Cap file...")
mv_m_raw = load_datastream_wide('DS_MV_T_USD_M_2025-2.xlsx')

mv_m = mv_m_raw[['NAME'] + valid_orig].copy()
mv_m.columns = ['NAME'] + valid_parsed
mv_m = mv_m[mv_m.index.isin(pac_isins)].copy()
mv_m[valid_parsed] = mv_m[valid_parsed].apply(pd.to_numeric, errors='coerce')

print(f"   Pacific firms: {len(mv_m)}")

# ============================================================
# 5. INVESTMENT SET CONSTRUCTION
# ============================================================

def get_stale_mask(return_matrix, all_dates, year_end, window_years=10, threshold=0.50):
    """Compute stale price filter."""
    window_end = pd.Timestamp(f'{year_end}-12-31')
    window_start = pd.Timestamp(f'{year_end - window_years}-12-31')
    window_cols = [d for d in all_dates if window_start < d <= window_end]
    
    sub = return_matrix[window_cols].astype(float)
    zero_fraction = (sub == 0).sum(axis=1) / sub.notna().sum(axis=1)
    stale_mask = zero_fraction > threshold
    return stale_mask

def build_investment_set(year_end, return_matrix, all_dates, ri_y_prices, 
                         pacific_isins, min_obs_months=36, window_years=10):
    """Build investment set for a given year."""
    window_end = pd.Timestamp(f'{year_end}-12-31')
    window_start = pd.Timestamp(f'{year_end - window_years}-12-31')
    window_cols = [d for d in all_dates if window_start < d <= window_end]
    
    eligible = []
    for isin in pacific_isins:
        if isin not in return_matrix.index:
            continue
        
        # Check annual price at year end
        yr_col = str(year_end)
        if yr_col in ri_y_prices.columns:
            price_end = ri_y_prices.at[isin, yr_col] if isin in ri_y_prices.index else np.nan
            if pd.isna(price_end):
                continue
        
        # Check minimum observations
        n_valid = return_matrix.loc[isin, window_cols].notna().sum()
        if n_valid < min_obs_months:
            continue
        
        eligible.append(isin)
    
    # Apply stale price filter
    stale = get_stale_mask(return_matrix, all_dates, year_end, window_years)
    stale_isins = set(stale[stale].index)
    eligible = [isin for isin in eligible if isin not in stale_isins]
    
    return eligible

def estimate_moments(isins, return_matrix, all_dates, year_end, window_years=10):
    """Estimate expected returns and covariance matrix."""
    window_end = pd.Timestamp(f'{year_end}-12-31')
    window_start = pd.Timestamp(f'{year_end - window_years}-12-31')
    window_cols = [d for d in all_dates if window_start < d <= window_end]
    
    R = return_matrix.loc[isins, window_cols].astype(float)
    R = R.fillna(0)
    
    mu = R.mean(axis=1).values
    R_dem = R.subtract(R.mean(axis=1), axis=0)
    T = R.shape[1]
    Sigma = (R_dem.values @ R_dem.values.T) / T
    
    return mu, Sigma, list(R.index)

# ============================================================
# 6. MINIMUM VARIANCE OPTIMIZATION
# ============================================================

def solve_min_variance(Sigma, N):
    """Solve long-only minimum variance portfolio."""
    
    def objective(alpha):
        return alpha @ Sigma @ alpha
    
    constraints = [{'type': 'eq', 'fun': lambda alpha: np.sum(alpha) - 1}]
    bounds = [(0, 1) for _ in range(N)]
    alpha0 = np.ones(N) / N
    
    result = minimize(objective, alpha0, method='SLSQP', 
                     bounds=bounds, constraints=constraints,
                     options={'maxiter': 1000, 'ftol': 1e-9})
    
    if not result.success:
        print(f"      ⚠️  Optimization warning: {result.message}")
    
    return result.x

# ============================================================
# 7. ROLLING PORTFOLIO CONSTRUCTION (2014-2025)
# ============================================================

print("\n[5/8] Building portfolios (2014-2025)...")

portfolio_data = {}
rebalance_years = list(range(2013, 2025))

for year_end in rebalance_years:
    print(f"\n   Year {year_end}:")
    
    inv_set = build_investment_set(year_end, ri_m_returns, valid_parsed,
                                   ri_y[year_cols], pac_isins)
    
    print(f"      Eligible firms: {len(inv_set)}")
    
    if len(inv_set) < 2:
        print("      Skipped (insufficient firms)")
        continue
    
    mu, Sigma, aligned_isins = estimate_moments(inv_set, ri_m_returns, 
                                                valid_parsed, year_end)
    
    alpha_mv = solve_min_variance(Sigma, len(aligned_isins))
    
    portfolio_data[year_end] = {
        'isins': aligned_isins,
        'weights': alpha_mv,
        'mu': mu,
        'Sigma': Sigma
    }
    
    print(f"      Optimized: {len(aligned_isins)} firms, max weight: {alpha_mv.max():.4f}")

# ============================================================
# 8. COMPUTE PORTFOLIO RETURNS
# ============================================================

print("\n[6/8] Computing portfolio returns...")

def compute_portfolio_returns(portfolio_data, returns_matrix, all_dates):
    """Compute monthly returns for minimum variance portfolio."""
    portfolio_returns = []
    dates = []
    
    for year in sorted(portfolio_data.keys()):
        year_start = pd.Timestamp(f'{year + 1}-01-01')
        year_end = pd.Timestamp(f'{year + 1}-12-31')
        
        year_months = [d for d in all_dates if year_start <= d <= year_end]
        
        isins = portfolio_data[year]['isins']
        weights = portfolio_data[year]['weights']
        
        # Starting weights
        current_weights = weights.copy()
        
        for month_date in year_months:
            if month_date not in returns_matrix.columns:
                continue
            
            # Get returns for this month
            monthly_rets = returns_matrix.loc[isins, month_date].fillna(0).values
            
            # Portfolio return
            port_ret = np.dot(current_weights, monthly_rets)
            portfolio_returns.append(port_ret)
            dates.append(month_date)
            
            # Update weights (drift)
            current_weights = current_weights * (1 + monthly_rets) / (1 + port_ret)
    
    return pd.Series(portfolio_returns, index=dates)

mv_returns = compute_portfolio_returns(portfolio_data, ri_m_returns, valid_parsed)
print(f"   MV portfolio returns: {len(mv_returns)} months")

# ============================================================
# 9. VALUE-WEIGHTED BENCHMARK
# ============================================================

print("\n[7/8] Computing value-weighted benchmark...")

def compute_vw_returns(returns_matrix, mv_matrix, all_dates, start_year=2014, end_year=2025):
    """Compute value-weighted portfolio returns."""
    vw_returns = []
    dates = []
    
    start_date = pd.Timestamp(f'{start_year}-01-01')
    end_date = pd.Timestamp(f'{end_year}-12-31')
    
    relevant_dates = [d for d in all_dates if start_date <= d <= end_date]
    
    for month_date in relevant_dates:
        if month_date not in returns_matrix.columns or month_date not in mv_matrix.columns:
            continue
        
        # Previous month for market caps
        prev_dates = [d for d in all_dates if d < month_date]
        if not prev_dates:
            continue
        prev_date = max(prev_dates)
        
        if prev_date not in mv_matrix.columns:
            continue
        
        # Market caps at t-1
        mv_prev = mv_matrix[prev_date].dropna()
        
        # Returns at t
        ret_current = returns_matrix[month_date]
        
        # Align
        common_isins = mv_prev.index.intersection(ret_current.index)
        if len(common_isins) == 0:
            continue
        
        mv_aligned = mv_prev.loc[common_isins]
        ret_aligned = ret_current.loc[common_isins].fillna(0)
        
        # Value weights
        weights = mv_aligned / mv_aligned.sum()
        
        # Portfolio return
        vw_ret = np.dot(weights, ret_aligned)
        vw_returns.append(vw_ret)
        dates.append(month_date)
    
    return pd.Series(vw_returns, index=dates)

vw_returns = compute_vw_returns(ri_m_returns, mv_m, valid_parsed)
print(f"   VW portfolio returns: {len(vw_returns)} months")

# ============================================================
# 10. PERFORMANCE METRICS
# ============================================================

print("\n[8/8] Computing performance metrics...")

def compute_performance_metrics(returns_series, name="Portfolio"):
    """Compute annualized performance metrics."""
    # Annualized return
    mean_monthly = returns_series.mean()
    ann_return = (1 + mean_monthly) ** 12 - 1
    
    # Annualized volatility
    std_monthly = returns_series.std()
    ann_vol = std_monthly * np.sqrt(12)
    
    # Sharpe ratio (assuming risk-free rate = 0)
    sharpe = ann_return / ann_vol if ann_vol > 0 else 0
    
    # Min and Max monthly returns
    min_ret = returns_series.min()
    max_ret = returns_series.max()
    
    # Cumulative return
    cum_ret = (1 + returns_series).prod() - 1
    
    print(f"\n   {name}:")
    print(f"      Annualized Return: {ann_return*100:.2f}%")
    print(f"      Annualized Volatility: {ann_vol*100:.2f}%")
    print(f"      Sharpe Ratio: {sharpe:.4f}")
    print(f"      Min Monthly Return: {min_ret*100:.2f}%")
    print(f"      Max Monthly Return: {max_ret*100:.2f}%")
    print(f"      Cumulative Return: {cum_ret*100:.2f}%")
    
    return {
        'Annualized Return': ann_return,
        'Annualized Volatility': ann_vol,
        'Sharpe Ratio': sharpe,
        'Min Monthly Return': min_ret,
        'Max Monthly Return': max_ret,
        'Cumulative Return': cum_ret
    }

mv_stats = compute_performance_metrics(mv_returns, "Minimum Variance Portfolio")
vw_stats = compute_performance_metrics(vw_returns, "Value-Weighted Portfolio")

# ============================================================
# 11. EXPORT RESULTS
# ============================================================

print("\n" + "=" * 60)
print("EXPORTING RESULTS")
print("=" * 60)

# Align return series
common_dates = mv_returns.index.intersection(vw_returns.index)
mv_returns_aligned = mv_returns.loc[common_dates]
vw_returns_aligned = vw_returns.loc[common_dates]

# Create output DataFrame
output = pd.DataFrame({
    'Date': common_dates,
    'MV_Return': mv_returns_aligned.values,
    'VW_Return': vw_returns_aligned.values
})

# Compute cumulative returns
output['MV_CumReturn'] = (1 + output['MV_Return']).cumprod()
output['VW_CumReturn'] = (1 + output['VW_Return']).cumprod()

# Save to CSV
output.to_csv('part1_results.csv', index=False)
print("\n✅ Saved: part1_results.csv")

# Save summary statistics
summary = pd.DataFrame({
    'Minimum Variance': mv_stats,
    'Value Weighted': vw_stats
}).T

summary.to_csv('part1_summary_statistics.csv')
print("✅ Saved: part1_summary_statistics.csv")

# Save portfolio compositions
portfolio_compositions = []
for year in sorted(portfolio_data.keys()):
    df = pd.DataFrame({
        'Year': year + 1,
        'ISIN': portfolio_data[year]['isins'],
        'Weight': portfolio_data[year]['weights']
    })
    portfolio_compositions.append(df)

portfolio_comp_df = pd.concat(portfolio_compositions, ignore_index=True)
portfolio_comp_df.to_csv('part1_portfolio_compositions.csv', index=False)
print("✅ Saved: part1_portfolio_compositions.csv")

print("\n" + "=" * 60)
print("PART I COMPLETE")
print("=" * 60)
print("\nNext steps:")
print("1. Review the output CSV files")
print("2. Fill in 'Template for Part I-SAAM.xlsx' with the results")
print("3. Create visualizations (cumulative returns plot)")
