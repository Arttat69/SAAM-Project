"""
SAAM Project — Part 2: Plotting Script (v3)
Reads from resultsPart1/ and ResultsPart2[_LW]/, saves all figures there.

Set USE_LEDOIT_WOLF to match what you used in part2_main_v3.py.
Run AFTER part2_main_v3.py.
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# ─── Match this flag to part2_main_v3.py ─────────────────────────────────────
USE_LEDOIT_WOLF = True    # ← flip to False if you ran without LW shrinkage

RESULTS_PART1 = "resultsPart1"
RESULTS_PART2 = "ResultsPart2_LW" if USE_LEDOIT_WOLF else "ResultsPart2"
os.makedirs(RESULTS_PART2, exist_ok=True)

Y0    = 2013
THETA = 0.10

COLORS = {"MV": "#1f77b4", "MV05": "#ff7f0e",
          "VW": "#2ca02c", "TE05": "#d62728", "NZ": "#9467bd"}

label_suffix = " (LW)" if USE_LEDOIT_WOLF else ""

# ── Load return series ────────────────────────────────────────────────────────
ret_df      = pd.read_csv(f"{RESULTS_PART1}/part1_results.csv",
                           parse_dates=["Date"]).set_index("Date")
mv_port_ret = ret_df["MV_Return"]
vw_port_ret = ret_df["VW_Return"]

ret_32 = pd.read_csv(f"{RESULTS_PART2}/returns_32_mv_carbon05.csv",
                     index_col=0, parse_dates=True).squeeze("columns")
ret_33 = pd.read_csv(f"{RESULTS_PART2}/returns_33_te_carbon05.csv",
                     index_col=0, parse_dates=True).squeeze("columns")
ret_41 = pd.read_csv(f"{RESULTS_PART2}/returns_41_netzero.csv",
                     index_col=0, parse_dates=True).squeeze("columns")

carbon  = pd.read_csv(f"{RESULTS_PART2}/all_carbon_metrics.csv",  index_col=0)
nz_path = pd.read_csv(f"{RESULTS_PART2}/nz_target_path.csv",      index_col=0).squeeze("columns")
top10   = pd.read_csv(f"{RESULTS_PART2}/top10_carbon_intensity.csv")


def cum_ret(r):
    return (1.0 + r).cumprod()

def fmt_xaxis(ax, rot=45):
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    plt.setp(ax.get_xticklabels(), rotation=rot, ha="right")

def save(fig, name):
    path = f"{RESULTS_PART2}/{name}"
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved {path}")


# ── Fig 1: Cumulative returns — MV vs MV(0.5) ────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
cum_ret(mv_port_ret).plot(ax=ax, color=COLORS["MV"],  lw=2, ls="-",
                          label=r"$P^{(mv)}_{oos}$")
cum_ret(ret_32).plot(     ax=ax, color=COLORS["MV05"], lw=2, ls="--",
                          label=r"$P^{(mv)}_{oos}(0.5)$" + label_suffix)
ax.set_title("Fig. 1 — Cumulative Returns: MV vs. MV with 50% CF Reduction", fontsize=12)
ax.set_ylabel("Cumulative Return (base = 1.0)"); ax.legend(); fmt_xaxis(ax); ax.grid(alpha=.3)
save(fig, "fig1_cumret_mv_vs_mv05.png")

# ── Fig 2: Cumulative returns — VW vs TE(0.5) ────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
cum_ret(vw_port_ret).plot(ax=ax, color=COLORS["VW"],   lw=2, ls="-",
                          label=r"$P^{(vw)}_{oos}$")
cum_ret(ret_33).plot(     ax=ax, color=COLORS["TE05"], lw=2, ls="--",
                          label=r"$P^{(vw)}_{oos}(0.5)$" + label_suffix)
ax.set_title("Fig. 2 — Cumulative Returns: VW vs. TE with 50% CF Reduction", fontsize=12)
ax.set_ylabel("Cumulative Return (base = 1.0)"); ax.legend(); fmt_xaxis(ax); ax.grid(alpha=.3)
save(fig, "fig2_cumret_vw_vs_te05.png")

# ── Fig 3: Cumulative returns — VW vs TE(0.5) vs NZ ─────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
cum_ret(vw_port_ret).plot(ax=ax, color=COLORS["VW"],   lw=2, ls="-",
                          label=r"$P^{(vw)}_{oos}$")
cum_ret(ret_33).plot(     ax=ax, color=COLORS["TE05"], lw=2, ls="--",
                          label=r"$P^{(vw)}_{oos}(0.5)$" + label_suffix)
cum_ret(ret_41).plot(     ax=ax, color=COLORS["NZ"],   lw=2, ls=":",
                          label=r"$P^{(vw)}_{oos}(NZ)$" + label_suffix)
ax.set_title("Fig. 3 — Cumulative Returns: VW vs. TE(0.5) vs. Net Zero", fontsize=12)
ax.set_ylabel("Cumulative Return (base = 1.0)"); ax.legend(); fmt_xaxis(ax); ax.grid(alpha=.3)
save(fig, "fig3_cumret_vw_te05_nz.png")

# ── Fig 4: WACI evolution ─────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
years = carbon.index.astype(int)
for col, lbl, c, ls in [
    ("WACI_MV",    r"$P^{(mv)}_{oos}$",      COLORS["MV"],   "-"),
    ("WACI_VW",    r"$P^{(vw)}_{oos}$",      COLORS["VW"],   "-"),
    ("WACI_MV_05", r"$P^{(mv)}_{oos}(0.5)$" + label_suffix, COLORS["MV05"], "--"),
    ("WACI_TE_05", r"$P^{(vw)}_{oos}(0.5)$" + label_suffix, COLORS["TE05"], "--"),
    ("WACI_NZ",    r"$P^{(vw)}_{oos}(NZ)$"  + label_suffix, COLORS["NZ"],   ":"),
]:
    ax.plot(years, carbon[col], ls=ls, lw=2, marker="o", ms=4, color=c, label=lbl)
ax.set_title("Fig. 4 — Weighted-Average Carbon Intensity (WACI)", fontsize=12)
ax.set_ylabel(r"WACI (tCO$_2$ / M USD Revenue)"); ax.set_xlabel("Year")
ax.legend(fontsize=9); ax.grid(alpha=.3)
save(fig, "fig4_waci_evolution.png")

# ── Fig 5: Carbon Footprint evolution + NZ target path ───────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
for col, lbl, c, ls, mk in [
    ("CF_MV",    r"$P^{(mv)}_{oos}$",      COLORS["MV"],   "-",  "o"),
    ("CF_VW",    r"$P^{(vw)}_{oos}$",      COLORS["VW"],   "-",  "s"),
    ("CF_MV_05", r"$P^{(mv)}_{oos}(0.5)$" + label_suffix, COLORS["MV05"], "--", "^"),
    ("CF_TE_05", r"$P^{(vw)}_{oos}(0.5)$" + label_suffix, COLORS["TE05"], "--", "v"),
    ("CF_NZ",    r"$P^{(vw)}_{oos}(NZ)$"  + label_suffix, COLORS["NZ"],   ":",  "D"),
]:
    ax.plot(years, carbon[col], ls=ls, lw=2, marker=mk, ms=4, color=c, label=lbl)
ax.plot(nz_path.index.astype(int), nz_path.values, ls="--", lw=1.5,
        color="black", alpha=.6, label="NZ target path")
ax.set_title("Fig. 5 — Carbon Footprint & Net-Zero Target Path", fontsize=12)
ax.set_ylabel(r"CF (tCO$_2$ / M USD Invested)"); ax.set_xlabel("Year")
ax.legend(fontsize=9); ax.grid(alpha=.3)
save(fig, "fig5_cf_evolution.png")

# ── Fig 6: Top-10 firms by average carbon intensity ──────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
labels = [f"{str(row['Name'])[:22]}\n({row['ISIN']})" for _, row in top10.iterrows()]
ax.barh(labels[::-1], top10["Avg_CI"].values[::-1],
        color="#d62728", alpha=0.8, edgecolor="white")
ax.set_xlabel(r"Average CI (tCO$_2$ / M USD Revenue)")
ax.set_title("Fig. 6 — Top 10 Firms by Average Carbon Intensity (2013–2024)", fontsize=12)
ax.grid(axis="x", alpha=.3)
save(fig, "fig6_top10_ci.png")

# ── Fig 7: Rolling 12-month Sharpe (VW-based strategies) ─────────────────────
def rolling_sharpe(r, w=12):
    mu = r.rolling(w).mean() * 12
    sd = r.rolling(w).std(ddof=1) * (12 ** .5)
    return mu / sd

fig, ax = plt.subplots(figsize=(10, 5))
for ret, lbl, c, ls in [
    (vw_port_ret, r"$P^{(vw)}_{oos}$",                          COLORS["VW"],   "-"),
    (ret_33,      r"$P^{(vw)}_{oos}(0.5)$" + label_suffix,      COLORS["TE05"], "--"),
    (ret_41,      r"$P^{(vw)}_{oos}(NZ)$"  + label_suffix,      COLORS["NZ"],   ":"),
]:
    rolling_sharpe(ret).plot(ax=ax, label=lbl, color=c, lw=2, ls=ls)
ax.axhline(0, color="black", lw=.8, alpha=.5)
ax.set_title("Fig. 7 — Rolling 12-Month Sharpe Ratio (VW-based portfolios)", fontsize=12)
ax.set_ylabel("Rolling Sharpe Ratio"); fmt_xaxis(ax); ax.legend(); ax.grid(alpha=.3)
save(fig, "fig7_rolling_sharpe.png")

# ── Fig 8 (LW only): Ledoit-Wolf shrinkage intensity δ̂ over time ─────────────
if USE_LEDOIT_WOLF and "LW_delta" in carbon.columns:
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    axes[0].bar(years, carbon["LW_delta"], color="#1f77b4", alpha=0.8, edgecolor="white")
    axes[0].set_title(r"LW Shrinkage Intensity $\hat{\delta}$ by Year", fontsize=12)
    axes[0].set_ylabel(r"$\hat{\delta}$"); axes[0].set_xlabel("Year")
    axes[0].axhline(0, color="black", lw=0.8)
    axes[0].grid(axis="y", alpha=.3)

    axes[1].bar(years, carbon["LW_r_bar"], color="#ff7f0e", alpha=0.8, edgecolor="white")
    axes[1].set_title(r"Average Sample Correlation $\bar{r}$ by Year", fontsize=12)
    axes[1].set_ylabel(r"$\bar{r}$"); axes[1].set_xlabel("Year")
    axes[1].grid(axis="y", alpha=.3)
    save(fig, "fig8_lw_parameters.png")

print(f"\nAll figures saved to '{RESULTS_PART2}/'")
