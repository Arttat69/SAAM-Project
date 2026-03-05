"""
SAAM Project — Part 2: Plotting Script
Reads from resultsPart1/ and ResultsPart2/, saves all figures to ResultsPart2/.
Run AFTER part2_main.py.
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

RESULTS_PART1 = "resultsPart1"
RESULTS_PART2 = "ResultsPart2"
os.makedirs(RESULTS_PART2, exist_ok=True)

Y0    = 2013
THETA = 0.10

COLORS = {"MV": "#1f77b4", "MV05": "#ff7f0e",
          "VW": "#2ca02c", "TE05": "#d62728", "NZ": "#9467bd"}

# ── Load return series ────────────────────────────────────────────────────────
ret_df = pd.read_csv(f"{RESULTS_PART1}/part1_results.csv", parse_dates=["Date"]).set_index("Date")
mv_port_ret = ret_df["MV_Return"]
vw_port_ret = ret_df["VW_Return"]

ret_32 = pd.read_csv(f"{RESULTS_PART2}/returns_32_mv_carbon05.csv",
                     index_col=0, parse_dates=True).squeeze("columns")
ret_33 = pd.read_csv(f"{RESULTS_PART2}/returns_33_te_carbon05.csv",
                     index_col=0, parse_dates=True).squeeze("columns")
ret_41 = pd.read_csv(f"{RESULTS_PART2}/returns_41_netzero.csv",
                     index_col=0, parse_dates=True).squeeze("columns")

carbon  = pd.read_csv(f"{RESULTS_PART2}/all_carbon_metrics.csv", index_col=0)
nz_path = pd.read_csv(f"{RESULTS_PART2}/nz_target_path.csv", index_col=0).squeeze("columns")
top10   = pd.read_csv(f"{RESULTS_PART2}/top10_carbon_intensity.csv")


def cum_ret(r):
    return (1.0 + r).cumprod()

def fmt_xaxis(ax, rot=45):
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    plt.setp(ax.get_xticklabels(), rotation=rot, ha="right")

# ── Figure 1: Cumulative returns — MV vs MV(0.5) ─────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
cum_ret(mv_port_ret).plot(ax=ax, color=COLORS["MV"],  lw=2, ls="-",
                          label=r"$P^{(mv)}_{oos}$")
cum_ret(ret_32).plot(     ax=ax, color=COLORS["MV05"], lw=2, ls="--",
                          label=r"$P^{(mv)}_{oos}(0.5)$")
ax.set_title("Fig. 1 — Cumulative Returns: MV vs. MV with 50 % CF Reduction", fontsize=12)
ax.set_ylabel("Cumulative Return (base = 1.0)"); ax.legend(); fmt_xaxis(ax); ax.grid(alpha=.3)
plt.tight_layout()
plt.savefig(f"{RESULTS_PART2}/fig1_cumret_mv_vs_mv05.png", dpi=150); plt.close()
print("Saved fig1")

# ── Figure 2: Cumulative returns — VW vs TE(0.5) ─────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
cum_ret(vw_port_ret).plot(ax=ax, color=COLORS["VW"],   lw=2, ls="-",
                          label=r"$P^{(vw)}_{oos}$")
cum_ret(ret_33).plot(     ax=ax, color=COLORS["TE05"], lw=2, ls="--",
                          label=r"$P^{(vw)}_{oos}(0.5)$")
ax.set_title("Fig. 2 — Cumulative Returns: VW vs. TE with 50 % CF Reduction", fontsize=12)
ax.set_ylabel("Cumulative Return (base = 1.0)"); ax.legend(); fmt_xaxis(ax); ax.grid(alpha=.3)
plt.tight_layout()
plt.savefig(f"{RESULTS_PART2}/fig2_cumret_vw_vs_te05.png", dpi=150); plt.close()
print("Saved fig2")

# ── Figure 3: Cumulative returns — VW vs TE(0.5) vs NZ ───────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
cum_ret(vw_port_ret).plot(ax=ax, color=COLORS["VW"],   lw=2, ls="-",
                          label=r"$P^{(vw)}_{oos}$")
cum_ret(ret_33).plot(     ax=ax, color=COLORS["TE05"], lw=2, ls="--",
                          label=r"$P^{(vw)}_{oos}(0.5)$")
cum_ret(ret_41).plot(     ax=ax, color=COLORS["NZ"],   lw=2, ls=":",
                          label=r"$P^{(vw)}_{oos}(NZ)$")
ax.set_title("Fig. 3 — Cumulative Returns: VW vs. TE(0.5) vs. Net Zero", fontsize=12)
ax.set_ylabel("Cumulative Return (base = 1.0)"); ax.legend(); fmt_xaxis(ax); ax.grid(alpha=.3)
plt.tight_layout()
plt.savefig(f"{RESULTS_PART2}/fig3_cumret_vw_te05_nz.png", dpi=150); plt.close()
print("Saved fig3")

# ── Figure 4: WACI evolution ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
years = carbon.index.astype(int)
for col, label, c, ls in [
    ("WACI_MV",    r"$P^{(mv)}_{oos}$",      COLORS["MV"],   "-"),
    ("WACI_VW",    r"$P^{(vw)}_{oos}$",      COLORS["VW"],   "-"),
    ("WACI_MV_05", r"$P^{(mv)}_{oos}(0.5)$", COLORS["MV05"], "--"),
    ("WACI_TE_05", r"$P^{(vw)}_{oos}(0.5)$", COLORS["TE05"], "--"),
    ("WACI_NZ",    r"$P^{(vw)}_{oos}(NZ)$",  COLORS["NZ"],   ":"),
]:
    ax.plot(years, carbon[col], ls=ls, lw=2, marker="o", ms=4, color=c, label=label)
ax.set_title("Fig. 4 — Weighted-Average Carbon Intensity (WACI)", fontsize=12)
ax.set_ylabel(r"WACI (tCO$_2$ / M USD Revenue)"); ax.set_xlabel("Year")
ax.legend(fontsize=9); ax.grid(alpha=.3)
plt.tight_layout()
plt.savefig(f"{RESULTS_PART2}/fig4_waci_evolution.png", dpi=150); plt.close()
print("Saved fig4")

# ── Figure 5: Carbon Footprint evolution + NZ target path ────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
for col, label, c, ls, mk in [
    ("CF_MV",    r"$P^{(mv)}_{oos}$",      COLORS["MV"],   "-",  "o"),
    ("CF_VW",    r"$P^{(vw)}_{oos}$",      COLORS["VW"],   "-",  "s"),
    ("CF_MV_05", r"$P^{(mv)}_{oos}(0.5)$", COLORS["MV05"], "--", "^"),
    ("CF_TE_05", r"$P^{(vw)}_{oos}(0.5)$", COLORS["TE05"], "--", "v"),
    ("CF_NZ",    r"$P^{(vw)}_{oos}(NZ)$",  COLORS["NZ"],   ":",  "D"),
]:
    ax.plot(years, carbon[col], ls=ls, lw=2, marker=mk, ms=4, color=c, label=label)
ax.plot(nz_path.index.astype(int), nz_path.values, ls="--", lw=1.5,
        color="black", alpha=.6, label="NZ target path")
ax.set_title("Fig. 5 — Carbon Footprint & Net-Zero Target Path", fontsize=12)
ax.set_ylabel(r"CF (tCO$_2$ / M USD Invested)"); ax.set_xlabel("Year")
ax.legend(fontsize=9); ax.grid(alpha=.3)
plt.tight_layout()
plt.savefig(f"{RESULTS_PART2}/fig5_cf_evolution.png", dpi=150); plt.close()
print("Saved fig5")

# ── Figure 6: Top-10 firms by average carbon intensity ───────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
labels = [f"{str(row['Name'])[:22]}\n({row['ISIN']})" for _, row in top10.iterrows()]
ax.barh(labels[::-1], top10["Avg_CI"].values[::-1], color="#d62728", alpha=0.8,
        edgecolor="white")
ax.set_xlabel(r"Average CI (tCO$_2$ / M USD Revenue)")
ax.set_title("Fig. 6 — Top 10 Firms by Average Carbon Intensity (2013–2024)", fontsize=12)
ax.grid(axis="x", alpha=.3)
plt.tight_layout()
plt.savefig(f"{RESULTS_PART2}/fig6_top10_ci.png", dpi=150); plt.close()
print("Saved fig6")

# ── Figure 7: Rolling 12-month Sharpe (VW-based strategies) ──────────────────
def rolling_sharpe(r, w=12):
    mu = r.rolling(w).mean() * 12
    sd = r.rolling(w).std(ddof=1) * (12 ** .5)
    return mu / sd

fig, ax = plt.subplots(figsize=(10, 5))
for ret, label, c, ls in [
    (vw_port_ret, r"$P^{(vw)}_{oos}$",      COLORS["VW"],   "-"),
    (ret_33,      r"$P^{(vw)}_{oos}(0.5)$", COLORS["TE05"], "--"),
    (ret_41,      r"$P^{(vw)}_{oos}(NZ)$",  COLORS["NZ"],   ":"),
]:
    rolling_sharpe(ret).plot(ax=ax, label=label, color=c, lw=2, ls=ls)
ax.axhline(0, color="black", lw=.8, alpha=.5)
ax.set_title("Fig. 7 — Rolling 12-Month Sharpe Ratio (VW-based portfolios)", fontsize=12)
ax.set_ylabel("Rolling Sharpe Ratio"); fmt_xaxis(ax); ax.legend(); ax.grid(alpha=.3)
plt.tight_layout()
plt.savefig(f"{RESULTS_PART2}/fig7_rolling_sharpe.png", dpi=150); plt.close()
print("Saved fig7")

print(f"\nAll figures saved to '{RESULTS_PART2}/'")
