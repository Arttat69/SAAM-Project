# SAAM PART I - QUICK START GUIDE
# Run this after setting up all files in the same directory

"""
FILES NEEDED IN YOUR WORKING DIRECTORY:
========================================
1. Static_2025.xlsx
2. DS_RI_T_USD_Y_2025-3.xlsx
3. DS_RI_T_USD_M_2025.xlsx
4. DS_MV_T_USD_M_2025-2.xlsx
5. saam_part1_complete.py (this script)
6. saam_part1_visualize.py (visualization script)

WHAT THIS SCRIPT DOES:
======================
✓ Filters 513 Pacific firms from 2545 total firms
✓ Cleans price data (removes RI < 0.5, handles delistings)
✓ Builds rolling 10-year investment sets (2014-2025)
✓ Solves minimum variance optimization annually
✓ Computes value-weighted benchmark
✓ Calculates performance metrics
✓ Exports results to CSV

EXPECTED OUTPUTS:
=================
→ part1_results.csv (144 months of returns)
→ part1_summary_statistics.csv (performance metrics)
→ part1_portfolio_compositions.csv (yearly weights)

EXPECTED PERFORMANCE (rough estimates):
=======================================
Minimum Variance Portfolio:
  - Annualized Return: 8-12%
  - Annualized Volatility: 12-18%
  - Sharpe Ratio: 0.4-0.8

Value-Weighted Portfolio:
  - Annualized Return: 7-11%
  - Annualized Volatility: 14-20%
  - Sharpe Ratio: 0.3-0.7

RUNTIME:
========
Total: 5-10 minutes
  - Data loading: 2-3 min
  - Portfolio construction: 2-4 min
  - Performance calculation: 1-2 min

TROUBLESHOOTING:
================
❌ Memory Error → Close other programs, your files are large (~6GB)
❌ Optimization Warning → Normal, check that weights sum to 1.0
❌ File Not Found → Check all 4 Excel files are in same folder

WHAT TO DO AFTER RUNNING:
==========================
1. Run: python saam_part1_complete.py
2. Check console for any warnings
3. Verify 3 CSV files were created
4. Run: python saam_part1_visualize.py
5. Check plots look reasonable
6. Run: python fill_template.py (auto-fills the Excel template!)
7. Review Template-Part-I-FILLED.xlsx
8. Submit by April 12 midnight

VALIDATION CHECKLIST:
=====================
□ part1_results.csv has 144 rows (12 months × 12 years)
□ All returns are between -1.0 and 5.0 (no crazy outliers)
□ Cumulative returns end positive (MV_CumReturn, VW_CumReturn > 1)
□ Summary statistics show MV has lower volatility than VW
□ Portfolio weights sum to 1.0 each year
□ No negative weights in portfolio compositions

NEXT STEPS (AFTER PART I):
===========================
For Part II (Due May 29):
  - Add DS_CO2_SCOPE1.xlsx (Scope 1 emissions)
  - Add DS_REV_USD_Y.xlsx (Revenues)
  - Add DS_MV_T_USD_Y.xlsx (Annual market caps)
  - Compute carbon intensity (CI = CO2 / Revenue)
  - Optimize with 50% carbon reduction constraint
  - Compare P^(mv) vs P^(mv)(0.5) vs P^(vw)(0.5) vs P^(vw)(NZ)

DEADLINE REMINDER:
==================
📅 April 12, 2026 - Preliminary Results (Part I)
📅 April 14, 2026 - In-class Presentation
📅 May 29, 2026 - Final Submission (Part I + Part II)

Good luck! 🚀
"""

# Quick installation check
print("=" * 60)
print("SAAM PROJECT - PART I QUICK START")
print("=" * 60)

import sys

required_packages = {
    'pandas': 'pandas',
    'numpy': 'numpy', 
    'scipy': 'scipy',
    'matplotlib': 'matplotlib',
    'seaborn': 'seaborn',
    'openpyxl': 'openpyxl'
}

missing = []
for package, pip_name in required_packages.items():
    try:
        __import__(package)
        print(f"✓ {package}")
    except ImportError:
        print(f"✗ {package} - MISSING")
        missing.append(pip_name)

if missing:
    print("\n" + "=" * 60)
    print("MISSING PACKAGES - RUN THIS COMMAND:")
    print("=" * 60)
    print(f"pip install {' '.join(missing)}")
    print("=" * 60)
    sys.exit(1)
else:
    print("\n" + "=" * 60)
    print("All packages installed! ✓")
    print("=" * 60)

# File check
import os

required_files = [
    'Static_2025.xlsx',
    'DS_RI_T_USD_Y_2025-3.xlsx',
    'DS_RI_T_USD_M_2025.xlsx',
    'DS_MV_T_USD_M_2025-2.xlsx',
    'saam_part1_complete.py'
]

print("\nCHECKING FILES:")
print("=" * 60)

missing_files = []
for file in required_files:
    if os.path.exists(file):
        size_mb = os.path.getsize(file) / (1024*1024)
        print(f"✓ {file} ({size_mb:.1f} MB)")
    else:
        print(f"✗ {file} - MISSING")
        missing_files.append(file)

if missing_files:
    print("\n" + "=" * 60)
    print("MISSING FILES - PLEASE ADD:")
    print("=" * 60)
    for file in missing_files:
        print(f"  - {file}")
    print("=" * 60)
    sys.exit(1)
else:
    print("\n" + "=" * 60)
    print("All files present! ✓")
    print("=" * 60)

print("\n" + "=" * 60)
print("READY TO RUN!")
print("=" * 60)
print("\nExecute:")
print("  python saam_part1_complete.py")
print("\nThen:")
print("  python saam_part1_visualize.py")
print("\n" + "=" * 60)
