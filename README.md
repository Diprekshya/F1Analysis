# What Wins F1 Championships?
### A Data Analysis of 21 F1 Seasons (2004–2024)

> **Portfolio project** — Business Analysis using real Ergast F1 data, Python, pandas/matplotlib, and Excel (openpyxl).

---

## Key Findings at a Glance

| Question | Finding |
|---|---|
| **Q1 — Does the best car always win?** | 86% match rate (18/21 seasons). The 3 exceptions (2008, 2021, 2024) are the most dramatic seasons in modern F1. |
| **Q2 — How much do DNFs cost titles?** | Champions averaged 1.3 DNFs vs 1.9 for runners-up. In 5 seasons, DNFs were mathematically championship-decisive. |
| **Q3 — Qualifying or race pace?** | Race execution outweighs Saturday qualifying. Hamilton 2019: 5 poles → 11 wins. Alonso 2013: 0 poles, still P2 in WDC. |
| **Q4 — Do pit stops decide races?** | 2.2s gap between fastest/slowest teams ≈ 97 seconds of cumulative pit time advantage across a full season. |
| **Q5 — Circuit DNA?** | Every constructor has a comfort zone — Red Bull dominates high-speed, Ferrari excels on technical layouts, McLaren's 2024 surge was fast-corner focused. |

---

## Project Structure

```
f1-analysis/
├── index.html              # GitHub Pages portfolio site
├── analyse.py              # Python analysis → generates charts + Excel
├── F1_Championship_Analysis.xlsx  # 5-sheet Excel workbook (auto-generated)
├── charts/                 # 12 PNG charts (auto-generated)
│   ├── q1_match_bar.png
│   ├── q1_exceptions.png
│   ├── q2_dnf_analysis.png
│   ├── q2_decisive.png
│   ├── q3_grid_vs_finish.png
│   ├── q3_poles_vs_wins.png
│   ├── q3_positions_gained.png
│   ├── q4_pit_avg.png
│   ├── q4_pit_trend.png
│   ├── q5_circuit_heatmap.png
│   ├── q5_circuit_profile.png
│   └── q5_circuit_normalised.png
└── data/                   # Raw Kaggle CSVs (not committed — see below)
    ├── results.csv
    ├── drivers.csv
    ├── constructors.csv
    ├── races.csv
    ├── driver_standings.csv
    ├── constructor_standings.csv
    ├── pit_stops.csv
    ├── qualifying.csv
    ├── circuits.csv
    └── status.csv
```

---

## Running the Analysis

**Requirements:** Python 3.8+

```bash
# Install dependencies
pip install pandas matplotlib seaborn openpyxl numpy

# Run analysis (generates all charts + Excel)
python3 analyse.py
```

This will:
1. Load 14 CSV files from the `data/` folder
2. Run Q1–Q5 analysis (2004–2024)
3. Export 12 PNG charts to `charts/`
4. Export a 5-sheet Excel workbook with data tables, colour-coded analysis, and insight boxes

---

## Data Sources

All data from the **Ergast Motor Racing Database** via [Kaggle F1 World Championship Dataset](https://www.kaggle.com/datasets/rohanrao/formula-1-world-championship-1950-2020).

14 CSV files used: `results`, `drivers`, `constructors`, `races`, `driver_standings`, `constructor_standings`, `pit_stops`, `qualifying`, `circuits`, `status`, `seasons`, `constructor_results`, `lap_times`, `sprint_results`.

---

## Methodology Notes

- **Scope:** 2004–2024 (21 seasons of the V8/V6 hybrid turbo era)
- **Q2 points estimate:** DNF cost calculated as driver's average finish position × points scale. Conservative, but consistent.
- **Q4 pit times:** Total pit lane transit time (entry + stationary + exit). Outliers >60s removed (safety car/mechanical). All teams measured on the same basis — cross-team comparisons are valid.
- **Q5 circuit classification:** Heuristic rules based on circuit name/ref — not official FIA categories. Street, High Speed, Slow/Technical, Mixed.
- **Limitations:** Pit stop data only from 2011. Qualifying coverage varies pre-2010. 2024 data may be incomplete depending on dataset version.

---

## Deploying to GitHub Pages

1. Push the repo to GitHub (exclude `data/` if you prefer — charts are pre-generated)
2. Go to **Settings → Pages → Source: main branch / root**
3. Your site will be live at `https://<your-username>.github.io/<repo-name>/`

---

*Built with Python (pandas, matplotlib, seaborn, openpyxl) · Ergast F1 data · 2026*
