# What Wins F1 Championships?
### A Data Analysis of 21 F1 Seasons (2004–2024)

> **Portfolio project** — Business Analysis using real Ergast F1 data, Python, pandas/matplotlib, and Excel.
> **Live Site→** https://diprekshya.github.io/F1Analysis/

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

*Built with Python · Ergast F1 data · 2026*
