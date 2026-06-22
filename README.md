## F1 Championship Analysis — Project Summary

**What I did:**
I analyzed 21 seasons (2004–2024) of real Formula 1 data from Kaggle's F1 World Championship dataset to answer one question: *what actually determines who wins a championship?* I worked across 14 raw CSV files (results, standings, qualifying, pit stops, status codes, etc.), built a structured Excel workbook with four research questions, and visualized the findings in an interactive dashboard.

**The 4 questions I investigated:**
1. **Car vs. driver** — Does the best constructor always produce the champion driver?
2. **Reliability** — How much do mechanical failures (DNFs) cost a championship?
3. **Qualifying vs. race pace** — Is Saturday or Sunday more predictive of winning?
4. **Pit strategy** — Do pit stop times create a measurable competitive edge?

**Key findings:**
- **Car performance dominates** — 86% of seasons saw the same team produce both the constructors' and drivers' champion. The 3 exceptions (2008, 2021, 2024) were among the most dramatic title fights in F1 history.
- **Reliability decides close championships** — DNFs were the deciding factor in several razor-thin title races (Alonso lost both 2010 and 2012 by single-digit point margins).
- **Race execution beats qualifying** — Hamilton won the 2019 title with only 5 poles but 11 wins, proving conversion matters more than starting position.
- **Operational details compound** — A 2.2-second pit stop gap between the fastest and slowest 2023 teams adds up to nearly two minutes of lost time across a season.

I also did a follow-up exercise classifying all 77 historical F1 circuits into street/high-speed/mixed/technical categories, then cross-checked that against the live 2026 calendar to separate active venues from retired ones — turning a static dataset into something tied to the current state of the sport.

**What I learned:**
- **Data wrangling** — joining multiple relational tables (races, results, standings, status codes) to answer business questions, not just describe data
- **Defining metrics from ambiguous concepts** — "DNF-decisive" and "pole-to-win conversion" had to be operationalized with clear, defensible rules
- **Translating data into a narrative** — the real BA skill wasn't running the numbers, it was turning them into a "what should you actually believe" conclusion
- **Knowing when data is stale** — recognizing the dataset stopped at 2024 and verifying current information rather than assuming it was still accurate

---

*Built with Python · 2026*
