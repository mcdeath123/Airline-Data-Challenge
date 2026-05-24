# U.S. Airline Market Entry Strategy & Data Analytics Pipeline
> **1Q2019 Aviation Market Analysis & Fleet Investment Recommendations**

---

## ✦ Project Overview

This repository contains the complete end-to-end data analytics pipeline, financial models, executive report, and premium interactive presentations for a U.S. domestic market entry strategy. 

The objective is to establish a highly competitive domestic airline utilizing an initial fleet of **5 aircraft** (backed by a **$90,000,000 upfront capital expenditure per aircraft** or $450 million in total portfolio CapEx). Anchored on our brand motto **"On time, for you,"** our investment strategy merges high-liquidity transcontinental corridors and high-yield regional trunks with real-world operational delay risk analysis to maximize yield and accelerate capital amortization.

---

## ✦ Recommended Widescreen Route Portfolio

Our analysis recommends allocating the 5 dedicated aircraft to the following high-performing round-trip routes, generating a total of **$502.2M in net profit** in the first quarter of operations:

| Dedicated Aircraft | Route Segment | Full Airport Locations Served | 1Q2019 Net Profit | Avg Round-Trip Fare | Operating Profit / RT | Breakeven (RTs) | projected payback |
| :---: | :---: | :--- | :---: | :---: | :---: | :---: | :---: |
| **Plane 1** | **JFK - LAX** | New York (JFK) &rarr; Los Angeles (LAX) | **$195.6M** | $964.35 | $61,902 | 1,454 RTs | **41 Days** |
| **Plane 2** | **EWR - SFO** | Newark (EWR) &rarr; San Francisco (SFO) | **$82.8M** | $1,050.82 | $68,324 | 1,317 RTs | **98 Days** |
| **Plane 3** | **JFK - SFO** | New York (JFK) &rarr; San Francisco (SFO) | **$82.8M** | $860.38 | $44,503 | 2,022 RTs | **98 Days** |
| **Plane 4** | **DCA - ORD** | Washington Reagan (DCA) &rarr; Chicago (ORD) | **$72.4M** | $535.33 | $39,213 | 2,295 RTs | **112 Days** |
| **Plane 5** | **ATL - CLT** | Atlanta (ATL) &rarr; Charlotte (CLT) | **$68.6M** | $508.35 | $44,602 | 2,018 RTs | **118 Days** |

### ✦ Strategic Highlights:
* **The JFK-LAX Cash Cow:** Undisputed premium transcontinental segment generating nearly 40% of our portfolio's profit. High passenger density amortizes its dedicated aircraft in only **41 operational days**.
* **High-Yield Tech Corridor (EWR-SFO):** Tech-corporate channel carrying the nation's highest average fare ($1,050.82) and generating the largest profit per individual flight segment ($68k/RT).
* **The Regional Powerhouse (ATL-CLT):** Highly resilient short-haul route bypassing normal distance-fare limits with premium corporate fares ($508.35) and exceptionally low operational overhead ($2,074 per leg).
* **Strategic Rejection of CLT-FLO:** Explicitly bypassed the regional monopoly route Charlotte (CLT) to Florence (FLO) despite high raw ticket fares. Operating 200-seat mainline aircraft in a tiny regional market (averaging 3 flights/day) is structurally unviable and would collapse local ticket pricing.

---

## ✦ Key Deliverables Map

The repository is organized into a clean, platform-independent structure:

* **[Airline_Analysis_Presentation.html](Airline_Analysis_Presentation.html)**: A premium, highly interactive dark-mode HTML slide deck featuring smooth fade-in-up micro-animations, keyboard arrow navigation, and a quick slide jumper dropdown.
* **[Airline_Analysis_Presentation_New.pptx](Airline_Analysis_Presentation_New.pptx)**: Widescreen native PowerPoint presentation deck built programmatically using custom slide designs and metric cards.
* **[Airline_Analysis_Report.md](Airline_Analysis_Report.md)**: A publication-quality executive Markdown report providing comprehensive data cleaning profiles, financial methodology, and future strategic KPIs.
* **[Airline_Analysis_Results.xlsx](Airline_Analysis_Results.xlsx)**: Comprehensive calculations workbook featuring five tabs:
  1. `Abbreviation Definitions` (Fully styled metadata dictionary and IATA code index)
  2. `Busiest Routes`
  3. `Most Profitable Routes`
  4. `Recommended Routes`
  5. `All Routes Summary`
* **`airline_analysis.py`**: The core data analytics execution pipeline.
* **[scripts/](scripts/)**: Folder containing modular utility scripts:
  * `airline_analysis.py`: Main analytical pipeline.
  * `add_abbreviation_sheet.py`: Styles and prepends the metadata sheet.
  * `generate_presentation.py`: Builds the native PowerPoint presentation.
* **[visualizations/](visualizations/)**: Four publication-quality high-resolution Matplotlib charts:
  * `chart_1_busiest_routes.png`: Top 10 Busiest Round-Trip Routes.
  * `chart_2_profitable_routes.png`: Grouped financial component comparison.
  * `chart_3_breakeven_analysis.png`: Payback speed for recommended aircraft.
  * `chart_4_profitability_drivers.png`: Load Factor vs. Ticket Fare bubble chart.

---

## ✦ Data Analytics & Financial Methodology

### 1. Robust Data Cleaning
Our pipeline processes over **1.9M flight records** and **1.1M ticket records** for 1Q2019 with zero errors:
* **Frequent Flyer Outliers:** Excluded all fares **under $50** (eliminating reward tickets and security fee distortions) and **above $5,000** (eliminating data corruptions).
* **Missing Distance Imputation:** Imputed 2,700 records with missing or corrupted distances using the **median distance** of all other flights on the exact same directed route.
* **Missing Arrival Delays:** Imputed diverted flights (valid departure delay but null arrival delay) using departure delay (`DEP_DELAY`) to preserve traffic volumes and delay penalties.
* **Mixed-Type Formatting:** Cleansed currency string prefixes, extra spaces, and special symbols (e.g. `'$$$'`) and coerced them to clean numeric float values.

### 2. Financial Modeling Formulas
* **Mainline Passenger Load:** `Passengers = 200 * Occupancy Rate` (representing mainline aircraft scaled by real-world route occupancies).
* **Ticket Revenue:** `Revenue = Passengers * (Average Round-Trip Fare / 2) per leg`.
* **Checked Bag Revenue:** `Bag Revenue = Passengers * 50% bag check rate * $35.00 leg bag fee`.
* **Physical Distance Cost:** `Cost = Distance * $9.18 per mile` (covers fuel, crew, maintenance at $8.00/mi + depreciation, insurance at $1.18/mi).
* **Landing Airport Costs:** `$10,000 per arrival leg for Large airports` and `$5,000 for Medium airports` (*source: Airport_Codes.csv*).
* **Operational Delay Cost:** `Delay Cost = (Minutes - 15) * $75.00 per minute` for arrival and departure delays exceeding 15 minutes.

---

## ✦ Execution Instructions

### Prerequisites
Ensure you have Python 3.8+ installed along with the required libraries:
```bash
pip install pandas numpy openpyxl matplotlib python-pptx
```

### Running the Data Pipeline
To execute the complete end-to-end analytical pipeline and regenerate all charts and Excel sheets:
```bash
python airline_analysis.py
```

### Formatting the Excel Metadata Sheet
To format the workbook and prepend the styled `Abbreviation Definitions` metadata tab:
```bash
python scripts/add_abbreviation_sheet.py
```

### Generating the Widescreen PowerPoint Presentation
To build the native widescreen `.pptx` presentation slides:
```bash
python scripts/generate_presentation.py
```

---

## ✦ Future Strategic Roadmap & KPIs

To transition our airline from market entry to sustainable scaling, we recommend tracking the following **Key Performance Indicators (KPIs)**:
1. **On-Time Performance (OTP - Target 90%+):** Punctuality is the direct guardian of operating margins. Since delays over 15 minutes cost **$75/minute**, keeping turnaround times tight directly protects profits.
2. **Route Load Factor (Target 70%+):** Occupancy is our primary scale scaler. Adding just 1% in load factor on the JFK-LAX route yields **$1.5M/quarter** in high-margin revenue.
3. **Yield per Available Seat Mile (YASM):** Total route revenue divided by Available Seat Miles (`200 seats * Distance`). Measures pricing efficiency and isolates premium transcontinental margins.
