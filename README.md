# Stock Market Portfolio Tracker and Return Analyzer

A simple, smart, and interactive Excel dashboard to track stock profits, analyze Indian companies, and manage investments in one place.

---

## Dashboard Preview

![Portfolio Dashboard Screenshot](screenshots/dashboard_screenshot.png)

---

## Why We Built This Project

Most investors buy shares across different trading apps such as Zerodha, Groww, and Angel One. Because the data is scattered across multiple places, it becomes very confusing to know:

1. How much total profit or loss have I made across all my shares?
2. Which company is performing best, and which one is lagging behind?
3. How much free cash am I receiving each year from company dividends?
4. How is my money divided across different industries like IT, Banking, and Railways?

This project solves all these problems by bringing everything together into one clean, automated Excel dashboard.

---

## What This Dashboard Solves

- **Centralized View**: Combines 11 top Indian stocks from 5 different sectors into a single master ledger.
- **Single-Click Dropdown**: Pick any company from the dropdown menu (like HCL Tech, TCS, or RVNL) to see its full profit story, purchase history, and 12-month price chart instantly.
- **Visual Charts Included**: Features a complete set of visual graphs (Donut chart, Pie chart, Bar chart, and Line charts) that explain the data simply without reading complicated tables.
- **Zero Number Clipping**: All columns have generous widths so numbers and currency figures display cleanly with no `###` hashtag errors.
- **100% Native Excel Formulas**: Uses standard `=XLOOKUP`, `=SUMIF`, and `=IF` formulas without needing complex macros or VBA scripts.

---

## Portfolio Breakdown (11 Indian Companies)

The tracker monitors 11 well-known Indian stocks across 5 key sectors:

| Company Name | Stock Code | Sector / Industry | Shares Held | Buy Price (INR) | Current Price (INR) | Total Profit (%) |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| HCL Technologies Ltd | HCLTECH | Information Technology | 150 | 1,120.00 | 1,745.50 | +55.85% |
| Tata Consultancy Services | TCS | Information Technology | 40 | 3,350.00 | 4,120.00 | +22.99% |
| Infosys Limited | INFY | Information Technology | 120 | 1,420.00 | 1,865.00 | +31.34% |
| Wipro Limited | WIPRO | Information Technology | 250 | 410.00 | 545.00 | +32.93% |
| Tech Mahindra Ltd | TECHM | Information Technology | 80 | 1,180.00 | 1,640.00 | +38.98% |
| HDFC Bank Limited | HDFCBANK | Banking & Financials | 100 | 1,490.00 | 1,680.00 | +12.75% |
| Axis Bank Limited | AXISBANK | Banking & Financials | 90 | 960.00 | 1,265.00 | +31.77% |
| Coal India Limited | COALINDIA | Energy & Mining | 300 | 280.00 | 495.00 | +76.79% |
| Rail Vikas Nigam Ltd (RVNL) | RVNL | Railways & Infrastructure | 350 | 175.00 | 565.00 | +222.86% |
| Indian Railway Finance Corp | IRFC | Railways & Infrastructure | 800 | 75.00 | 185.00 | +146.67% |
| Aditya Birla Capital Ltd | ABCAPITAL | Conglomerate & Financials | 400 | 195.00 | 405.00 | +107.69% |

---

## Visual Charts and Graphs

The dashboard includes 5 visual charts that help anyone understand the numbers quickly:

1. **Sector Allocation Donut Chart**: Shows how much money is invested across IT, Banking, Energy, Railways, and Conglomerate sectors.
2. **Asset Share Pie Chart**: Shows which company holds the largest percentage of your total money.
3. **Sector Market Value Bar Chart**: Horizontal bar graph comparing total market value for each sector.
4. **Dynamic 12-Month Area Chart**: Automatically redraws the 1-year price trajectory for whichever stock you select in the dropdown menu.
5. **Top Bluechips Multi-Line Comparison Chart**: Shows a 1-year price comparison between 5 market leaders (TCS, Infosys, HDFC Bank, HCL Tech, and RVNL) on the same graph.
6. **In-Cell Data Bars and Sparklines**: Mini graphs and visual weight bars inside the table cells for quick scanning.

---

## Simple Formulas Used

All calculations are built using clear, standard Excel formulas:

### 1. Market Value
```excel
= Shares_Held * Current_Price
= D6 * E6
```

### 2. Total Profit or Loss
```excel
= Current_Market_Value - Total_Cost_Spent
= J2 - I2
```

### 3. Return on Investment (ROI %)
```excel
= Total_Profit / Total_Cost_Spent
= K2 / I2
```

### 4. Interactive Dropdown Filter
```excel
=IF($C$3="ALL PORTFOLIO (MASTER VIEW)", SUM(F6:F16), XLOOKUP($C$3, B6:B16, F6:F16))
```

### 5. Sector Total
```excel
=SUMIF(Holdings_Data!$D$2:$D$12, A2, Holdings_Data!$J$2:$J$12)
```

### 6. Yearly Dividend Cash
```excel
= Current_Market_Value * Dividend_Yield
= J2 * O2
```

---

## Key Insights Learned from the Data

1. **Big Winners (Railways)**: RVNL grew by **+222.86%** and IRFC grew by **+146.67%** because of major government investments in Indian railway infrastructure.
2. **High Yearly Cash Dividends**: Coal India gave a **5.20% cash dividend** every year on top of a **+76.79% stock price gain**.
3. **Safe Foundation**: Tech giants like TCS, Infosys, and HCL Tech gave steady 23% to 55% profits and protected the portfolio during market downturns.
4. **Strong Annual Growth**: The entire portfolio achieved an estimated **+28.4% annual growth rate (CAGR)**, which is nearly 4 times higher than a regular bank fixed deposit.

---

## Project File Structure

```text
Stock-Market-Portfolio-Tracker/
│
├── Stock_Portfolio_Master_Complete_v2.xlsx    # Main Excel dashboard with all 5 charts and dropdowns
├── Stock_Market_Portfolio_Presentation_With_Logo.pptx  # 10-slide PowerPoint presentation with CoE AI logo
│
├── data/
│   ├── portfolio_holdings_data.csv            # Data for all 11 stocks
│   └── sector_summary.csv                     # Sector values and percentages
│
├── screenshots/
│   ├── dashboard_screenshot.png               # Full dashboard image
│   └── coe_ai_logo.png                        # Centre of Excellence for AI logo
│
├── docs/
│   ├── Project_Report.md                      # Detailed project report
│   └── Panel_Defense_Viva_Guide.md            # Simple Q&A guide for presentations and viva
│
├── scripts/
│   ├── generate_dashboard.py                  # Python script to build the Excel workbook
│   └── generate_presentation.py               # Python script to build the PowerPoint deck
│
└── README.md                                  # Project overview and instructions
```

---

## How to Open and Use the Project

1. Download or clone this repository:
   ```bash
   git clone https://github.com/priyanshushekhar1319/Stock-Market-Portfolio-Tracker.git
   ```
2. Open `Stock_Portfolio_Master_Complete_v2.xlsx` in Microsoft Excel.
3. On the `Dashboard` sheet, click cell **`C3`**.
4. Choose any company name (e.g. `HCL TECH`, `RVNL`, `TCS`, `HDFC BANK`) to see its profit and 12-month graph update instantly.
5. Select `ALL PORTFOLIO (MASTER VIEW)` to see the entire portfolio combined.
