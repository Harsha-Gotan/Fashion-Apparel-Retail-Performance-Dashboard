# 👗 Fashion & Apparel Retail Performance Dashboard

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Power Query](https://img.shields.io/badge/ETL-Power%20Query-F5A623?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=for-the-badge)
![Domain](https://img.shields.io/badge/Domain-Retail%20%26%20Finance-2E86AB?style=for-the-badge)

---

## 📌 Project Overview

A comprehensive **end-to-end Excel analytics project** simulating a real-world Fashion & Apparel retail business operating in the **India market (INR ₹)**. This project covers the full data analyst workflow — from raw data ingestion and ETL, to advanced formula-based analysis, to an executive-level interactive dashboard.

The dataset spans **12 months of FY2024** with **1,500 sales transactions** across **30 products** and **6 categories**, designed to reflect realistic Indian retail seasonality patterns including Diwali, Navratri, and summer peaks.

---

## 🧩 Business Problem

A mid-sized Fashion & Apparel retail brand operating across multiple product categories in India needed a **centralized performance tracking system** to answer the following business questions:

- Which months and categories are driving the most revenue?
- How are we performing against our monthly budget targets?
- Which products are our top performers and which are underperforming?
- What is our gross margin across categories — where are we most profitable?
- Which categories have the highest return rates and why?
- How is our cumulative revenue tracking across the financial year?
- How do sales trends shift across weeks and months?

Without a structured dashboard, the business was relying on raw spreadsheet data with no visibility into trends, profitability, or budget variance — making it impossible to make data-driven decisions quickly.

---

## 🛠️ Tech Stack

| Tool | Usage |
|------|-------|
| **Microsoft Excel** | Primary analytics and dashboard tool |
| **Power Query (M Language)** | ETL — data cleaning, transformation, merging |
| **Pivot Tables** | Aggregation and dynamic data summarization |
| **Pivot Charts** | Interactive visual reporting |
| **Advanced Excel Formulas** | SUMIFS, INDEX-MATCH, LARGE, OFFSET, IFERROR |
| **Conditional Formatting** | Heatmaps, color scales, variance flagging |
| **Excel Slicers** | Interactive dashboard filtering |

---

## 📁 Data Architecture

The project is built on **4 raw data tables:**

```
Fashion_Apparel_Dataset.xlsx
│
├── Sales_Raw          → 1,500 daily sales transactions (Jan–Dec 2024)
├── Product_Master     → 30 products across 6 categories with COGS & price
├── Budget_Plan        → Monthly budgeted revenue & units per category
└── Returns_Raw        → 104 return records (~7% return rate)
```

### Data Dictionary

**Sales_Raw**
| Column | Description |
|--------|-------------|
| Sale_ID | Unique transaction identifier |
| Date | Transaction date |
| Week_Number | ISO week number (1–52) |
| Month | Month number (1–12) |
| Month_Name | Month name (January–December) |
| Product_ID | Product identifier (P001–P030) |
| Category | Product category |
| Units_Sold | Number of units sold |
| Sale_Price | Actual selling price (post discount) |
| Discount_% | Discount applied (0%, 5%, 10%, 15%) |

**Product_Master**
| Column | Description |
|--------|-------------|
| Product_ID | Unique product identifier |
| Product_Name | Product name |
| Category | One of 6 categories |
| COGS | Cost of Goods Sold per unit |
| Launch_Date | Product launch date |
| Sale_Price | Standard selling price |

---

## ⚙️ ETL Process (Power Query)

All raw data was ingested and transformed using **Power Query** before analysis:

```
Raw Data Sources (4 tables)
        ↓
Power Query Editor
        ↓
├── Fix data types (Date, Whole Number, Decimal)
├── Remove blank rows
├── Merge Sales_Raw + Product_Master on Product_ID
│       └── Brings in: Product_Name, COGS
├── Add calculated columns:
│       ├── Revenue_INR = Units_Sold × Sale_Price
│       ├── Total_COGS = Units_Sold × COGS
│       ├── Gross_Profit = Revenue_INR - Total_COGS
│       └── Gross_Margin_% = Gross_Profit / Revenue_INR
├── Merge Returns_Raw on Sale_ID (Left Outer Join)
│       └── Brings in: Units_Returned
├── Handle nulls: Units_Returned → 0 where null
├── Add Net calculations:
│       ├── Net_Units_Sold = Units_Sold - Units_Returned
│       └── Net_Revenue_INR = Net_Units_Sold × Sale_Price
└── Load clean master table → Sales_Master (1,500 rows, 19 columns)
```

---

## 📊 Dashboard Walkthrough

The executive dashboard is built on a **dark navy theme (#1F2D40)** with gold (#F5A623) and blue (#2E86AB) accents, designed for clarity and visual impact.

### 🔢 KPI Cards (Top Row)
Six headline metrics giving an instant business snapshot:

| KPI | Value | Insight |
|-----|-------|---------|
| Total Net Revenue | ₹54,28,419 | Full year net revenue after returns |
| Units Sold | 4,491 | Total units across all categories |
| Avg Gross Margin | 50.16% | Healthy margin across portfolio |
| Return Rate | 4.85% | Low return rate — good product quality |
| Best Month | November | Diwali season peak performance |
| Top Category | Topwear | Highest revenue generating category |

---

### 📈 Chart 1 — Monthly Net Revenue Trend + 3M Moving Average
**Type:** Line Chart with Moving Average Trendline
- Tracks month-by-month revenue performance across FY2024 with a smoothed moving average overlay to highlight the underlying trend beneath seasonal noise.

---

### 🍩 Chart 2 — Revenue by Category (Donut Chart)
**Type:** Donut Chart
- Shows each category's share of total revenue, making it instantly clear which categories dominate the portfolio and which are underrepresented.

---

### 🏆 Chart 3 — Top 5 Products by Revenue (Bar Chart)
**Type:** Horizontal Clustered Bar Chart
- Ranks the five highest-earning products by net revenue, helping identify which SKUs are carrying the business.
---

### 📉 Chart 4 — Budget vs Actual Revenue by Month
**Type:** Clustered Column Chart
- Compares monthly actual revenue against budgeted targets side by side, highlighting how far off the business was from its plan each month.
---

### 📐 Chart 5 — Cumulative Revenue Growth (Area Chart)
**Type:** Stacked Area Chart
- Visualizes how total revenue compounds across the year, making Q4 festive season acceleration immediately visible.

---

### 🔴 Chart 6 — Return Rate by Category (Bar Chart)
**Type:** Horizontal Bar Chart
- Ranks categories by return rate percentage, flagging where product quality or sizing issues may be hurting customer satisfaction.
  
---

### 🎛️ Interactive Slicers
Two slicers control all interactive charts simultaneously:

- **Month Slicer** — filter by any month (January–December)
- **Category Slicer** — filter by any product category

Clicking any combination instantly updates the Trend chart, Top 5 Products, Donut chart, and Returns chart in real time.

---

## 💡 Business Insights & Recommendations

### Insight 1 — Festive Season is the Business Lifeline
**Finding:** October and November combined account for **24% of annual revenue**

**Recommendation:**
- Increase inventory procurement by 30–40% starting September
- Launch targeted marketing campaigns 3 weeks before Diwali
- Introduce festive-exclusive product bundles (Ethnic Sherwani + Jutties combo)
- Negotiate early supplier contracts to avoid festive season stock-outs

---

### Insight 2 — Accessories is the Hidden Profit Driver
**Finding:** Accessories has the highest gross margin at **60.92%** but only 11% revenue share

**Recommendation:**
- Cross-sell Accessories at checkout with every Topwear purchase
- Bundle Wrist Watch + Leather Belt + Wallet as a "Complete Look" package
- Increase Accessories inventory and marketing spend
- A 5% increase in Accessories revenue share could improve overall margin by ~2%

---

### Insight 3 — Bottomwear Returns Need Investigation
**Finding:** Bottomwear return rate of **6.55%** is highest across all categories

**Recommendation:**
- Add detailed size guide charts to all Bottomwear product pages
- Introduce "Try Before You Buy" option for Jeans and Trousers
- Analyze return reason data — "Size Issue" is likely the top reason
- Consider adding half-sizes or stretch-fit options

---

### Insight 4 — Budget Planning Needs Recalibration
**Finding:** Business came in **under budget every month** in FY2024

**Recommendation:**
- Reduce FY2025 budget targets by 10–15% to set realistic benchmarks
- Build separate budget models for festive vs non-festive months
- Introduce rolling forecasts updated quarterly instead of annual fixed budgets
- Use MoM growth rates from this analysis to build data-driven FY2025 targets

---

### Insight 5 — Monsoon Season (Jul–Aug) Needs a Strategy
**Finding:** July and August show the **steepest revenue decline** (–15% MoM in July)

**Recommendation:**
- Launch monsoon-specific product lines (rainwear, waterproof accessories)
- Run clearance sales in July to move slow-moving summer inventory
- Shift marketing spend from July–August to September pre-festive buildup
- Introduce loyalty rewards redeemable during monsoon months to maintain engagement

---

### Insight 6 — Winter Collection is Underperforming
**Finding:** Winter Collection has only **17% revenue share** despite premium pricing

**Recommendation:**
- Launch Winter Collection earlier (October instead of November)
- Target northern India markets more aggressively where winter demand is higher
- Bundle Winter Collection items with festive wear for October Diwali shoppers

---

## 📂 File Structure

```
Fashion_Apparel_Dashboard/
│
├── Fashion_Apparel_Dataset.xlsx     ← Main Excel file
│   ├── DASHBOARD                   ← Executive dashboard (start here)
│   ├── Sales_Master                ← Clean merged data (Power Query output)
│   ├── Analysis                    ← All formula calculations
│   ├── Pivot_Data                  ← Pivot tables feeding dashboard
│   ├── Sales_Raw                   ← Raw transaction data
│   ├── Product_Master              ← Product reference table
│   ├── Budget_Plan                 ← Monthly budget targets
│   └── Returns_Raw                 ← Returns transaction data
│
└── README.md                       ← This file
```

---

## 📈 Key Metrics Summary

| Metric | Value |
|--------|-------|
| Total Net Revenue | ₹54,28,419 |
| Total Units Sold | 4,491 |
| Average Gross Margin | 50.16% |
| Overall Return Rate | 4.85% |
| Best Month | November (₹6,98,528) |
| Worst Month | August (₹3,21,812) |
| Top Category by Revenue | Topwear (25%) |
| Top Category by Margin | Accessories (60.92%) |
| Top Product | Wrist Watch (₹8,12,925) |
| Months Under Budget | 12/12 (100%) |

---

## 🎓 Skills Demonstrated

- ✅ **ETL & Data Cleaning** — Power Query M language, multi-table merging
- ✅ **Data Modeling** — Relational table design, calculated columns
- ✅ **Advanced Formulas** — SUMIFS, INDEX-MATCH, LARGE, OFFSET, IFERROR
- ✅ **Statistical Analysis** — Moving averages, running totals, growth rates
- ✅ **Financial Analysis** — Gross margin, budget variance, profitability
- ✅ **Data Visualization** — 6 chart types
- ✅ **Business Intelligence** — KPI tracking, slicer interactivity
- ✅ **Business Storytelling** — Translating data into actionable recommendations

---

## Screenshot

<img width="1442" height="750" alt="Screenshot 2026-05-15 114246" src="https://github.com/user-attachments/assets/575cb330-3e48-4e9b-bbba-070581a88b64" />


