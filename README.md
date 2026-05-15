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

## 🧮 Analysis Engine (Formulas)

All metrics were calculated in a dedicated **Analysis sheet** using advanced Excel formulas:

### Monthly Metrics
```excel
-- Total Revenue per Month
=SUMIFS(Sales_Master[Net_Revenue_INR], Sales_Master[Month], A3)

-- Gross Margin % per Month
=SUMIFS(Sales_Master[Gross_Profit], Sales_Master[Month], A3) / 
 SUMIFS(Sales_Master[Revenue_INR], Sales_Master[Month], A3)
```

### Running Total (Cumulative Revenue)
```excel
=SUMIFS(Sales_Master[Net_Revenue_INR], Sales_Master[Month], "<="&J3)
```

### Month-over-Month Growth
```excel
=(E4-E3)/E3
```

### 3-Month Moving Average
```excel
=AVERAGE(E3:E5)  -- shifts with each month row
```

### Top 5 Products (Dynamic Ranking)
```excel
-- Get Nth highest revenue value
=LARGE($C$43:$C$72, A19)

-- Match product name to that value
=INDEX($B$43:$B$72, MATCH(C19, $C$43:$C$72, 0))
```

### Budget vs Actual Variance
```excel
-- Variance INR
=Actual_Revenue - Budgeted_Revenue

-- Variance %
=(Actual - Budget) / Budget

-- Status Flag
=IF(Variance > 0, "✅ Over", "❌ Under")
```

### Return Rate by Category
```excel
=SUMIFS(Sales_Master[Units_Returned_Clean], Sales_Master[Category], G19) /
 SUMIFS(Sales_Master[Units_Sold], Sales_Master[Category], G19)
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

**What it shows:**
- Month-by-month Net Revenue fluctuation across FY2024
- 3-Month Moving Average overlay smoothing out seasonal noise
- Clear visualization of the Diwali peak in October–November

**Key Insight:**
> Revenue dips significantly in July–August (monsoon slowdown) before recovering sharply in September and peaking in November at ₹6,98,528 — a **41% MoM jump from September to October** driven by festive season demand.

---

### 🍩 Chart 2 — Revenue by Category (Donut Chart)
**Type:** Donut Chart

**What it shows:**
- Revenue share of each of the 6 product categories
- Topwear dominates at 25% of total revenue
- Summer Collection has the smallest share at 5%

**Key Insight:**
> Topwear and Accessories together account for **36% of total revenue** despite Accessories having the highest gross margin (60.92%) — indicating an opportunity to push Accessories sales further.

---

### 🏆 Chart 3 — Top 5 Products by Revenue (Bar Chart)
**Type:** Horizontal Clustered Bar Chart

**What it shows:**
- The 5 highest revenue-generating products
- Ranked from highest to lowest
- Data labels showing exact revenue values

| Rank | Product | Revenue |
|------|---------|---------|
| 1 | Wrist Watch | ₹8,12,925 |
| 2 | Ethnic Sherwani | ₹5,66,280 |
| 3 | Formal Blazer | ₹3,78,048 |
| 4 | Quilted Jacket | ₹3,59,415 |
| 5 | Running Sneakers | ₹3,20,364 |

**Key Insight:**
> Wrist Watch alone contributes **14.98% of total revenue** — making it the single most important product in the portfolio. Its high COGS-to-price ratio also makes it one of the highest margin products.

---

### 📉 Chart 4 — Budget vs Actual Revenue by Month
**Type:** Clustered Column Chart

**What it shows:**
- Side-by-side comparison of Actual vs Budgeted revenue for each month
- Gold bars = Budget Target, Blue bars = Actual Revenue
- Clear visual of budget variance month by month

**Key Insight:**
> The business came in **under budget every single month** in FY2024 — suggesting budget targets were set too aggressively or actual demand was softer than forecast. November had the largest absolute gap (₹1,15,172 under budget) despite being the best revenue month — indicating even higher sales were expected during Diwali.

---

### 📐 Chart 5 — Cumulative Revenue Growth (Area Chart)
**Type:** Stacked Area Chart

**What it shows:**
- Running total of net revenue building across 12 months
- Exponential curve shape showing accelerating growth in Q4
- Full year total of ₹54,28,419

**Key Insight:**
> The first 6 months (Jan–Jun) contributed only **44% of annual revenue** while the last 6 months (Jul–Dec) contributed **56%** — driven heavily by festive and winter season demand, confirming strong Q4 seasonality.

---

### 🔴 Chart 6 — Return Rate by Category (Bar Chart)
**Type:** Horizontal Bar Chart

**What it shows:**
- Return rate percentage for each category
- Sorted from highest to lowest return rate
- Red color scheme indicating returns as a negative metric

| Category | Return Rate |
|----------|------------|
| Bottomwear | 6.55% |
| Topwear | 5.80% |
| Summer Collection | 5.17% |
| Accessories | 4.53% |
| Footwear | 3.96% |
| Winter Collection | 2.85% |

**Key Insight:**
> Bottomwear has the highest return rate at 6.55% — likely driven by sizing issues (the most common return reason in the dataset). Winter Collection has the lowest at 2.85% — suggesting better product-market fit or clearer sizing guidance.

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

## 🚀 How to Use This Dashboard

1. **Download** `Fashion_Apparel_Dataset.xlsx`
2. **Open** in Microsoft Excel (2016 or later recommended)
3. **Navigate** to the `DASHBOARD` tab
4. **Enable** editing if prompted
5. **Use Month slicer** on the right to filter by specific months
6. **Use Category slicer** to drill down into specific product categories
7. **Click "Clear Filter"** (X icon on slicer) to reset all filters

> ⚠️ Do not edit cells on the DASHBOARD sheet directly — all data flows from Sales_Master via Pivot Tables and formulas.

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
- ✅ **Data Visualization** — 6 chart types, dark theme dashboard design
- ✅ **Business Intelligence** — KPI tracking, slicer interactivity
- ✅ **Business Storytelling** — Translating data into actionable recommendations

---

## 👨‍💻 Author

**[Your Name]**
Data Analyst | Excel | SQL | Python | Power BI

📧 [your.email@gmail.com]
🔗 [LinkedIn Profile URL]
🐙 [GitHub Profile URL]

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

---

*This project was built as part of a data analytics portfolio to demonstrate end-to-end Excel analytics capabilities — from raw data to executive dashboard.*
