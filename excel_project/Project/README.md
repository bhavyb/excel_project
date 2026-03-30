# 📊 Bhavy Excel Dashboard — Superstore Sales 2023

An interactive Excel workbook demonstrating end-to-end data analytics skills — from raw data import and cleaning through pivot analysis, advanced formulas, data visualization, and a polished KPI dashboard.

**File:** `BhavyExcelDashboard.xlsx`  
**Dataset:** Superstore Sales 2023 (Kaggle / UCI ML Repository)  
**Total Records:** 200 orders | **Total Sheets:** 7

---

## 📁 Workbook Structure

```
BhavyExcelDashboard.xlsx
├── 1_Dataset              → 200-row raw Superstore Sales dataset
├── 2_DataCleaning         → 10-step cleaning log + summary statistics
├── 3_PivotTables          → Sales aggregations by Category, Region, Segment
├── 4_AdvancedFormulas     → VLOOKUP, INDEX-MATCH, IF, COUNTIF/SUMIF, TEXT
├── 5_DataVisualization    → Bar chart, Pie chart, Conditional Formatting
└── 6_Dashboard            → Interactive KPI dashboard with charts & tables
```

---

## 📊 Key Metrics (from Dashboard)

| KPI | Value |
|-----|-------|
| 💰 Total Sales | $301,247.38 |
| 📈 Total Profit | $45,095.84 |
| 📦 Total Orders | 200 |
| 🏷️ Avg Discount | 9.65% |
| 💹 Profit Margin | 14.97% |

---

## 🗂️ Sheet-by-Sheet Breakdown

### 1️⃣ `1_Dataset` — Raw Data 

**Skills:** Data Import, Initial Cleaning

- **200 rows** of Superstore Sales data for year 2023
- **13 columns:** Order ID, Order Date, Ship Date, Customer Name, Segment, Region, Category, Sub-Category, Product Name, Sales, Quantity, Discount, Profit
- Auto-filter enabled on all columns
- Alternating row shading for readability
- Profit values colour-coded: 🔴 red (loss) / 🟢 green (profit)

**Sample Data:**

| Order ID | Customer | Category | Sales | Profit |
|----------|----------|----------|-------|--------|
| ORD-0001 | Leo Patel | Furniture | $749.78 | -$6.75 |
| ORD-0003 | Eva Green | Technology | $2,281.25 | $271.57 |
| ORD-0008 | Frank Lee | Office Supplies | $2,060.15 | $574.21 |

---

### 2️⃣ `2_DataCleaning` — Data Preparation
**Skills:** Remove duplicates, Handle missing values, Format data types

**10 Cleaning Steps Completed:**

| Step | Action | Status |
|------|--------|--------|
| 1 | Check for Duplicates (Order ID) | ✔ Done |
| 2 | Handle Missing Values (all columns) | ✔ Done |
| 3 | Format Order Date → YYYY-MM-DD | ✔ Done |
| 4 | Format Ship Date → YYYY-MM-DD | ✔ Done |
| 5 | Validate Discount Range [0, 0.3] | ✔ Done |
| 6 | Flag Negative Profits | ✔ Done |
| 7 | Standardize Text Case (PROPER) | ✔ Done |
| 8 | Remove Leading/Trailing Spaces (TRIM) | ✔ Done |
| 9 | Validate Quantity ≥ 1 | ✔ Done |
| 10 | Verify Numeric Data Types (Sales, Profit) | ✔ Done |

**Dataset Summary Statistics:**

| Metric | Value |
|--------|-------|
| Total Records | 200 |
| Total Sales | $301,247.38 |
| Total Profit | $45,095.84 |
| Average Discount | 9.65% |
| Unique Customers | 12 |
| Profitable Orders | 174 |
| Loss-making Orders | 26 |

---

### 3️⃣ `3_PivotTables` — Analysis
**Skills:** Data Analysis, Summarization

**Sales & Profit by Category:**

| Category | Total Sales | Total Profit | Avg Discount | Orders |
|----------|------------|--------------|--------------|--------|
| Furniture | $98,809.24 | $13,726.07 | 9.84% | 64 |
| Office Supplies | $95,608.21 | $14,144.27 | 10.00% | 65 |
| Technology | $106,829.93 | $17,225.50 | 9.15% | 71 |
| **Grand Total** | **$301,247.38** | **$45,095.84** | **9.65%** | **200** |

**Sales & Profit by Region:**

| Region | Total Sales | Total Profit | Orders |
|--------|------------|--------------|--------|
| Central | $90,039.33 | $14,210.86 | 56 |
| East | $79,782.68 | $11,968.70 | 49 |
| South | $67,162.95 | $9,256.37 | 51 |
| West | $64,262.42 | $9,659.91 | 44 |

**Sales & Profit by Customer Segment:**

| Segment | Total Sales | Total Profit |
|---------|------------|--------------|
| Consumer | $95,569.27 | $13,594.12 |
| Corporate | $103,344.90 | $17,549.60 |
| Home Office | $102,333.21 | $13,952.12 |

---

### 4️⃣ `4_AdvancedFormulas` — Formula Showcase

**Skills:** VLOOKUP, INDEX-MATCH, IF conditions, Nested functions

**VLOOKUP — Order Lookup (ORD-0042):**
- Customer Name → Grace Kim
- Segment → Home Office
- Sales → $1,230.10

**INDEX-MATCH — Customer Lookup (Alice Johnson):**
- First Order ID → ORD-0004
- Category → Technology
- Max Sales → $2,652.36

**Nested IF — Profit Rating Logic:**
```excel
=IF(Profit > 500, "High", IF(Profit > 0, "Medium", "Loss"))
=IF(Sales > 1000, "Top Seller", IF(Sales > 500, "Good", "Below Avg"))
```

**COUNTIF / SUMIF — Aggregation by Category:**

| Category | Count | Total Sales | Total Profit | Avg Sales |
|----------|-------|------------|--------------|-----------|
| Furniture | 64 | $98,809.24 | $13,726.07 | $1,543.89 |
| Office Supplies | 65 | $95,608.21 | $14,144.27 | $1,470.90 |
| Technology | 71 | $106,829.93 | $17,225.50 | $1,504.65 |

**TEXT Functions — Date Formatting:**

| Sample Date | Formatted | Month | Year |
|-------------|-----------|-------|------|
| 2023-03-15 | March 15, 2023 | March | 2023 |
| 2023-07-04 | July 04, 2023 | July | 2023 |
| 2023-11-28 | November 28, 2023 | November | 2023 |

---

### 5️⃣ `5_DataVisualization` — Charts & Formatting
**Skills:** Charts, Conditional Formatting

- 📊 **Clustered Bar Chart** — Sales & Profit by Category
- 🥧 **Pie Chart** — Sales Distribution by Region
- 🎨 **Data Bars** — Applied on Sales column (blue gradient)
- 🌈 **3-Color Scale** — Applied on Profit column (Red → Yellow → Green)

---

### 6️⃣ `6_Dashboard` — Interactive Dashboard

**Skills:** Slicers, Dynamic ranges, Dashboard design

**5 KPI Cards:**

| 💰 Total Sales | 📈 Total Profit | 📦 Total Orders | 🏷️ Avg Discount | 💹 Profit Margin |
|---|---|---|---|---|
| $301,247.38 | $45,095.84 | 200 | 9.65% | 14.97% |

**Top 5 Sub-Categories by Sales:**

| Sub-Category | Sales | Profit | Orders |
|---|---|---|---|
| Tables | $27,674.33 | $4,098.11 | 15 |
| Phones | $26,191.29 | $5,404.12 | 20 |
| Chairs | $25,465.15 | $2,758.96 | 17 |
| Accessories | $25,011.35 | $4,617.30 | 14 |
| Binders | $21,552.23 | $3,000.88 | 17 |
---
