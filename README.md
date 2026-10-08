# 👟 Adidas Sales Dashboard — Excel Project

An interactive Excel dashboard analyzing Adidas sales across the Middle East (Egypt, Iraq, KSA, Lebanon, Oman). It covers revenue, profit, and order performance by region, sales channel, payment method, product, and month.

## 📊 Dashboard Headline KPIs

| KPI | Value |
|---|---|
| Total Revenue | $287,378 |
| Total Profit | $85,070 |
| Total Orders | 1,200 |

## 🗺️ Project Stages

| Stage | Title | Status |
|---|---|---|
| 1 | Data Exploration in Excel | ✅ This document |
| 2 | Data Cleaning & Preparation | ⏳ Coming next |
| 3 | Data Analysis (Pivot Tables) | ⏳ Planned |
| 4 | Dashboard Design & Visualization | ⏳ Planned |
| 5 | Insights & Recommendations | ⏳ Planned |

---

# Stage 1: Data Exploration in Excel

## 🎯 Objective

Understand the structure, quality, and content of the raw Adidas sales data **before** any cleaning or analysis. The goal is to know what each column means, spot problems early, and define the business questions the dashboard will answer.

## 🧰 Tools Used

- **Microsoft Excel** (Tables, Filters, Conditional Formatting, Data Validation, Pivot Tables for quick profiling)
- Excel functions: `COUNTA`, `COUNTBLANK`, `COUNTIF`, `UNIQUE`, `MIN`, `MAX`, `AVERAGE`, `MEDIAN`, `TEXT`, `ISNUMBER`

## 📁 Dataset Overview

The dataset contains **1,200 order records**. Based on the final dashboard, the key fields are:

| Column | Type | Description | Example Values |
|---|---|---|---|
| `Order_ID` | Text / Number | Unique identifier of each order | — |
| `Date` / `Month` | Date | When the order was placed | January – December |
| `Region` | Categorical | Country where the order was made | Egypt, Iraq, KSA, Lebanon, Oman |
| `Sales_Channel` | Categorical | How the order was sold | Online, Retail, Outlet, Wholesale |
| `Payment_Method` | Categorical | How the customer paid | Apple Pay, Cash, Credit Card, Debit Card, Google Pay, NetBanking, PayPal, UPI |
| `Product_Name` | Categorical | Adidas product sold | Predator Freak, Ultraboost Light, Primegreen Jacket, Copa Sense, Ultraboost 22, Stan Smith, Superstar, Gym Bag, etc. |
| `Quantity` | Numeric | Units per order | — |
| `Revenue` | Currency (USD) | Total sales value | — |
| `Profit` | Currency (USD) | Revenue minus cost | — |

> 📝 **Note:** Adjust column names and types in this table to match your exact worksheet headers.

## 🔎 Exploration Steps

### 1. Load and Inspect the Data
- Open the raw file and review each sheet.
- Convert the range to an **Excel Table** (`Ctrl + T`) so formulas and pivots update automatically.
- Freeze the header row (`View → Freeze Panes`) and autofit columns.

### 2. Check Size and Structure
- Count rows and columns:
  ```excel
  =ROWS(Table1[Order_ID])
  =COLUMNS(Table1[#Headers])
  ```
- Confirm there are **1,200 records**.

### 3. Verify Data Types
- Make sure dates are real dates, and Revenue, Profit, and Quantity are numbers (not text).
- Test with:
  ```excel
  =COUNTIF(Table1[Revenue], "<>") - SUMPRODUCT(--ISNUMBER(Table1[Revenue]))
  ```
  (A result above 0 means some values are stored as text.)

### 4. Look for Missing Values
```excel
=COUNTBLANK(Table1[Region])
=COUNTBLANK(Table1[Revenue])
```
Highlight blanks with **Conditional Formatting → Highlight Cells Rules → Blanks**.

### 5. Detect Duplicates
- Use `Data → Remove Duplicates` preview, or Conditional Formatting → *Duplicate Values* on `Order_ID`.
- Record how many duplicates exist (do not delete yet; that happens in Stage 2).

### 6. Profile Categorical Columns
List the unique values of each category and check for spelling or case inconsistencies:
```excel
=UNIQUE(Table1[Region])
=UNIQUE(Table1[Payment_Method])
=UNIQUE(Table1[Sales_Channel])
=UNIQUE(Table1[Product_Name])
```
Check that:
- 5 regions are present (Egypt, Iraq, KSA, Lebanon, Oman)
- 4 sales channels (Online, Retail, Outlet, Wholesale)
- 8 payment methods
- No variants like `"egypt "`, `"EGYPT"`, `"Gooogle Pay"`

### 7. Profile Numeric Columns
| Statistic | Formula |
|---|---|
| Minimum | `=MIN(Table1[Revenue])` |
| Maximum | `=MAX(Table1[Revenue])` |
| Average | `=AVERAGE(Table1[Revenue])` |
| Median | `=MEDIAN(Table1[Revenue])` |
| Std. Deviation | `=STDEV.S(Table1[Revenue])` |

Repeat for `Profit` and `Quantity`. Flag **negative values**, **zeros**, and **outliers** (values far beyond the average) for review.

### 8. Logical Consistency Checks
- `Profit` should not exceed `Revenue`.
- `Quantity` should be a positive whole number.
- Dates should fall within a single year range.
- Add a helper column:
  ```excel
  =IF([@Profit]>[@Revenue],"Check","OK")
  ```

### 9. Create Helper Metrics (for exploration only)
```excel
Profit Margin % = [@Profit] / [@Revenue]
Month Name      = TEXT([@Date], "mmmm")
Revenue / Order = [@Revenue] / [@Quantity]
```

### 10. Quick Exploratory Pivot Tables
Build rough pivots to get a feel for the data:
- Revenue and Profit by **Region**
- Revenue by **Sales Channel**
- Revenue by **Payment Method**
- Top products by Revenue and Profit
- Revenue and Profit by **Month**

## ❓ Business Questions Defined

These questions drive the rest of the project and the dashboard visuals:

1. Which region generates the **highest and lowest profit**?
2. Which region generates the **highest and lowest revenue**?
3. Which **sales channel** performs best, and which performs worst?
4. Which **payment method** leads, and which ranks lowest?
5. What are the **top 5 best-selling** products by revenue?
6. What are the **top 5 most profitable** products?
7. How do **revenue and profit change month by month**?

## 👀 Early Observations (from the exploration)

| Area | Observation |
|---|---|
| Regions | Egypt appears strongest in both revenue and profit; Lebanon the weakest |
| Channels | Online leads in sales; Wholesale is the lowest |
| Payments | PayPal leads; Google Pay is the lowest |
| Products | Predator Freak and Ultraboost Light rank at the top for both revenue and profit |
| Trend | Monthly revenue shows a general downward trend through the year; profit fluctuates more |

> These are preliminary and are validated in Stage 3.

## ⚠️ Data Quality Log

Fill this in as you explore:

| # | Issue Found | Column | Count | Planned Action (Stage 2) |
|---|---|---|---|---|
| 1 | Missing values | — | — | — |
| 2 | Duplicate orders | `Order_ID` | — | — |
| 3 | Inconsistent text | — | — | — |
| 4 | Wrong data types | — | — | — |
| 5 | Outliers / negatives | — | — | — |

## ✅ Stage 1 Deliverables

- [x] Raw data loaded and converted to an Excel Table
- [x] Column dictionary documented
- [x] Missing values, duplicates, and data types checked
- [x] Categorical and numeric columns profiled
- [x] Exploratory pivot tables created
- [x] Business questions defined
- [x] Data quality issues logged for Stage 2

## ➡️ Next Stage

**Stage 2: Data Cleaning & Preparation** — fix the issues logged above (remove duplicates, standardize text, correct data types, handle missing values) and prepare the final dataset for analysis.
