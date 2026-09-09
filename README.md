# Data Analytics Projects: Client Transactions & FMCG Excel Analysis

A collection of practical Excel data analytics projects showcasing data cleaning, transformation, financial calculations, advanced lookup functions, and exploratory data analysis.

---

## 📋 Overview

This repository contains two hands-on data analytics projects built in Excel:
1. **Client Transactions Analysis & Data Cleaning**: Cleaning, standardizing, and restructuring raw client transaction data for financial reporting.
2. **FMCG Sales Analysis & Lookup Practice**: Analyzing a large-scale Fast-Moving Consumer Goods (FMCG) dataset (2022–2024) and practicing production-level lookup techniques (`VLOOKUP` and `INDEX-MATCH`).

---

## 📁 Project 1: Client Transactions Analysis & Data Cleaning

### 1. Dataset Summary
* **Source**: Raw client transaction records (`Client_Transactions.xlsx`).
* **Timeframe**: May 30, 2023 – June 2, 2023.
* **Size**: 31 raw records reduced to 28 clean, unique transactions.
* **Key Fields**: Date, Client Name, Contact, Department, Payment Method, Revenue, Profit, Profit Margin.

### 2. What I Did & Steps Taken
- **AutoFit Formatting**: Auto-adjusted row heights and column widths (`Alt + H + O + I` / `Alt + H + O + A`) to clean up cramped layouts.
- **Text Standardizing**:
  - Removed stock tickers in parentheses `(*)` from client names using **Find & Replace**.
  - Standardized company names to lowercase using `=LOWER()`.
  - Cleaned contact names to proper casing and stripped extra spaces using `=TRIM(PROPER())`.
- **Delimited Data Splitting**: Split combined department-location strings (e.g., `Cloud Tech_Texas`) into separate `Department` and `Region` columns using **Text to Columns** (`_` delimiter).
- **Deduplication**: Removed duplicate records using Excel's **Remove Duplicates** tool.
- **Missing Value & Error Handling**:
  - Filled blank fields with `NA` flags for clearer auditability.
  - Calculated `Profit Margin` safely using `=IFERROR(Profit/Revenue, "NA")` to prevent `#DIV/0!` errors on missing revenue values.
- **Visual Presentation**: Disabled worksheet gridlines for a clean presentation dashboard layout.

---

## 📁 Project 2: FMCG Sales Analysis & Advanced Lookup Practice

### 1. Dataset Summary
* **Source**: Master FMCG transactional dataset (`FMCG_2022_ 2024.xlsx`).
* **Timeframe**: January 21, 2022 – December 31, 2024.
* **Size**: 190,757 rows covering 3,799,824 total units sold and **$19.95M+** in revenue.
* **Dimensions**: 30 SKUs, 14 Brands, 5 Product Categories (Yogurt, Milk, ReadyMeal, SnackBar, Juice), 3 Sales Channels, and 3 Regions.

### 2. What I Did & Practice Focus
Used this dataset to practice analyzing sales distributions while mastering real-world Excel lookup techniques across dedicated practice sheets.

#### **Data Analysis Highlights**
- Aggregated sales revenue and volume across product categories, sales channels (Retail, Discount, E-commerce), and regions (Central, North, South).
- Highlighted performance trends — **Yogurt** emerged as the top-performing category ($8.22M revenue, 1.56M units sold).

#### **Lookup Practice (`VLOOKUP` Sheet)**
- **Basic Lookups**: Mapped SKUs to Brand names using exact matching.
- **Handling Messy Inputs**: Wrapped lookups in `=TRIM()` to fix leading/trailing spaces in source strings.
- **Multi-Criteria Lookups**: Created composite helper keys (`SKU & "_" & Region`) to look up values across multiple conditions.
- **Text Formatting In-Formula**: Combined `=LEFT()` and `=RIGHT()` to repair broken SKU formats dynamically.
- **Dynamic Headers**: Nested `=MATCH()` and `=COLUMNS()` to make column indexes responsive to layout changes.
- **Price Tiering**: Built numeric price buckets (*Value*, *Mid*, *Premium*) using approximate matching (`TRUE/1`).
- **Error Handling**: Replaced `#N/A` errors gracefully using `=IFERROR()`, `=IFNA()`, and `=COUNTIF()` presence checks.

#### **Advanced Lookup Practice (`INDEX-MATCH` Sheet)**
- **Left Lookups**: Retrieved data columns located to the left of lookup keys where `VLOOKUP` fails.
- **Two-Way Lookups**: Combined row and column `MATCH` functions for dynamic grid data retrieval.
- **Multi-Condition Array Matching**: Applied boolean logic `MATCH(1, (Range1=Cond1)*(Range2=Cond2), 0)` to fetch single records under multiple rules.
- **Dynamic Peak Analysis**: Nested `=MAX()` within `=MATCH()` to dynamically locate top-selling SKUs (e.g., identifying SKU `SN-019`).

---

## 🛠️ Tools & Skills Used
* **Software**: Microsoft Excel
* **Functions & Formulas**: `VLOOKUP`, `INDEX`, `MATCH`, `IFERROR`, `IFNA`, `COUNTIF`, `LOWER`, `TRIM`, `PROPER`, `MAX`
* **Features**: Text to Columns, Remove Duplicates, Find & Replace, Go To Special (Blanks), AutoFit Layouts
* **Analytics**: Data Cleaning, Data Standardization, Multi-criteria Analysis, Financial Metric Calculation
