# E-commerce Sales Dataset Analysis in Excel

> End-to-end Excel project: data cleaning → star-schema modeling → descriptive statistics → PivotTables → interactive sales dashboard.

![Excel](https://img.shields.io/badge/Excel-Data_Analysis_Toolpak%20%7C%20PivotTables%20%7C%20Dashboard-217346?logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-success)
![Dataset](https://img.shields.io/badge/Records-2000_Sales_%7C_500_Customers_%7C_100_Products_%7C_20_Stores-blue)

## 📌 Overview

This project analyzes an **E-commerce Sales Dataset** entirely in **Microsoft Excel**, showcasing core data-analyst skills:

1. **Data cleaning & preparation** with Excel formulas
2. **Dimensional modeling** (Customer / Product / Store dims + Sales fact)
3. **Descriptive statistics** using the Analysis ToolPak
4. **PivotTables** for business questions
5. **Interactive Dashboard** with PivotCharts + Slicers

**Key outcome:** **$2.17M total revenue from 2,000 orders**, **$1,087 average order value**, with **Sports (25%)** and **Electronics** driving almost half of all orders. 2024 (only full year in data) drives ~50% of revenue.

## 📂 Workbook Structure

`Ecommerce_Sales_Dataset_Project.xlsx`

| Sheet | Purpose |
| :--- | :--- |
| `Customer_Dim_Raw` / `Customer_Dim_Final` | 500 customers: ID, Name, Age, Gender, City/State/Country, Loyalty_Level |
| `Product_Dim_Raw` / `Product_Dim_Final` | 100 products: ID, Name, Category, Sub_Category, Brand, Cost, Stock |
| `Store_Dim_Raw` / `Store_Dim_Final` | 20 stores: ID, Name, Region (East/North/South/West), City, Store_Type (Flagship/Outlet) |
| `Sales_Fact_Raw` / `Sales_Fact_Final` | 2,000 transactions: Sales_ID, Order_Date, Customer_ID, Product_ID, Store_ID, Quantity, Unit_Price, Discount, Payment_Type, Total_Amount |
| `Descriptive Analysis` | ToolPak stats for `Total_Amount` and `Unit_Price` |
| `Pivot Table` | 6 business pivots + written insights |
| `Dashboard` | One-page interactive Ecommerce Sales Dashboard |

> `*_Raw` = as-received data. `*_Final` = cleaned + `*_Clean` columns preserved for audit.

## 🧹 Data Cleaning (Excel Only)

All cleaning is formula-based — no data deleted, originals preserved in `*_Clean` columns.

**Customers (500 rows):**
- Names: `=PROPER(TRIM(A2))` 
  - e.g. `"  sheena bonilla  "` → `Sheena Bonilla`, `"dr. Benjamin Price"` → `Benjamin Price`
- Loyalty gaps (4 blanks): `=IF([@Loyalty_Level]="", "Platinum", [@Loyalty_Level])` → new `Loyality_Level_Clean`
  - Result: Platinum 149, Bronze 125, Gold 122, Silver 104

**Products (100 rows):**
- `Cost` (6 blanks) → `Cost_Clean = IFERROR(Cost, AVERAGE(Cost_range))` → mean imputed ~$296.83
- `Stock` (3 blanks) → `Stock_Clean = IFERROR(Stock, MEDIAN(Stock_range))`
- Names standardized with `PROPER()` — e.g. `Men's Graphic T-Shirt`

**Sales Fact (2,000 rows):**
- `Quantity` (5 blanks), `Unit_Price` (3 blanks), `Discount` (4 blanks), `Total_Amount` (4 blanks)
- Clean columns: `Quantity_Clean`, `Unit_price_Clean`, `Discount_Clean`, `Total_Amount_Clean` using `IFERROR + AVERAGE/MEDIAN`
- Recalculated check: `=Quantity_Clean * Unit_price_Clean * (1-Discount_Clean)` vs `Total_Amount_Clean`
- Date range: **Oct 2023 – Oct 2025** (2023 starts mid-Oct, 2025 ends mid-Oct; 2024 is the only full year)

Functions used: `TRIM, PROPER, SUBSTITUTE, UPPER/LOWER, IF, IFERROR, ISBLANK, AVERAGE, MEDIAN, MODE, COUNT, COUNTA, COUNTBLANK, XLOOKUP/VLOOKUP, YEAR, MONTH, TEXT`

## 📊 Descriptive Analysis

Generated via **Data → Data Analysis → Descriptive Statistics**.

**Total Sales Amount (per order):**

| Metric | Value |
| :--- | :--- |
| Mean | $1,123.99 |
| Median | $940.15 |
| Std Dev | $819.50 |
| Skewness | 0.83 (right-skewed — few large orders pull mean > median) |
| Kurtosis | -0.15 |
| Min / Max | $24.44 / $3,576.04 |
| Range | $3,551.60 |
| Sum (sample) | $560,870.85 |
| Count | 499* |

**Unit Price:**

| Metric | Value |
| :--- | :--- |
| Mean | $514.13 |
| Median | $522.75 |
| Mode | $871.34 |
| Std Dev | $282.30 |
| Min / Max | $20.09 / $999.89 |

\* Full cleaned fact = 2,000 orders, **$2,174,394.70 total, $1,087.20 AOV**. Descriptive sheet uses a 499-row sample window — use Pivot totals for KPIs.

## 📈 PivotTable Insights

### 1. Gender-wise Loyalty Count
| Gender | Bronze | Gold | Platinum | Silver | Total |
| :--- | ---: | ---: | ---: | ---: | ---: |
| Female | 60 | 68 | 83 | 51 | 262 |
| Male | 65 | 54 | 66 | 53 | 238 |
| **Total** | **125** | **122** | **149** | **104** | **500** |

> **Platinum is #1 at 149 (30%)** and leads for both genders. Females skew premium (Platinum+Gold = 58% vs 50% for males).

### 2. Sales by Payment Type
PayPal ~35%, Credit Card ~33%, COD ~32% — **essentially even (±7%)**. No single method dominates; keep all three.

### 3. Year-wise Total Sales
| Year | Revenue |
| :--- | ---: |
| 2023* | $223,813 |
| 2024 | $1,093,912 (~50%) |
| 2025* | $856,670 |
| **Total** | **$2,174,395** |

\*Partial years. 2024 is the only full calendar year — YoY not directly comparable.

### 4. Region-wise Average Order Value
North ~$1,177 | South ~$1,171 | East ~$1,096 | West ~$1,060 | **Overall ~$1,087–$1,139**
> Gap is only ~10% — order value is consistent. Revenue differences come from **order volume**, not basket size.

### 5. Category-wise Order Count
Sports **498 (25%)** > Electronics **454** > Home **410** > Clothing **334** > Beauty **304**
> **Sports + Electronics = 48% of all orders** — core focus categories.

### 6. Monthly Average Sales
AOV ranges **~$918–$1,304** around $1,087 mean. Peak: **Nov 2024 ($1,304)**. Trough: **May 2024 ($918)** — May is weak in both years. No strong trend; Oct 2023 / Oct 2025 are partial months.

## 📊 Dashboard

<img width="948" height="382" alt="Dashboard" src="https://github.com/user-attachments/assets/f20a49ce-b9e9-46e3-a485-e7759fefc8fc" />

One-page `Dashboard` sheet with:
- KPI cards: Total Revenue, Total Orders, AOV, Avg Unit Price
- Charts: Revenue by Year, Orders by Category, Revenue by Payment Type, AOV by Month, Avg Sales by Region, Loyalty by Gender
- Slicers: Year, Region, Category, Payment_Type, Gender, Loyalty_Level, Store_Type

> Select any slicer to filter all charts — e.g. click `2024` + `Sports` to see peak-category performance.

## 💡 Key Takeaways & Recommendations

1. **Protect Sports + Electronics** — half the volume; prioritize stock, promo, assortment.
2. **Push Platinum loyalty** — largest tier (30%) and female-led; expand Gold→Platinum upgrades.
3. **November playbook** — highest AOV month; replicate Nov-2024 pricing/bundles in Q4.
4. **Fix May dip** — weak both years; test spring campaign.
5. **Keep all payment options** — even split means removing one risks ~1/3 revenue.
6. **Grow West order count** — AOV is healthy everywhere; West needs traffic, not discount.

## 🛠️ Tools & Skills Demonstrated

- Excel Data Cleaning (`TRIM/PROPER/IF/IFERROR/XLOOKUP`)
- Star Schema (`Dims + Fact`, `_Raw` vs `_Final` audit trail)
- Analysis ToolPak (Mean/Median/Mode/StdDev/Skew/Kurtosis)
- PivotTables + PivotCharts + Slicers + Timelines
- Dashboard design (KPI cards, consistent formatting)
- Business storytelling (insight callouts next to each pivot)

## 🚀 How to Use

1. Clone / download this repo + open `Ecommerce_Sales_Dataset_Project.xlsx`
2. Enable **Data → Data Analysis ToolPak** to reproduce `Descriptive Analysis`
3. Use `Pivot Table` sheet to drill any question; refresh after changing `*_Final`
4. Open `Dashboard` → use Slicers to interact

## 📁 Repo Suggested Layout

```text  
/  
├── README.md  
├── Ecommerce_Sales_Dataset_Project.xlsx
├── Ecommerce_Sales_Documentation
└── Ecommerce_Sales_Screenshot 
