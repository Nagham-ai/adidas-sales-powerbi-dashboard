# Adidas US Sales Performance Dashboard

An end-to-end Power BI project that turns raw Adidas US retail sales data into a clean **Star Schema** model and a single-page interactive dashboard covering sales, profitability, products, regions, retailers and sales channels.

![Dashboard](dashboard.png)

> **Tools:** Power BI Desktop · Power Query (M) · DAX · Excel (data source)

---

## Table of Contents

1. [Business Questions](#business-questions)
2. [Dataset](#dataset)
3. [Data Model](#data-model)
4. [Data Preparation (Power Query)](#data-preparation-power-query)
5. [DAX Measures](#dax-measures)
6. [Dashboard Walkthrough](#dashboard-walkthrough)
7. [Key Insights](#key-insights)
8. [Assumptions and Limitations](#assumptions-and-limitations)
9. [Repository Structure](#repository-structure)
10. [How to Use](#how-to-use)
11. [Skills Demonstrated](#skills-demonstrated)
12. [About the Author](#about-the-author)

---

## Business Questions

The dashboard was designed to answer:

1. How are total sales, operating profit and operating margin performing, and how do they compare with the previous year?
2. Which regions, states, products and retailers drive the most revenue?
3. Where is profitability strongest and weakest?
4. Which sales channel (In-store, Outlet, Online) contributes the most?
5. Is there seasonality in monthly sales?

---

## Dataset

- **Source file:** `Adidas US Sales Datasets.xlsx` (public dataset)
- **Size:** about 9.6K sales records
- **Period:** 2020 to 2021
- **Coverage:** 6 retailers, 6 products, 5 regions, US states and cities, 3 sales methods

| Column | Description |
|---|---|
| Retailer, Retailer ID | Retail partner name and the ID provided in the source |
| Invoice Date | Date of the sale |
| Region, State, City | Location of the sale |
| Product | Product category (e.g. Men's Street Footwear, Women's Apparel) |
| Price per Unit | Selling price per unit |
| Units Sold | Number of units sold |
| Total Sales | Revenue (price x units) |
| Operating Profit | Profit from operations |
| Operating Margin | Operating profit as a share of sales |
| Sales Method | In-store, Outlet or Online |

---

## Data Model

The model follows a **Star Schema**: one fact table in the center and dimension tables around it. Every relationship is **One-to-Many** with **Single** filter direction, flowing from the dimension to the fact table.

![Data Model](data_model.png)

```
 Dim_Product ──┐
               │
 Dim_Location ─┼──►  Fact_Sales  ◄── Dim_Date
               │
 Dim_Retailer ─┘

 Key Measure  (measures only, no relationships)
```

| Table | Type | Key | Description |
|---|---|---|---|
| **Fact_Sales** | Fact | Product_ID, Location_ID, Retailer_Key, Invoice Date | One row per sales record, with the numeric columns and Sales Method |
| **Dim_Product** | Dimension | Product_ID | One row per product |
| **Dim_Location** | Dimension | Location_ID | One row per Region, State and City combination |
| **Dim_Retailer** | Dimension | Retailer_Key | One row per retailer |
| **Dim_Date** | Dimension | Date | Calendar table built with DAX (Year, Month Name, Month Number, Quarter) |
| **Key Measure** | Measures | None | Holds every DAX measure in one place |

| Relationship | Cardinality |
|---|---|
| `Dim_Product[Product_ID]` to `Fact_Sales[Product_ID]` | One to Many |
| `Dim_Location[Location_ID]` to `Fact_Sales[Location_ID]` | One to Many |
| `Dim_Retailer[Retailer_Key]` to `Fact_Sales[Retailer_Key]` | One to Many |
| `Dim_Date[Date]` to `Fact_Sales[Invoice Date]` | One to Many |

---

## Data Preparation (Power Query)

1. **Imported** the Excel file, promoted headers, removed the extra top rows and removed error rows.
2. **Set data types** for every column (date, whole number, decimal, percentage, text).
3. **Created a raw query (`Row_sales`)** with *Enable load* turned off. Every dimension and the fact table are built from it with **Reference**, so the source is cleaned once and the raw data is not loaded twice into the model.
4. **Built the dimension tables** from the raw query: *Remove Other Columns*, then *Remove Duplicates*, then *Add Index Column* (starting from 1) to create a numeric key.
5. **Built the fact table** from the raw query and used *Merge Queries* (Left Outer) to bring the numeric keys in from each dimension. Only the key column was expanded, and the descriptive text columns were removed from the fact table.
6. **Created `Dim_Date` with DAX** using `CALENDAR`, covering full calendar years, and sorted `Month Name` by `Month Number`.
7. **Validated the model:** checked that the fact table row count matches the source file, that no key column contains nulls, and that every relationship is One-to-Many.
8. **Cleaned the report view:** hid technical columns (keys, helper columns and raw numeric columns) so report authors only use measures.

### Data quality issues found and how they were handled

| Issue | What was found | Solution |
|---|---|---|
| **Retailer ID is not unique** | Several different retailers (Foot Locker, Walmart and Amazon) share the same ID in the source, so it cannot identify a retailer. Relating on it would create a Many-to-Many relationship and wrong totals. | Created a **surrogate key** (`Retailer_Key`) in `Dim_Retailer` and merged it into the fact table using the retailer name. |
| **Location table collapsed** | Removing duplicates on `Region` alone reduced the location table to 5 rows and dropped most states and cities. | Removed duplicates on **Region + State + City together**, then created `Location_ID`. |
| **Wrong format on Units Sold** | Units were displayed with a currency symbol. | Changed the format to whole number. |
| **Margin must not be averaged** | Averaging the row-level margin column gives a wrong overall margin. | Created `Operating Margin %` as total profit divided by total sales. |

---

## DAX Measures

All measures live in the `Key Measure` table.

```DAX
-- Core KPIs
Total Sales = SUM ( Fact_Sales[Total Sales] )

Total Operating Profit = SUM ( Fact_Sales[Operating Profit] )

Units Sold = SUM ( Fact_Sales[Units Sold] )

Operating Margin % = DIVIDE ( [Total Operating Profit], [Total Sales] )

Average Selling Price = DIVIDE ( [Total Sales], [Units Sold] )
```

```DAX
-- Time intelligence (year over year)
Sales LY = CALCULATE ( [Total Sales], SAMEPERIODLASTYEAR ( Dim_Date[Date] ) )

Sales YoY % = DIVIDE ( [Total Sales] - [Sales LY], [Sales LY] )

Profit LY = CALCULATE ( [Total Operating Profit], SAMEPERIODLASTYEAR ( Dim_Date[Date] ) )

Profit YoY % = DIVIDE ( [Total Operating Profit] - [Profit LY], [Profit LY] )

Margin YoY pp =
    ( [Operating Margin %] - DIVIDE ( [Profit LY], [Sales LY] ) ) * 100
```

```DAX
-- Date dimension
Dim_Date =
ADDCOLUMNS (
    CALENDAR (
        DATE ( YEAR ( MIN ( Fact_Sales[Invoice Date] ) ), 1, 1 ),
        DATE ( YEAR ( MAX ( Fact_Sales[Invoice Date] ) ), 12, 31 )
    ),
    "Year", YEAR ( [Date] ),
    "Month Number", MONTH ( [Date] ),
    "Month Name", FORMAT ( [Date], "MMMM" ),
    "Quarter", "Q" & QUARTER ( [Date] )
)
```

**Why measures instead of the ready-made columns?** A measure is calculated at report time and respects every slicer and filter, so ratios such as the operating margin stay correct at any level (region, state, retailer, month).

---

## Dashboard Walkthrough

The dashboard is a single page with dark theme and a consistent color palette.

| Area | Visual | Purpose |
|---|---|---|
| Header | Title and Slicers (Year, Region, Retailer, Product) | Filter the whole page |
| KPI cards | Total Sales, Total Operating Profit, Operating Margin %, Units Sold, Average Selling Price | Headline numbers, with year-over-year change on the first three |
| Table | Regional Performance Overview | Sales, profit and margin by region, expandable to state and city |
| Line chart | Monthly Sales Trend by Year | Seasonality and 2020 vs 2021 comparison |
| Bar chart | Top 10 States by Sales | Geographic concentration |
| Pie chart | Sales Channel Performance | Share of In-store, Outlet and Online |
| Clustered column | Sales vs Profit by Region | Compare revenue and profit side by side |
| Bar chart | Sales by Product | Product ranking |
| Bar chart | Operating Margin % by Retailer | Retailer profitability |

---

## Key Insights

- **Overall:** $900M in sales and $332M in operating profit, a **37% operating margin**, at an average selling price of **$363 per unit**.
- **West is the largest region but the least profitable.** It brings in **$270M (30% of sales)** with a **33% margin**, while the **South** earns the highest margin at **42%** on $145M. Improving West margins is the clearest profit opportunity.
- **Four states drive about a quarter of sales:** New York ($64M), California ($60M), Florida ($59M) and Texas ($46M) together make up roughly 25% of total sales.
- **Footwear is the core of the business.** The four footwear categories add up to about **two-thirds of sales**, led by **Men's Street Footwear at $209M**. Women's Apparel is the strongest apparel line at $179M.
- **In-store is the biggest channel (40%)**, ahead of Outlet (33%) and Online (28%). Online has the smallest share, which makes it worth investigating as a growth area.
- **Retailer margins are tightly grouped (35% to 41%).** Sports Direct has the highest margin at 41%, while West Gear and Walmart are lowest at 35%.
- **Seasonality:** in 2021, monthly sales dip in March and peak in July and December, and 2021 runs above 2020 across the year.

---

## Assumptions and Limitations

- The data covers **only two years (2020 to 2021)**, so year-over-year figures are meaningful for 2021 only. Selecting 2020 has no prior year to compare against.
- Operating margin is calculated on totals (profit divided by sales), not as an average of row-level margins.
- This is a public sample dataset, so findings illustrate the analysis approach and are not real business advice.

---

## Repository Structure

```
.
├── README.md
├── Adidas_US_Sales_Dashboard.pbix
├── data/
│   └── README.md            
├── images/
│   ├── dashboard.png
│   └── data_model.png
└── docs/
    └── dax_measures.md     
```

---

## How to Use

1. Download `Adidas_US_Sales_Dashboard.pbix` and open it with **Power BI Desktop**.
2. If the data source path has changed, go to **Transform data > Data source settings > Change Source** and point it to the Excel file.
3. Click **Refresh** to reload the data.

---

## Skills Demonstrated

- Data cleaning and transformation with **Power Query**
- **Dimensional modeling:** Star Schema, surrogate keys, One-to-Many relationships
- Data quality checks and validation
- **DAX:** calculations, ratios and time intelligence
- Dashboard design, KPI storytelling and visual consistency
- Turning data into business insights

---

## About the Author

Nagham Hesham Hamed: Physical Therapist transitioning into Data Analytics, working with Excel, SQL and Power BI.

- LinkedIn:[Nagham Hesham](https://www.linkedin.com/in/nagham-hesham1357/)
- GitHub: [Nagham-ai](https://github.com/Nagham-ai)

---

*This is an independent portfolio project built for learning. Adidas and its logo belong to their respective owners, and the project is not affiliated with or endorsed by Adidas.*

