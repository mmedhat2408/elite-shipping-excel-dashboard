# Elite Shipping Dashboard | Excel Sales Analytics

An end-to-end Excel analytics project: raw CSV files are cleaned in **Power Query**, modeled as a **star schema in Power Pivot**, summarized with **Pivot Tables**, and presented in an interactive **dashboard** with slicers and a timeline.

---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Tools & Skills](#2-tools--skills)
3. [Data Sources](#3-data-sources)
4. [Stage 1: Data Cleaning (Power Query)](#4-stage-1-data-cleaning-power-query)
5. [Stage 2: Data Modeling (Power Pivot)](#5-stage-2-data-modeling-power-pivot)
6. [Stage 3: Pivot Tables](#6-stage-3-pivot-tables)
7. [Stage 4: Dashboard](#7-stage-4-dashboard)
8. [Key Insights](#8-key-insights)
9. [Repository Structure](#9-repository-structure)
10. [How to Use](#10-how-to-use)

---

## 1. Project Overview

The goal of this project is to turn messy, multi-file sales data into a clean analytical model and an interactive dashboard that answers questions such as:

- How much did we sell, and how many orders, units and customers were involved?
- Which products and categories drive revenue, and which underperform?
- How do sales and orders change by month and by day of the week?
- How are sales distributed across regions?

The workflow follows four stages: **Clean → Model → Analyze → Visualize**, all inside Excel.

## 2. Tools & Skills

| Area | Details |
|---|---|
| Tool | Microsoft Excel |
| ETL | Power Query (M language) |
| Modeling | Power Pivot, star schema, relationships, Calendar table, DAX measures |
| Analysis | Pivot Tables, Top/Bottom N filters, % of grand total |
| Visualization | Pivot Charts, Slicers, Timeline, custom dark-theme dashboard |

## 3. Data Sources

| File | Columns | Notes |
|---|---|---|
| `Customers.csv` | CustomerID, CustomerName, Region | Customer master data |
| `Products.csv` | ProductID, ProductName, Category, Price | Product master data |
| `Orders/` (folder) | OrderID, CustomerID, ProductID, Quantity, OrderDate | One CSV file per month; the month number is in the file name (e.g. `Orders_1.csv`) |

---

## 4. Stage 1: Data Cleaning (Power Query)

All transformations are recorded as Applied Steps (full M code is in `Power Query/Applied Steps For Power Query.txt`).

### 4.1 Customers table
| Step | Purpose |
|---|---|
| Promote first row to headers | Use the CSV header row as column names |
| Change data types | `CustomerID` → Whole Number; `CustomerName`, `Region` → Text |
| Capitalize Each Word + Trim (CustomerName) | Standardize name formatting and remove extra spaces |
| Replace Value (Region): space → empty | Remove stray spaces inside region values |
| Trim + Capitalize Each Word (Region) | Standardize region labels (Cairo, Giza, Aswan, Alexandria, Mansoura) |

### 4.2 Products table
| Step | Purpose |
|---|---|
| Promote headers and set types | `ProductID` → Whole Number; other columns → Text |
| Trim + Capitalize Each Word (ProductName, Category) | Consistent naming and categories |
| Replace `$` with empty in Price | Remove currency symbols so the value can be converted |
| Replace `Two Hundred` with `200` | Fix a price typed as text |
| Change Price type to Currency | Make Price usable in calculations |

### 4.3 Orders table (the most complex query)
**Combine files**
1. Connect with the **Folder** connector and filter out hidden files.
2. Invoke the custom function (**Transform File**) on each file's content, keep `Source.Name` and expand the combined table.

**Clean the data**
3. Set `OrderID`, `CustomerID`, `ProductID` and `Quantity` to Whole Number.
4. Filter out rows where `ProductID` or `Quantity` is null or empty.

**Fix inconsistent dates (OrderDate)**
5. Replace `/` with `-` so every date uses the same delimiter.
6. Split `OrderDate` by `-` into three parts (`OrderDate.1`, `OrderDate.2`, `OrderDate.3`) and set them to Whole Number.
7. **Year**: a conditional column picks the part greater than 31, since only the year can exceed 31.
8. **Month**: duplicate `Source.Name`, extract the text between `_` and `.`, and convert it to a number (the month comes from the file name).
9. **Day**: a custom column picks the part that is ≤ 31 and different from the Month.
10. Merge Day-Month-Year with the `en-GB` culture and convert to a proper **Date** type as `Order Date`.
11. Remove the helper columns (`OrderDate.1/2/3`, `Year`, `Month`, `Day`).

**Rebuild a unique key**
12. Remove the original `OrderID` (duplicates across files) and add an **Index column** as the new `Order ID`.

**Final Orders columns:** `Source.Name`, `Order ID`, `CustomerID`, `ProductID`, `Quantity`, `Order Date`.

---

## 5. Stage 2: Data Modeling (Power Pivot)

The cleaned tables are loaded into the Data Model and arranged as a **star schema**.

| Table | Type | Key columns |
|---|---|---|
| **Orders** | Fact | Order ID, CustomerID, ProductID, Quantity, Order Date |
| **Customers** | Dimension | CustomerID, CustomerName, Region |
| **Products** | Dimension | ProductID, ProductName, Category, Price |
| **Calendar** | Dimension (date table) | Date, Year, Month Number, Month, MMM-YYYY, Day Of Week Number, Day Of Week, Date Hierarchy (Year → Month → Date) |

**Relationships (one-to-many, single direction)**
- `Customers[CustomerID]` 1 → \* `Orders[CustomerID]`
- `Products[ProductID]` 1 → \* `Orders[ProductID]`
- `Calendar[Date]` 1 → \* `Orders[Order Date]`

**Measures created in the Orders table**
- **Total Sales**: sum of Quantity × Price
- **Total Quantities**: sum of Quantity
- **Total Orders**: count of orders
- **Total Customers**: distinct count of customers

Additional calculations such as *Average Sales per Order*, *Average Sales by Category/Region* and *% of Grand Total* are built on top of these measures.

> Note: the exact DAX formulas are stored in the Power Pivot model inside the workbook.

---

## 6. Stage 3: Pivot Tables

Pivot Tables connected to the Data Model feed every chart on the dashboard:

| Pivot Table | Purpose |
|---|---|
| Top 5 vs Bottom 5 Selling Products | Top/Bottom N filter on Total Sales by Product |
| Monthly Total Sales | Sales trend across 12 months |
| Total Orders vs Average Sales per Order (by Month) | Order volume vs. order value |
| Total Sales by Region (% of total) | Regional contribution |
| Total Sales by Category (% of total) | Category contribution |
| Category Total Sales vs Category Average | Category performance against the average |
| Day of Week Total Sales vs Total Orders | Weekly pattern |
| Monthly Total Sales vs Total Orders | Sales and orders side by side |
| Product × Region matrices | Total Sales, Quantity, Orders and Customers by product and region |

---

## 7. Stage 4: Dashboard

The **Elite Shipping Dashboard** is built entirely with Excel Pivot Charts.

**Controls**
- Slicers: **Region** and **Category**
- Timeline: **Order Date** (by month)

**KPI cards**

| KPI | Value |
|---|---|
| Total Sales | $101,111,318 |
| Total Quantity | 108,087 |
| Total Orders | 19,586 |
| Total Customers | 300 |

**Charts**
1. Top 5 vs Bottom 5 Selling Products (bar)
2. Monthly Total Sales (line)
3. Total Orders vs Order Average (column + line)
4. Total Sales by Region (donut)
5. Category Total Sales vs Category Average (bar)
6. Day of Week Total Sales vs Total Orders (column + line)
7. Monthly Total Sales vs Total Orders (column + line)
8. Total Sales by Category (pie)

![Dashboard](images/dashboard.png)

---

## 8. Key Insights

- **Category:** Clothing leads with 28.78% of sales, followed by Toys (25.10%), Electronics (18.80%), Books (17.00%) and Furniture (10.32%).
- **Top products:** Doll ($8.15M), Shirt ($7.21M), Pants ($6.57M), Tablet ($6.35M) and Shoes ($6.32M).
- **Bottom products:** Sofa ($0.26M) and Headphones ($0.89M) are the weakest sellers, followed by Chair, Notebook and Lego (about $1.9M to $2.1M each).
- **Region:** Sales are evenly spread across regions: Mansoura 20.78%, Aswan 20.34%, Cairo 20.23%, Giza 20.19% and Alexandria 18.46%.
- **Monthly trend:** Sales stay in a narrow band of about $8.3M to $8.7M per month, peaking in May ($8.66M) and lowest in December ($8.29M). Orders per month are very stable at roughly 1,620 to 1,640.
- **Day of week:** Wednesday is the strongest day ($14.86M, 2,899 orders) and Tuesday the weakest ($13.97M, 2,695 orders).

---

## 9. Repository Structure

```
.
├── Data/
│   ├── Customers.csv
│   ├── Products.csv
│   └── Orders/                  # monthly order CSV files
├── Power Query/
│   └── Applied Steps For Power Query.txt
├── Excel/
│   └── Elite Shipping Dashboard.xlsx
├── images/
│   ├── star-schema.png          # Power Pivot data model
│   ├── pivot-tables.png         # Pivot tables sheets
│   └── dashboard.png            # Final dashboard
└── README.md
```

## 10. How to Use

1. Clone or download the repository.
2. Open the Excel workbook.
3. If the file paths differ, update the data sources in **Data → Get Data → Data Source Settings** (the queries currently point to a local folder).
4. Refresh with **Data → Refresh All**.
5. Use the Region and Category slicers and the Order Date timeline to explore the dashboard.

---

*Built with Microsoft Excel (Power Query, Power Pivot, Pivot Tables).*
