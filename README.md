# 📊 SuperStore Sales Dashboard — Microsoft Excel

An interactive, single-page **Sales Analysis Dashboard** built entirely in Microsoft Excel (Power Query → Power Pivot Data Model → PivotTables → PivotCharts → Slicers), developed as part of the **Digital Egypt Pioneers Initiative (DEPI)** — Data Analysis Track, 2026.

The dashboard analyzes the classic **Superstore** retail dataset (9,994 sales line items, Jan 2016 – Dec 2019) and turns it into a management cockpit of KPIs, charts, and interactive slicers — with **zero code**.

![Sales Analysis Dashboard](Screenshots/Dashboard.PNG)

---

## 🎯 Objective

Turn raw order-level sales data into clear, actionable answers about:

- 💰 **Overall performance** — total Sales, total Profit, and profit margin
- 🗺️ **Regional performance** — which region sells the most
- 🏷️ **Category profitability** — where the profit actually comes from
- 👥 **Customer segmentation** — who buys (Consumer / Corporate / Home Office)
- 📅 **Monthly trends** — seasonality across the year

---

## 🔢 Dashboard at a Glance (key figures in the shipped workbook)

| KPI / Insight | Value |
|---|---|
| Records analyzed | **9,994** order line items |
| Date coverage | **03 Jan 2016 → 30 Dec 2019** |
| Total **Sales** | **$2,297,200.86** |
| Total **Profit** | **$286,397.02** (≈ **12.47%** margin) |
| **Top Region** | **West** — $725,457.82 (31.6% of sales) |
| Most profitable **Category** | **Technology** — $145,454.95 profit |
| Largest **Segment** | **Consumer** — 50.56% of sales |
| Best **Month** | **November** — $352,461.07 |
| Weakest **Month** | **February** — $59,751.25 |
| Returned orders (sales value) | **$180,504.28** (~7.9% of sales) |

**Sales by Region** — West $725,457.82 · East $678,781.24 · Central $501,239.89 · South $391,721.91
**Profit by Category** — Technology $145,454.95 · Office Supplies $122,490.80 · Furniture $18,451.27
**Sales by Segment** — Consumer 50.56% · Corporate 30.74% · Home Office 18.70%
**Sales by Year** — 2016: $484,247.50 · 2017: $470,532.51 · 2018: $609,205.60 · 2019: $733,215.26

---

## 🧭 How to Use the Dashboard

1. Open `Sales_Analysis_Dashboard.xlsx` and go to the **`DashBoard`** tab (it opens pre-selected; gridlines are hidden and the view is set to 60% zoom for a canvas-like look).
2. Filter with the four **slicers** on the left rail:
   - **Order Date (Month)** (Jan–Dec)
   - **Segment** (Consumer / Corporate / Home Office)
   - **Region** (Central / East / South / West)
   - **Category** (Furniture / Office Supplies / Technology)
3. Every slicer is **connected to all seven pivot tables**, so all charts and the Sales/Profit KPI cards re-filter together. Use multi-select (Ctrl/Cmd-click) or clear a slicer to reset it.
4. The companion **`Pivot Tables`** tab holds the same analysis as titled charts plus its own copy of the four slicers — handy for exploring the underlying tables.

> **Note:** the **Sales** and **Profit** KPI cards are text boxes linked to `'Pivot Tables'!$C$75` / `'Pivot Tables'!$D$75`, so they update live with the slicers. The **Top Region** card is a static label (the West region is the overall leader) and does not re-rank when slicers change.

---

## 🗂️ Workbook Architecture

| Sheet | Purpose |
|---|---|
| **DashBoard** | The presentation canvas: title banner, 3 KPI cards, 4 charts, 4 slicers, icon/panel shapes. Contains no cells of its own — everything is floating objects over a clean gridline-free canvas. |
| **Pivot Tables** | The engine room: 7 PivotTables + 4 titled charts + a copy of the 4 slicers. |
| **Sales (S&R)** | The cleaned dataset loaded by Power Query: an Excel Table named `Sales__S_R` with **9,994 rows × 22 columns**. |
| **Sheet4** *(hidden)* | A diagnostic drill-through dump of the Data Model (`Sum of Profit`, first 1,000 rows), used while building the model. |

### 📄 Dataset schema (`Sales__S_R`, columns A–V)

`Order ID`, `Order Date`, `Ship Date`, `Ship Mode`, `Customer ID`, `Customer Name`, `Segment`, `Country/Region`, `City`, `State`, `Postal Code`, `Region`, `Product ID`, `Category`, `Sub-Category`, `Product Name`, `Sales`, `Quantity`, `Discount`, `Profit`, `Profit Margin`, `Returned`

The **Data Model** adds four calculated date columns used by the slicers/charts: `Order Date (Year)`, `Order Date (Quarter)`, `Order Date (Month)`, `Order Date (Month Index)`.

### 🔄 ETL pipeline (Power Query)

Four queries are embedded in the workbook (see `customXml/item1.xml` → `Formulas/Section1.m`):

1. **Sales Orders** — loads the raw orders table, sets data types, and adds a calculated column **Profit Margin = [Profit] / [Sales]**.
2. **Returns** — loads the returns table, **removes duplicate Order IDs** (`Table.Distinct`) and sorts by Order ID.
3. **People** — loads the regional people/manager reference table.
4. **Sales (S&R)** — the main query: **left-outer joins** `Sales Orders` ⨝ `Returns` on `Order ID`, expands the `Returned` flag, then replaces `null → "Sold"` and `"Yes" → "Returned"`. Its output is loaded both to the `Sales (S&R)` sheet table **and** to the Power Pivot Data Model.

### 📐 PivotTables on `Pivot Tables` (all fed by the Data Model, all slicer-connected)

| PivotTable | Range | Measure |
|---|---|---|
| Regional Sales | A3:B8 | Sum of Sales by Region |
| Category Profitability | A22:B26 | Sum of Profit by Category |
| Customer Segmantation | A37:B41 | % of Sales by Segment |
| Monthly Trends | A55:B68 | Sum of Sales by calendar Month |
| PivotTable2 (Sales total) | C74:C75 | Grand total Sales → feeds the **Sales** KPI card |
| Total Profit | D74:D75 | Grand total Profit → feeds the **Profit** KPI card |
| Top Region | F74:G79 | % share of Sales by Region (West leads) |

### 📈 Charts

- `Pivot Tables` sheet: **Regional Sales** (column), **Category Profit** (column), **Sales per Month** (line), **Customer Sugmentation** (pie with % data labels).
- `DashBoard` sheet: the same four visualizations, un-titled and re-styled to match the dashboard panels (Regional Sales, Category Profit, Monthly Sales, Customer Segmentation).

---

## 🛠️ Tools & Techniques

- **Microsoft Excel** (Microsoft 365)
- **Power Query (M)** — join, dedupe, type-conversion, calculated column
- **Power Pivot / Data Model** — calculated date hierarchy columns, slicer-aware measures
- **PivotTables & PivotCharts**
- **Slicers** (4 caches × 2 sheets, report connections to all pivots)
- **Dashboard design** — shape-based canvas, cell-linked KPI text boxes, icons, gridline-free layout

## 💡 Key Skills Demonstrated

Data cleaning & shaping · relational joins in Power Query · Data Modeling · KPI design · interactive filtering · business storytelling · dashboard UX in Excel

## 📁 Repository Contents

```
sales-analysis-dashboard-excel/
├── Sales_Analysis_Dashboard.xlsx   # The dashboard workbook
├── Screenshots/
│   └── Dashboard.PNG               # Dashboard preview
└── README.md                       # This file
```

## ⚠️ Notes & Limitations

- The workbook ships with all query outputs **loaded**, so it is fully usable as-is. The Power Query sources (`Orders`, `Table3`, `Table7` via `Excel.CurrentWorkbook()`) are **not embedded**, so a full refresh requires re-attaching the original source tables/workbook.
- The **Top Region** KPI is a static text label (see note above).
- Some object names keep their original (slightly misspelled) workbook names, e.g. the `Customer Segmantation` / `Customer Sugmentation` pivots/charts.
- Requires Excel 2010+ for slicers; Data Model features require Excel 2013+ (Microsoft 365 recommended).

## 🎓 Program

**Digital Egypt Pioneers Initiative (DEPI)**
Data Analysis Track | 2026

*Author: Mostafa Mahrous Moghazy*
