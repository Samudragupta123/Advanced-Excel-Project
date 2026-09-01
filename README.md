# E-Commerce Order & Customer Analytics — Advanced Excel Dashboard

An interactive Excel dashboard that turns 100 rows of raw e-commerce order data into a full analytics solution — KPI cards, trend charts, and slicer-based filtering across category, region, customer segment, payment method, and order status.

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/status-complete-brightgreen)

---

## 📊 Project Overview

This project transforms a raw e-commerce order dataset into an interactive, Excel-based analytics workbook. It summarizes revenue, orders, customer activity, product performance, regional performance, payment methods, and customer segments — all surfaced through a single interactive dashboard with KPI cards, charts, and slicers for exploring the data and spotting business trends.

## 🗂️ Repository Contents

| File | Description |
|---|---|
| `Final_project_dashboard_refined.xlsx` | The complete Excel workbook (dashboard, data, and backend calculations) |
| `README.md` | This file |

## 🧾 Dataset Description

| Attribute | Detail |
|---|---|
| Records | 100 orders |
| Fields | 22 columns |
| Customers | 20 unique customers |
| Products | 12 products |
| Categories | 6 categories |
| Regions | 4 regions (East, North, South, West) |
| Cities | 10 cities |
| States | 8 states |
| Payment methods | 5 methods |
| Order statuses | Delivered, Pending, Returned, Cancelled |
| Customer segments | New Customer, Regular Customer, Loyal Customer |
| Date range | January – August 2026 |

## 📁 Workbook Structure

| Sheet | Purpose |
|---|---|
| **Dashboard** | Interactive management dashboard with KPI cards, charts, and slicers |
| **Raw_Data** | Primary 100-row source dataset, stored as the `SalesData` Excel Table |
| **Sheet3** | Supporting summary/pivot output used in the workbook |
| **Sheet5** | Supporting copy/analysis sheet of the dataset |
| **Backend** | Supporting aggregated analysis powering the dashboard visualizations |
| **Data_Dictionary** | Definitions and meanings of all 22 dataset columns |

## ✨ Key Features

- **KPI Dashboard** — Total Revenue, Total Orders, Customers, Quantity Sold, Average Order Value, and Average Rating
- **Revenue Trend** — Monthly revenue visualized across January–August 2026
- **Category Analysis** — Revenue compared across six product categories
- **Product Analysis** — Top products ranked by net revenue
- **Regional Analysis** — Revenue compared across East, North, South, and West
- **Customer Segment Analysis** — Revenue broken down by New, Regular, and Loyal customers
- **Payment Analysis** — Payment method distribution across five methods
- **Interactive Filtering** — Slicers for Category, City, Customer Segment, Order Status, Payment Method, and Region
- **Data Documentation** — A dedicated Data Dictionary sheet explaining every source and calculated field

## 🛠️ Tools & Techniques

- Microsoft Excel with a structured Excel Table (`SalesData`)
- Structured references and formulas for all KPI calculations
- `SUM` — total revenue and quantity
- `COUNTA` — order counts
- `SUMPRODUCT` + `COUNTIF` — distinct customer counting
- `AVERAGE` — average customer rating
- Pivot-style backend summaries by category, month, region, product, payment method, and customer segment
- Charts/PivotCharts for trend, comparison, and distribution analysis
- Interactive slicers for on-sheet filtering

## 🧮 Formula & Calculation Logic

| KPI / Measure | Formula |
|---|---|
| Total Revenue | `=SUM(SalesData[Net Revenue])` |
| Total Orders | `=COUNTA(SalesData[Order ID])` |
| Unique Customers | `=SUMPRODUCT((SalesData[Customer ID]<>"")/COUNTIF(SalesData[Customer ID],SalesData[Customer ID]&""))` |
| Quantity Sold | `=SUM(SalesData[Quantity])` |
| Average Order Value | `=SUM(SalesData[Net Revenue])/COUNTA(SalesData[Order ID])` |
| Average Rating | `=AVERAGE(SalesData[Rating])` |

Calculated source columns:
- **Gross Sales** = Quantity × Unit Price
- **Discount Amount** = Gross Sales × Discount
- **Net Revenue** = Gross Sales − Discount Amount + Shipping Cost

## 📈 Key Analytical Results

| Metric | Result |
|---|---|
| Total Net Revenue | ₹393,143.60 |
| Total Orders | 100 |
| Unique Customers | 20 |
| Quantity Sold | 257 |
| Average Order Value | ₹3,931.44 |
| Average Rating | 3.98 / 5 |
| Highest Revenue Category | Electronics — ₹140,692.10 |
| Highest Revenue Region | West — ₹138,824.45 |
| Highest Revenue Customer Segment | Loyal Customer — ₹141,179.75 |
| Highest Revenue Product | Smart Watch — ₹72,413.25 |
| Highest Revenue Month | July — ₹68,669.85 |

## 🖥️ Dashboard Preview

The dashboard combines KPI cards, a trend chart, and category/product/region/customer/payment visualizations, with slicer objects for interactive filtering.

> **Note:** Slicers are native Excel objects and may not render correctly in non-Excel spreadsheet viewers (e.g., Google Sheets, LibreOffice, GitHub's file preview). Open the workbook in **Microsoft Excel** for the full interactive experience.

 

## 🚀 How to Use

1. Download or clone this repository.
2. Open `Final_project_dashboard_refined.xlsx` in Microsoft Excel.
3. Go to the **Dashboard** sheet.
4. Use the slicers to filter by Category, City, Customer Segment, Order Status, Payment Method, or Region.
5. Refer to the **Data_Dictionary** sheet for definitions of any column.

## 🎓 Learning Outcomes

- Organizing and documenting a structured e-commerce dataset in Excel
- Using structured references and formulas for dynamic KPI calculations
- Summarizing large datasets with PivotTable-style analysis and backend summaries
- Selecting appropriate chart types for trend, comparison, and distribution analysis
- Building interactive, slicer-driven dashboards
- Improving dashboard layout, formatting, and business-oriented data presentation
- Understanding how a Data Dictionary improves clarity and maintainability

## 📌 Conclusion

The E-Commerce Order & Customer Analytics project converts raw order-level data into a structured, interactive Excel analytics solution. The final workbook delivers clear KPIs and multiple visual analyses across revenue, products, categories, regions, customer segments, and payment methods — demonstrating practical application of Advanced Excel skills through documented calculations and slicer-based filtering.

---

**Project Duration:** 4 Days · **Category:** Advanced Excel Dashboard & Analytics
