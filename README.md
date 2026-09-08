# 🛍️ E-Commerce Sales Analytics

Professional **Microsoft Excel** workbook analyzing a full year (2025) of e-commerce sales for a fictional Egyptian store — **400 orders · 8 categories · 32 products**, built entirely with native Excel formulas (no macros).

> **Tool:** Microsoft Excel (SUMIFS, INDEX/MATCH, Conditional Formatting, Charts) | **Data Generation:** PowerShell 5.1 + Excel COM

---

## 📌 Overview

The workbook simulates a complete **Data Analysis → Dashboard** pipeline inside a single `.xlsx` file:

```
Raw_Data  →  Summary  →  Dashboard
 (source)    (logic)     (insights)
```

Every KPI and chart is formula-linked back to the raw sheet, so the dashboard refreshes automatically whenever the source data changes.

---

## 🗂️ Workbook Contents (3 Sheets)

### Sheet 1 — Raw_Data
- **400 simulated e-commerce orders** for 2025 with 13 columns (Order ID, Date, Region, Category, Product, Channel, Customer Type, Units, Unit Price, etc.).
- Formula-driven `Total` column and **3-color scale** conditional formatting for instant visual scanning.

### Sheet 2 — Summary
Live breakdowns, all computed with native formulas:
| View | Formula approach |
|------|------------------|
| Sales by Region | `SUMIFS` |
| Sales by Category | `SUMIFS` |
| Sales by Channel | `SUMIFS` |
| Sales by Month | `SUMIFS` + date logic |
| Sales by Customer Type | `COUNTIFS` / `SUMIFS` |
| Product Performance | `SUMIFS` + ranking |
| Top 5 Products | `LARGE` + `INDEX/MATCH` |

### Sheet 3 — Dashboard
- **6 KPI cards** — Total Sales, Total Orders, Avg Order Value, Top Category, Top Region, Best Channel.
- **4 interactive charts** — category, channel, monthly trend, and product performance.
- Fully **macro-free**: every element is a live Excel formula or native chart.

---

## 🔍 Key Insights

- 400 orders spread across **8 categories** and **32 products**.
- Dashboard highlights best-performing products and channels at a glance.
- `INDEX/MATCH` + `LARGE` power the dynamic **Top 5** list without hardcoding.

---

## 📥 Getting Started

1. **Download** the workbook: [`E-Commerce_Sales_Analytics_v2.xlsx`](E-Commerce_Sales_Analytics_v2.xlsx).
2. Open in **Microsoft Excel** (2016+).
3. Edit any raw value → summaries and dashboard update live.

```text
e-commerce-sales-analytics/
├── E-Commerce_Sales_Analytics_v2.xlsx   # Main workbook (3 sheets)
└── README.md                            # Project documentation
```

---

## 📬 Connect

- **Author:** Mohamed Hany
- **LinkedIn:** [Mohamed Hany Abdelfattah](https://www.linkedin.com/in/mohamed-hany-abdelfattah)
- **GitHub:** [Mohamed-Hany-Abdelfattah](https://github.com/Mohamed-Hany-Abdelfattah)