# 📊 Sales Data Analysis & Visualization — Excel Project

> An Excel-based sales analytics project using a 1,000-row retail orders dataset, summary tables, visual analysis, and advanced formulas.

## 📌 Project Overview

This project explores retail sales, profit, product categories, customer segments, locations, and payment methods. The workbook is organized into three sheets:

- **`DataSet`** — raw order-level data with 16 fields.
- **`Data Analysis & Visualisation`** — summary tables covering city/state/region sales, category sales and quantity, payment methods, profit by category, customer segments, and more.
- **`Advanced Formulas`** — formula-based order classifications and lookups.

## 📷 Visual Gallery

### Sales by product category
![Total sales by category](assets/sales_by_category.png)

### Sales by region
![Total sales by region](assets/sales_by_region.png)

### Top 10 cities by sales
![Top 10 cities by sales](assets/top_cities_sales.png)

### Sales vs. profit
![Sales vs profit scatter plot](assets/sales_vs_profit.png)

### Monthly sales trend
![Monthly sales trend](assets/monthly_sales_trend.png)

## 🔎 Key Dataset Snapshot

| Metric | Value |
|---|---:|
| Orders / records | 1,000 |
| Total sales | ₹6,974,348.00 |
| Total profit | ₹1,612,382.24 |
| Average sales per order | ₹6,974.35 |
| Cities represented | 15 |
| Product categories | 4 |
| Regions represented | 5 |
| Order date range | 2024-01-03 to 2026-09-27 |

*Figures above are calculated from the `DataSet` sheet. The charts are generated from the same dataset for README display.*

## 💡 What the Analysis Covers

- Compare sales across **cities, states, and regions**.
- Identify high-performing **product categories** by sales and profit.
- Compare quantity across categories and individual products.
- Explore sales collected through **Cash, Credit Card, Debit Card, Net Banking, and UPI**.
- Compare **Consumer, Corporate, and Home Office** segments.
- Inspect the relationship between sales and profit at order level.

## 🧮 Excel Skills Demonstrated

The `Advanced Formulas` sheet includes examples of:

- `IF` and nested `IF` — classify profit/performance and sales bands.
- `AND` — evaluate multiple order conditions.
- `VLOOKUP` — retrieve customer details from the dataset.
- `INDEX` + `MATCH` — retrieve matching order and shipping dates.
- Formula-driven order status and sales-category labels.

Example formulas used in the workbook:

```excel
=IF(C2>200,"Profit","Loss")
=IF(B2>6000,"High",IF(B2>=3000,"Medium","Low"))
=VLOOKUP(A2,DataSet!$1:$1048576,5,FALSE)
=INDEX(DataSet!B:B,MATCH('Advanced  Formulas'!A2,DataSet!A:A,0))
=IF(AND(B2>3000,C2>0),"Good Order","Review")
```

## 🗂️ Repository Structure

```text
excel-sales-analysis/
├── README.md
└── assets/
    ├── sales_by_category.png
    ├── sales_by_region.png
    ├── top_cities_sales.png
    ├── sales_vs_profit.png
    └── monthly_sales_trend.png
```

## 🚀 How to Use

1. Download or clone this repository.
2. Open `MohammadTaki_Practical_Exam.xlsx` in Microsoft Excel or a compatible spreadsheet app.
3. Review the `DataSet` sheet first.
4. Explore summary tables in `Data Analysis & Visualisation`.
5. Inspect formulas in `Advanced  Formulas`.
6. View the charts in this README for a quick visual overview.

## ⚠️ Notes

- Charts in `assets/` are static visualizations for the GitHub README; they are not interactive Excel charts.
- If you edit the source data, regenerate the chart images so the README stays in sync.
- Validate formula results after opening in a spreadsheet application that recalculates formulas.

## 👨‍💻 Author

**Panjwani Mohammad Taqi**

---

⭐ If you find this project useful, feel free to star the repository.
