# Superstore Sales Performance Dashboard

**An End-to-End Retail Analytics Project Using Microsoft Excel and Power BI**

## 1. Project Overview

The Superstore Sales Performance Dashboard is a business intelligence project designed to analyse retail sales, profitability, product categories, customer segments, and monthly sales trends.

The project began in Microsoft Excel, where the source data was inspected, numeric values were corrected, and PivotTables were created to summarise business performance. The prepared dataset was then imported into Power BI, where DAX measures and interactive visualisations were developed.

The objective was to transform raw transaction data into meaningful business insights that can support data-driven decision-making.

## 2. Dashboard Preview

![Superstore Sales Performance Dashboard](screenshots/superstore_dashboard.png)

**Dashboard Screenshot:** [Click here to view the dashboard image](screenshots/superstore_dashboard.png)

## 3. Project Objectives

- Prepare and validate retail sales data for analysis.
- Correct numeric values affected by decimal-separator differences.
- Use Excel PivotTables to summarise sales, profit, and customer performance.
- Identify profitable and loss-making product sub-categories.
- Compare sales and profitability across customer segments.
- Create DAX measures for key performance indicators.
- Develop an interactive Power BI dashboard.
- Translate analytical findings into actionable business recommendations.

## 4. Tools and Technologies

| Tool | Purpose |
|---|---|
| Microsoft Excel | Data preparation, numeric conversion, PivotTables, and preliminary analysis |
| Microsoft Power BI Desktop | Data modelling, DAX calculations, and interactive visualisation |
| DAX | Business measures and performance calculations |
| GitHub | Version control and portfolio presentation |

## 5. Dataset Description

The project uses the Sample Superstore retail dataset, containing **9,994 transaction rows**. The dataset includes information about orders, customers, products, locations, sales, discounts, quantities, and profit.

**Currency note:** The reported financial figures retain the dataset's original currency and should not be interpreted as South African rand.

## 6. Excel Data Preparation

### 6.1 Inspecting the Source Data

The source dataset was opened in Excel and inspected to understand its structure and identify the fields required for analysis. The worksheet was named `Sales`, and the dataset was converted into an Excel Table named `tblSales`.

### 6.2 Correcting Decimal-Separator Differences

The source data used decimal points, while the local Excel regional settings expected decimal commas. The `NUMBERVALUE` function was used to convert the affected values.

For example:

```excel
=NUMBERVALUE(P2;".";",")
```

This formula specifies the decimal point and thousands separator explicitly.

### 6.3 Validating Numeric Columns

Excel's `COUNT` function confirmed 9,994 numeric values in each of these columns:

- `Sales`
- `Profit`
- `Discount`
- `Quantity`

### 6.4 PivotTables and PivotCharts

PivotTables were used to summarise sales, profit, quantities, product categories, sub-categories, and customer segments. PivotCharts helped communicate the results and identify areas requiring further investigation.

## 7. Power BI Dashboard

The prepared `tblSales` table was imported into Power BI Desktop.

The dashboard includes KPI cards, sales by category, profit by sub-category, sales by customer segment, profit margin by segment, sales share by category, monthly sales trends, and interactive slicers for Region, Segment, and Order Date.

### 7.1 DAX Measures

**Total Sales**

```dax
Total Sales = SUM(tblSales[Sales])
```

**Total Profit**

```dax
Total Profit = SUM(tblSales[Profit])
```

**Total Orders**

```dax
Total Orders = DISTINCTCOUNT(tblSales[Order ID])
```

**Profit Margin**

```dax
Profit Margin = DIVIDE([Total Profit], [Total Sales], 0)
```

## 8. Key Business Findings

| Metric | Result |
|---|---:|
| Total Sales | 2,297,200.86 |
| Total Profit | 286,397.02 |
| Transaction Rows | 9,994 |
| Average Discount | 15.59% |
| Profit Margin | 12.47% |

### Product Category Performance

| Category | Sales | Profit | Profit Margin |
|---|---:|---:|---:|
| Technology | 836,154.03 | 145,454.95 | 17.40% |
| Office Supplies | 719,047.03 | 122,490.80 | 17.04% |
| Furniture | 741,999.80 | 18,451.27 | 2.49% |

**Finding:** Technology generated the highest sales and profit. Furniture generated substantial sales but had a considerably lower profit margin.

### Customer Segment Performance

| Segment | Sales | Profit | Profit Margin |
|---|---:|---:|---:|
| Consumer | 1,161,401.34 | 134,119.21 | 11.55% |
| Corporate | 706,146.37 | 91,979.13 | 13.03% |
| Home Office | 429,653.15 | 60,298.68 | 14.03% |

**Finding:** Consumer customers contributed the highest sales and total profit, while Home Office customers achieved the highest profit margin.

### Sub-Category Profitability

- **Copiers:** Highest sub-category profit at 55,617.82.
- **Phones:** Sales of 330,007.05 and profit of 44,515.73.
- **Tables:** Loss of 17,725.48.
- **Bookcases:** Loss of 3,472.56.
- **Supplies:** Loss of 1,189.10.

## 9. Business Recommendations

1. Investigate Furniture profitability by reviewing pricing, discounts, costs, and product mix.
2. Examine loss-making sub-categories, particularly Tables, Bookcases, and Supplies.
3. Investigate the products and customer groups contributing to Technology's strong performance.
4. Study Home Office customers' purchasing patterns and margin sustainability.
5. Monitor monthly sales trends to support planning.
6. Evaluate whether discounts increase sales enough to justify their impact on profit.

These recommendations identify areas for further investigation; the analysis does not independently establish the causes of the results.

## 10. Project Limitations

- The analysis describes historical performance and does not prove why a result occurred.
- The dataset may not contain every operating cost needed to explain product profitability.
- The monthly sales chart is historical analysis, not a forecasting model.
- Business decisions should consider additional context before action is taken.

## 11. Author

**Kefasi Vurumu**

Aspiring Data Analyst | Business Intelligence | Data Analytics
