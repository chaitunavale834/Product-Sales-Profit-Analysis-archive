# Product Sales and Profit Analysis Using Python and Excel

> **B.Sc. Data Science Mini Project**  
> **Domain:** Data Science / Data Analysis  
> **Tools Used:** Microsoft Excel and Python (Pandas, NumPy, Matplotlib, Seaborn, Openpyxl, Python-Docx)  
> **Author GitHub:** [chaitunavale834](https://github.com/chaitunavale834)  
> **Repository:** [Product-Sales-Profit-Analysis](https://github.com/chaitunavale834/Product-Sales-Profit-Analysis)

---

## 📋 Table of Contents
1. [Project Synopsis (From Specification)](#-project-synopsis)
   - [1. Introduction](#1-introduction)
   - [2. Problem Statement](#2-problem-statement)
   - [3. Objectives](#3-objectives)
   - [4. Tools and Technologies](#4-tools-and-technologies)
   - [5. Dataset / Proposed Fields](#5-dataset--proposed-fields)
   - [6. Important Formulas](#6-important-formulas)
   - [7. Methodology](#7-methodology)
   - [8. Expected Output](#8-expected-output)
   - [9. Advantages](#9-advantages)
   - [10. Conclusion](#10-conclusion)
2. [Quick Start & Execution](#-quick-start--execution)
3. [Repository Structure](#-repository-structure)
4. [Visualizations & Key Insights](#-visualizations--key-insights)
5. [Excel Dashboard Overview](#-excel-dashboard-overview)
6. [Formal Project Report](#-formal-project-report)

---

## 📑 Project Synopsis

| Project Title | Product Sales and Profit Analysis Using Python and Excel |
| :--- | :--- |
| **Project Type** | Mini Project |
| **Domain** | Data Science / Data Analysis |
| **Tools Used** | Microsoft Excel and Python |
| **Prepared For** | B.Sc. Data Science Mini Project |

### 1. Introduction
**Product Sales and Profit Analysis** is a data analysis project designed to understand sales performance, cost, revenue, and profit for different products. The project uses Microsoft Excel for data entry, cleaning, formulas, sorting, filtering, and basic visualization, while Python is used for data processing, calculation, statistical analysis, and graphical representation. The analysis helps identify high-selling products, profitable products, low-performing products, and important sales trends.

### 2. Problem Statement
Businesses generate large amounts of sales data, but raw data alone does not clearly show which products are performing well. Manual analysis can be time-consuming and may lead to errors. Therefore, this project aims to analyze product sales data systematically using Excel and Python and present useful information about sales, revenue, expenses, and profit.

### 3. Objectives
- To collect and organize product sales data.
- To calculate total sales, revenue, cost, and profit.
- To compare the performance of different products.
- To identify the best-selling and most profitable products.
- To analyze monthly or category-wise sales trends.
- To use Excel formulas, Pivot Tables, and charts for analysis.
- To use Python for data cleaning, calculations, and visualization.
- To present the results in a simple and understandable form.

### 4. Tools and Technologies
| Tool | Purpose |
| :--- | :--- |
| **Microsoft Excel** | Data entry, cleaning, formulas, Pivot Tables, charts and basic dashboard |
| **Python** | Data cleaning, analysis, calculations and visualization |
| **Pandas** | Reading, manipulating and analyzing tabular data |
| **NumPy** | Numerical calculations |
| **Matplotlib** | Creating graphs and charts |

### 5. Dataset / Proposed Fields
The dataset contains transactional sales fields:
- `Product ID` (e.g., P101 to P114)
- `Product Name` (e.g., Laptop Pro 15, Ergonomic Chair, USB-C Hub)
- `Category` (Electronics, Furniture, Office Supplies, Accessories)
- `Date` / `Order_Date` (Transactions across Jan–Dec 2024)
- `Quantity Sold` (Units per order)
- `Selling Price` (Per unit selling price)
- `Cost Price` (Per unit manufacturing/procurement cost)
- `Total Sales/Revenue` (Calculated revenue per order)
- `Total Cost` (Calculated total cost of goods)
- `Profit` (Net profit generated per order)
- `Profit Margin (%)` (Percentage profit margin)

### 6. Important Formulas
$$\text{Sales Revenue} = \text{Quantity Sold} \times \text{Selling Price}$$

$$\text{Total Cost} = \text{Quantity Sold} \times \text{Cost Price}$$

$$\text{Profit} = \text{Sales Revenue} - \text{Total Cost}$$

$$\text{Profit Margin (\%)} = \left(\frac{\text{Profit}}{\text{Sales Revenue}}\right) \times 100$$

### 7. Methodology
1. **Data Collection:** Collect real or sample retail/commercial sales data.
2. **Data Cleaning:** Detect and eliminate duplicates, handle nulls/missing fields, and remove illogical numeric values (non-positive quantity/prices).
3. **Excel Processing:** Utilize formulas, Pivot Tables, and conditional formatting to summarize sales and profits.
4. **Python Pipeline:** Ingest cleaned CSV data with Pandas, compute KPIs, perform statistical aggregations, and compute financial metrics.
5. **Visualization:** Generate high-resolution charts (bar charts, time-series line charts, category comparisons) with Matplotlib.
6. **Interpretation & Reporting:** Synthesize business insights, ranking best/worst performers, and compile findings into an interactive Excel Dashboard and formal Word report.

### 8. Expected Output
The project produces a comprehensive analysis of:
- **Total Revenue:** ₹6,409,213.93
- **Total Cost:** ₹4,572,730.00
- **Total Profit:** ₹1,836,483.93
- **Overall Profit Margin:** 28.65%
- **Product-Wise Performance:** Ranking 14 products by volume, revenue, and margin.
- **Category-Wise Distribution:** Electronics, Furniture, Accessories, and Office Supplies breakdown.
- **Monthly Trends:** Seasonality patterns across all 12 months.
- **Executive Visual Dashboard:** Native Excel interactive charts, KPI scorecards, and formatted analytical sheets.

### 9. Advantages
- **Easy Business Understanding:** Visual scorecards and charts provide immediate clarity.
- **Identify High-Profit Products:** Highlights which products drive margins vs. which drive sheer volume.
- **Data-Driven Decisions:** Enables optimized pricing, procurement, and inventory planning.
- **Reduces Manual Effort:** End-to-end Python script automates the entire pipeline in seconds.
- **Combined Excel & Python Power:** Demonstrates the synergy of high-speed programmatic analysis with intuitive spreadsheet dashboarding.
- **Visual Reports:** Ready-to-present charts, formatted tables, and comprehensive documentation.

### 10. Conclusion
Product Sales and Profit Analysis using Python and Excel is a practical mini project for learning data analysis. Excel provides an easy environment for organizing and summarizing data, while Python provides powerful programming-based analysis and visualization. By combining both tools, the project converts raw sales records into meaningful information that helps understand product performance and profitability.

---

## 🚀 Quick Start & Execution

### 1. Clone the Repository
```bash
git clone https://github.com/chaitunavale834/Product-Sales-Profit-Analysis.git
cd Product-Sales-Profit-Analysis
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the Automated Pipeline
You can run the complete project with a single command:
```bash
py main.py
```
*(Or `python main.py`)*

This script automatically:
1. Loads and validates raw data from `data/sales_data.csv`.
2. Computes all financial metrics (Revenue, Cost, Profit, Margin).
3. Calculates KPIs and group aggregations.
4. Generates 5 high-resolution charts in `output/`.
5. Builds the multi-sheet analytical Excel file `output/Sales_Analysis.xlsx`.
6. Creates the executive interactive dashboard `dashboard/Product_Sales_Dashboard.xlsx`.
7. Compiles the formal 19-chapter project report `report/Project_Report.docx`.

---

## 📁 Repository Structure

```text
Product-Sales-Profit-Analysis/
├── data/
│   └── sales_data.csv                 # Raw & validated transaction logs (300+ records)
├── output/
│   ├── Sales_Analysis.xlsx            # Cleaned data, aggregations & formulas in Excel
│   ├── product_profit.png             # Horizontal bar chart: Net profit by product
│   ├── monthly_sales.png              # Line chart: Monthly revenue, cost & profit trends
│   ├── category_sales.png             # Grouped bar chart: Category revenue vs profit
│   ├── region_sales.png               # Donut/Bar chart: Regional revenue distribution
│   └── profit_comparison.png          # Overall financial breakdown chart
├── dashboard/
│   └── Product_Sales_Dashboard.xlsx   # Executive KPI cards & native Excel charts
├── report/
│   └── Project_Report.docx            # Complete academic project report (19 chapters)
├── analysis.ipynb                     # Interactive Jupyter Notebook for EDA & analysis
├── main.py                            # Master automated Python analysis pipeline
├── PROJECT_SYNOPSIS.md                # Project synopsis document (PDF specification)
├── requirements.txt                   # Required Python libraries
└── README.md                          # Main repository documentation
```

---

## 📊 Visualizations & Key Insights

| Chart Name | Preview & Description |
| :--- | :--- |
| **`product_profit.png`** | Ranks all 14 products by net profit. Top performer: *Laptop Pro 15* (₹405,450.00 profit). |
| **`category_sales.png`** | Compares gross revenue vs net profit across Electronics, Furniture, Accessories, and Office Supplies. |
| **`monthly_sales.png`** | Longitudinal time series tracking revenue, cost, and profit peaks throughout 2024. |
| **`region_sales.png`** | Analyzes revenue contribution across Central, East, North, South, and West regions. |
| **`profit_comparison.png`** | High-level macro breakdown of Total Revenue, Total Cost, and Net Profit. |

---

## 🖥️ Excel Dashboard Overview
The workbook `dashboard/Product_Sales_Dashboard.xlsx` features:
- **Executive KPI Cards:** Total Revenue, Total Cost, Total Profit, Profit Margin %, Total Orders, and Units Sold.
- **Dedicated Worksheets:** `Dashboard`, `Raw_Data`, `Cleaned_Data`, `Product_Analysis`, `Category_Analysis`, `Monthly_Analysis`, `Region_Analysis`.
- **Dynamic Charts:** Native Excel bar charts and line graphs connected to summary tables.
- **Slicer & PivotTable Guidance:** Built-in instructions for interactive filtering.

---

## 📝 Formal Project Report
A complete academic report is located in `report/Project_Report.docx`, including:
- Cover page, Certificate, and Acknowledgements.
- Abstract and Table of Contents.
- Chapters covering Introduction, System Architecture, Mathematical Formulations, EDA, Findings, Limitations, and Future Scope.

---

## 👨‍💻 Author & Repository Link
- **Author:** [chaitunavale834](https://github.com/chaitunavale834)
- **GitHub Repository:** [https://github.com/chaitunavale834/Product-Sales-Profit-Analysis](https://github.com/chaitunavale834/Product-Sales-Profit-Analysis)
