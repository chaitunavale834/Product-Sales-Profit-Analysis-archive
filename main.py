"""
========================================================================================
Project: Product Sales and Profit Analysis Using Python and Excel
Academic Level: B.Sc. Data Science Mini Project
Module: python/main.py
Description: End-to-end automated pipeline for data cleaning, mathematical calculation,
             business aggregation, professional visualization, Excel workbook generation,
             interactive dashboard creation, and complete project report compilation.
========================================================================================
"""

import os
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --------------------------------------------------------------------------------------
# 1. PATH CONFIGURATION
# --------------------------------------------------------------------------------------
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
# Check if running from inside python folder or project folder
if os.path.basename(CURRENT_DIR) == "python":
    PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
else:
    PROJECT_ROOT = CURRENT_DIR

DATA_DIR = os.path.join(PROJECT_ROOT, "data")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")
DASHBOARD_DIR = os.path.join(PROJECT_ROOT, "dashboard")
REPORT_DIR = os.path.join(PROJECT_ROOT, "report")

# Ensure required output directories exist
for folder in [OUTPUT_DIR, DASHBOARD_DIR, REPORT_DIR]:
    os.makedirs(folder, exist_ok=True)

CSV_FILE_PATH = os.path.join(DATA_DIR, "sales_data.csv")

# --------------------------------------------------------------------------------------
# 2. DATA CLEANING & VALIDATION MODULE
# --------------------------------------------------------------------------------------
def clean_and_prepare_data(csv_path):
    print("=" * 70)
    print("STEP 1: LOADING & CLEANING RAW DATASET")
    print("=" * 70)
    
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Source file not found at: {csv_path}")

    raw_df = pd.read_csv(csv_path)
    total_raw_rows = len(raw_df)
    print(f"[*] Successfully loaded raw dataset: {total_raw_rows} records.")

    # 1. Check Missing Values
    null_counts = raw_df.isnull().sum()
    print("[*] Missing values check:")
    for col, count in null_counts.items():
        if count > 0:
            print(f"    - Column '{col}': {count} missing value(s)")
    
    # 2. Check Duplicates
    duplicate_mask = raw_df.duplicated()
    num_duplicates = duplicate_mask.sum()
    print(f"[*] Duplicate records found: {num_duplicates}")

    # 3. Numeric & Logic Validation
    invalid_qty_mask = (raw_df["Quantity"].isna()) | (raw_df["Quantity"] <= 0)
    invalid_price_mask = (raw_df["Selling_Price"].isna()) | (raw_df["Selling_Price"] <= 0) | (raw_df["Cost_Price"].isna()) | (raw_df["Cost_Price"] <= 0)
    invalid_region_mask = raw_df["Region"].isna()

    print(f"[*] Non-positive / invalid Quantity records: {invalid_qty_mask.sum()}")
    print(f"[*] Non-positive / invalid Price records: {invalid_price_mask.sum()}")
    print(f"[*] Missing Region records: {invalid_region_mask.sum()}")

    # Begin cleaning sequence
    clean_df = raw_df.copy()

    # Drop exact duplicates
    clean_df = clean_df.drop_duplicates()

    # Drop records missing critical categoricals or invalid numeric values
    clean_df = clean_df.dropna(subset=["Order_ID", "Order_Date", "Product_ID", "Product_Name", "Category", "Region", "Quantity", "Selling_Price", "Cost_Price"])
    clean_df = clean_df[(clean_df["Quantity"] > 0) & (clean_df["Selling_Price"] > 0) & (clean_df["Cost_Price"] > 0)]

    # Validate and standardize Date format
    clean_df["Order_Date"] = pd.to_datetime(clean_df["Order_Date"])
    clean_df = clean_df.sort_values("Order_Date").reset_index(drop=True)

    # Convert Quantity to integer
    clean_df["Quantity"] = clean_df["Quantity"].astype(int)

    total_clean_rows = len(clean_df)
    affected_rows = total_raw_rows - total_clean_rows
    print(f"[+] Data Cleaning complete:")
    print(f"    - Initial records: {total_raw_rows}")
    print(f"    - Removed records: {affected_rows} (duplicates & invalid entries)")
    print(f"    - Final validated records: {total_clean_rows}")
    print("-" * 70)

    # Document cleaning metrics dictionary
    cleaning_audit = {
        "total_raw": total_raw_rows,
        "duplicates_removed": int(num_duplicates),
        "missing_region": int(invalid_region_mask.sum()),
        "invalid_quantity": int(invalid_qty_mask.sum()),
        "invalid_price": int(invalid_price_mask.sum()),
        "total_clean": total_clean_rows,
        "records_retained_pct": round((total_clean_rows / total_raw_rows) * 100, 2)
    }

    return raw_df, clean_df, cleaning_audit

# --------------------------------------------------------------------------------------
# 3. CORE CALCULATIONS
# --------------------------------------------------------------------------------------
def perform_calculations(df):
    print("=" * 70)
    print("STEP 2: PERFORMING BUSINESS FINANCIAL CALCULATIONS")
    print("=" * 70)
    
    df = df.copy()

    # Revenue = Quantity * Selling_Price
    df["Revenue"] = df["Quantity"] * df["Selling_Price"]

    # Total_Cost = Quantity * Cost_Price
    df["Total_Cost"] = df["Quantity"] * df["Cost_Price"]

    # Profit = Revenue - Total_Cost
    df["Profit"] = df["Revenue"] - df["Total_Cost"]

    # Profit_Margin = (Profit / Revenue) * 100 (Safe division handled with numpy)
    df["Profit_Margin"] = np.where(df["Revenue"] > 0, (df["Profit"] / df["Revenue"]) * 100.0, 0.0)

    # Date breakdown
    df["Year"] = df["Order_Date"].dt.year
    df["Month_Num"] = df["Order_Date"].dt.month
    df["Month_Name"] = df["Order_Date"].dt.strftime("%b")
    df["Year_Month"] = df["Order_Date"].dt.strftime("%Y-%m")

    # Round currency and margin values to 2 decimals
    df["Revenue"] = df["Revenue"].round(2)
    df["Total_Cost"] = df["Total_Cost"].round(2)
    df["Profit"] = df["Profit"].round(2)
    df["Profit_Margin"] = df["Profit_Margin"].round(2)

    print("[+] Revenue, Total Cost, Profit, and Profit Margin calculated successfully.")
    return df

# --------------------------------------------------------------------------------------
# 4. PYTHON STATISTICAL ANALYSIS & AGGREGATIONS
# --------------------------------------------------------------------------------------
def analyze_data(df):
    print("=" * 70)
    print("STEP 3: COMPUTING KEY PERFORMANCE INDICATORS & AGGREGATIONS")
    print("=" * 70)

    # Overall KPIs
    total_revenue = df["Revenue"].sum()
    total_cost = df["Total_Cost"].sum()
    total_profit = df["Profit"].sum()
    overall_profit_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0.0
    total_orders = df["Order_ID"].nunique()
    total_quantity = df["Quantity"].sum()
    avg_order_value = df["Revenue"].mean()

    kpis = {
        "Total_Revenue": round(total_revenue, 2),
        "Total_Cost": round(total_cost, 2),
        "Total_Profit": round(total_profit, 2),
        "Overall_Profit_Margin": round(overall_profit_margin, 2),
        "Total_Orders": int(total_orders),
        "Total_Quantity_Sold": int(total_quantity),
        "Avg_Order_Value": round(avg_order_value, 2)
    }

    print(f"[*] Total Revenue:        INR {kpis['Total_Revenue']:,.2f}")
    print(f"[*] Total Cost:           INR {kpis['Total_Cost']:,.2f}")
    print(f"[*] Total Profit:         INR {kpis['Total_Profit']:,.2f}")
    print(f"[*] Overall Profit Margin: {kpis['Overall_Profit_Margin']}%")
    print(f"[*] Total Orders:          {kpis['Total_Orders']:,}")
    print(f"[*] Total Quantity Sold:   {kpis['Total_Quantity_Sold']:,} units")
    print(f"[*] Average Order Value:   INR {kpis['Avg_Order_Value']:,.2f}")

    # 1. Product-wise Analysis
    product_summary = df.groupby(["Product_ID", "Product_Name", "Category"]).agg(
        Orders=("Order_ID", "count"),
        Quantity_Sold=("Quantity", "sum"),
        Revenue=("Revenue", "sum"),
        Total_Cost=("Total_Cost", "sum"),
        Profit=("Profit", "sum")
    ).reset_index()
    product_summary["Profit_Margin"] = np.where(
        product_summary["Revenue"] > 0,
        (product_summary["Profit"] / product_summary["Revenue"] * 100).round(2),
        0.0
    )
    product_summary = product_summary.sort_values("Profit", ascending=False).reset_index(drop=True)

    # 2. Category-wise Analysis
    category_summary = df.groupby("Category").agg(
        Orders=("Order_ID", "count"),
        Quantity_Sold=("Quantity", "sum"),
        Revenue=("Revenue", "sum"),
        Total_Cost=("Total_Cost", "sum"),
        Profit=("Profit", "sum")
    ).reset_index()
    category_summary["Profit_Margin"] = np.where(
        category_summary["Revenue"] > 0,
        (category_summary["Profit"] / category_summary["Revenue"] * 100).round(2),
        0.0
    )
    category_summary["Revenue_Share_Pct"] = (category_summary["Revenue"] / total_revenue * 100).round(2)
    category_summary = category_summary.sort_values("Revenue", ascending=False).reset_index(drop=True)

    # 3. Region-wise Analysis
    region_summary = df.groupby("Region").agg(
        Orders=("Order_ID", "count"),
        Quantity_Sold=("Quantity", "sum"),
        Revenue=("Revenue", "sum"),
        Total_Cost=("Total_Cost", "sum"),
        Profit=("Profit", "sum")
    ).reset_index()
    region_summary["Profit_Margin"] = np.where(
        region_summary["Revenue"] > 0,
        (region_summary["Profit"] / region_summary["Revenue"] * 100).round(2),
        0.0
    )
    region_summary = region_summary.sort_values("Revenue", ascending=False).reset_index(drop=True)

    # 4. Monthly Analysis
    monthly_summary = df.groupby(["Year_Month", "Month_Num", "Month_Name"]).agg(
        Orders=("Order_ID", "count"),
        Quantity_Sold=("Quantity", "sum"),
        Revenue=("Revenue", "sum"),
        Total_Cost=("Total_Cost", "sum"),
        Profit=("Profit", "sum")
    ).reset_index()
    monthly_summary = monthly_summary.sort_values("Year_Month").reset_index(drop=True)
    monthly_summary["Profit_Margin"] = np.where(
        monthly_summary["Revenue"] > 0,
        (monthly_summary["Profit"] / monthly_summary["Revenue"] * 100).round(2),
        0.0
    )

    # Performance Highlights
    best_selling_by_qty = product_summary.sort_values("Quantity_Sold", ascending=False).iloc[0]
    best_selling_by_rev = product_summary.sort_values("Revenue", ascending=False).iloc[0]
    most_profitable = product_summary.sort_values("Profit", ascending=False).iloc[0]
    lowest_performing = product_summary.sort_values("Profit", ascending=True).iloc[0]

    highlights = {
        "best_selling_by_qty": best_selling_by_qty,
        "best_selling_by_rev": best_selling_by_rev,
        "most_profitable": most_profitable,
        "lowest_performing": lowest_performing
    }

    print("\n[*] Key Highlights:")
    print(f"    - Best-Selling Product (by Units):   {best_selling_by_qty['Product_Name']} ({best_selling_by_qty['Quantity_Sold']} units)")
    print(f"    - Best-Selling Product (by Revenue): {best_selling_by_rev['Product_Name']} (INR {best_selling_by_rev['Revenue']:,.2f})")
    print(f"    - Most Profitable Product:           {most_profitable['Product_Name']} (INR {most_profitable['Profit']:,.2f})")
    print(f"    - Lowest-Performing Product:         {lowest_performing['Product_Name']} (INR {lowest_performing['Profit']:,.2f}) [Criteria: Lowest Total Profit]")

    return {
        "kpis": kpis,
        "product_summary": product_summary,
        "category_summary": category_summary,
        "region_summary": region_summary,
        "monthly_summary": monthly_summary,
        "highlights": highlights
    }

# --------------------------------------------------------------------------------------
# 5. VISUALIZATION GENERATION (MATPLOTLIB)
# --------------------------------------------------------------------------------------
def generate_visualizations(analysis_results):
    print("=" * 70)
    print("STEP 4: GENERATING PROFESSIONAL VISUALIZATION CHARTS")
    print("=" * 70)

    # Set aesthetics
    plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")
    plt.rcParams["font.sans-serif"] = "DejaVu Sans"
    plt.rcParams["axes.edgecolor"] = "#cccccc"
    plt.rcParams["axes.linewidth"] = 0.8

    prod_df = analysis_results["product_summary"]
    cat_df = analysis_results["category_summary"]
    reg_df = analysis_results["region_summary"]
    mon_df = analysis_results["monthly_summary"]

    chart_paths = {}

    # Chart 1: Product-wise Profit (Horizontal Bar Chart)
    fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
    sorted_prod = prod_df.sort_values("Profit", ascending=True)
    colors = ["#2b5c8f" if p > 0 else "#d9534f" for p in sorted_prod["Profit"]]
    bars = ax.barh(sorted_prod["Product_Name"], sorted_prod["Profit"] / 1000, color=colors, edgecolor="black", height=0.65)
    ax.set_title("Product-Wise Net Profit (INR in Thousands)", fontsize=14, fontweight="bold", pad=15)
    ax.set_xlabel("Net Profit (Thousands INR)", fontsize=11, fontweight="bold")
    ax.set_ylabel("Product Name", fontsize=11, fontweight="bold")
    ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f"{int(x)}k"))

    # Add data labels
    for bar in bars:
        width = bar.get_width()
        offset = 5 if width >= 0 else -15
        ax.annotate(f"INR {width:.1f}k",
                    xy=(width, bar.get_y() + bar.get_height() / 2),
                    xytext=(offset, 0), textcoords="offset points",
                    ha='left' if width >= 0 else 'right', va='center',
                    fontsize=9, fontweight="bold")

    plt.tight_layout()
    p1 = os.path.join(OUTPUT_DIR, "product_profit.png")
    plt.savefig(p1, bbox_inches="tight")
    plt.close()
    chart_paths["product_profit"] = p1
    print(f"[+] Saved: {p1}")

    # Chart 2: Category-wise Revenue & Profit (Grouped Bar Chart)
    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=300)
    x = np.arange(len(cat_df))
    width = 0.35

    bar1 = ax.bar(x - width/2, cat_df["Revenue"] / 100000, width, label="Revenue", color="#1f77b4", edgecolor="#0e4b75")
    bar2 = ax.bar(x + width/2, cat_df["Profit"] / 100000, width, label="Net Profit", color="#2ca02c", edgecolor="#1b631b")

    ax.set_title("Category-Wise Performance: Revenue vs. Net Profit (INR in Lakhs)", fontsize=13, fontweight="bold", pad=15)
    ax.set_ylabel("Amount (INR in Lakhs)", fontsize=11, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(cat_df["Category"], fontsize=10, fontweight="bold")
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda val, p: f"{val:.1f}L"))
    ax.legend(frameon=True, facecolor="white", edgecolor="#cccccc")

    for bars in [bar1, bar2]:
        for b in bars:
            h = b.get_height()
            ax.annotate(f"{h:.2f}L",
                        xy=(b.get_x() + b.get_width() / 2, h),
                        xytext=(0, 4), textcoords="offset points",
                        ha='center', va='bottom', fontsize=8.5, fontweight="bold")

    plt.tight_layout()
    p2 = os.path.join(OUTPUT_DIR, "category_sales.png")
    plt.savefig(p2, bbox_inches="tight")
    plt.close()
    chart_paths["category_sales"] = p2
    print(f"[+] Saved: {p2}")

    # Chart 3: Monthly Sales Trend Line Chart
    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)
    ax.plot(mon_df["Month_Name"], mon_df["Revenue"] / 1000, marker="o", linewidth=2.5, color="#0d6efd", label="Revenue", markersize=6)
    ax.plot(mon_df["Month_Name"], mon_df["Total_Cost"] / 1000, marker="s", linewidth=2, linestyle="--", color="#dc3545", label="Total Cost", markersize=5)
    ax.plot(mon_df["Month_Name"], mon_df["Profit"] / 1000, marker="^", linewidth=2.5, color="#198754", label="Net Profit", markersize=6)

    ax.fill_between(mon_df["Month_Name"], mon_df["Profit"] / 1000, color="#198754", alpha=0.15)
    ax.set_title("Monthly Sales, Cost & Profit Trends (2024)", fontsize=13, fontweight="bold", pad=15)
    ax.set_xlabel("Month", fontsize=11, fontweight="bold")
    ax.set_ylabel("Amount (INR Thousands)", fontsize=11, fontweight="bold")
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, p: f"{int(v)}k"))
    ax.legend(loc="upper left", frameon=True)
    ax.grid(True, linestyle=":", alpha=0.6)

    plt.tight_layout()
    p3 = os.path.join(OUTPUT_DIR, "monthly_sales.png")
    plt.savefig(p3, bbox_inches="tight")
    plt.close()
    chart_paths["monthly_sales"] = p3
    print(f"[+] Saved: {p3}")

    # Chart 4: Region-wise Sales Performance
    fig, ax = plt.subplots(figsize=(8.5, 5), dpi=300)
    palette = ["#4e79a7", "#f28e2c", "#e15759", "#76b7b2", "#59a14f"]
    bars = ax.bar(reg_df["Region"], reg_df["Revenue"] / 100000, color=palette[:len(reg_df)], edgecolor="black", width=0.55)
    ax.set_title("Region-Wise Sales Revenue Distribution", fontsize=13, fontweight="bold", pad=15)
    ax.set_xlabel("Territory / Region", fontsize=11, fontweight="bold")
    ax.set_ylabel("Revenue (INR in Lakhs)", fontsize=11, fontweight="bold")
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, p: f"{v:.1f}L"))

    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"INR {h:.2f}L",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 4), textcoords="offset points",
                    ha='center', va='bottom', fontsize=9, fontweight="bold")

    plt.tight_layout()
    p4 = os.path.join(OUTPUT_DIR, "region_sales.png")
    plt.savefig(p4, bbox_inches="tight")
    plt.close()
    chart_paths["region_sales"] = p4
    print(f"[+] Saved: {p4}")

    # Chart 5: Profit Comparison (Waterfall / Component Breakdown)
    fig, ax = plt.subplots(figsize=(9, 5.2), dpi=300)
    kpis = analysis_results["kpis"]
    components = ["Total Revenue", "Total Cost", "Net Profit"]
    values = [kpis["Total_Revenue"] / 100000, kpis["Total_Cost"] / 100000, kpis["Total_Profit"] / 100000]
    colors = ["#007acc", "#d9534f", "#28a745"]

    bars = ax.bar(components, values, color=colors, width=0.45, edgecolor="black")
    ax.set_title("Overall Financial Structure: Revenue, Cost & Profit", fontsize=13, fontweight="bold", pad=15)
    ax.set_ylabel("Amount (INR in Lakhs)", fontsize=11, fontweight="bold")
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, p: f"{v:.1f}L"))

    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"INR {h:.2f} Lakhs",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 5), textcoords="offset points",
                    ha='center', va='bottom', fontsize=9.5, fontweight="bold")

    plt.tight_layout()
    p5 = os.path.join(OUTPUT_DIR, "profit_comparison.png")
    plt.savefig(p5, bbox_inches="tight")
    plt.close()
    chart_paths["profit_comparison"] = p5
    print(f"[+] Saved: {p5}")

    return chart_paths

# --------------------------------------------------------------------------------------
# 6. EXCEL ANALYTICAL WORKBOOK GENERATION (output/Sales_Analysis.xlsx)
# --------------------------------------------------------------------------------------
def generate_sales_analysis_workbook(cleaned_df, analysis_results):
    print("=" * 70)
    print("STEP 5: CREATING ANALYTICAL WORKBOOK (Sales_Analysis.xlsx)")
    print("=" * 70)

    target_path = os.path.join(OUTPUT_DIR, "Sales_Analysis.xlsx")
    wb = openpyxl.Workbook()
    
    # Styles
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Sheet 1: Executive KPI Summary
    ws_kpi = wb.active
    ws_kpi.title = "Executive_Summary"
    ws_kpi.views.sheetView[0].showGridLines = True

    ws_kpi.merge_cells("A1:D1")
    title_cell = ws_kpi["A1"]
    title_cell.value = "PRODUCT SALES & PROFIT ANALYSIS - EXECUTIVE SUMMARY"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_kpi.row_dimensions[1].height = 30

    kpi_data = [
        ("Metric Description", "Value", "Unit / Format"),
        ("Total Sales Revenue", analysis_results["kpis"]["Total_Revenue"], "INR (Currency)"),
        ("Total Cost of Goods Sold", analysis_results["kpis"]["Total_Cost"], "INR (Currency)"),
        ("Total Operating Profit", analysis_results["kpis"]["Total_Profit"], "INR (Currency)"),
        ("Overall Profit Margin", f"{analysis_results['kpis']['Overall_Profit_Margin']}%", "Percentage"),
        ("Total Volume Sold", analysis_results["kpis"]["Total_Quantity_Sold"], "Units"),
        ("Total Valid Orders Processed", analysis_results["kpis"]["Total_Orders"], "Orders"),
        ("Average Order Value (AOV)", analysis_results["kpis"]["Avg_Order_Value"], "INR (Currency)"),
    ]

    for r_idx, row in enumerate(kpi_data, start=3):
        for c_idx, val in enumerate(row, start=1):
            cell = ws_kpi.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if r_idx == 3:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.font = bold_font if c_idx == 1 else data_font
                if c_idx == 2 and isinstance(val, (int, float)):
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right")
                else:
                    cell.alignment = Alignment(horizontal="left" if c_idx == 1 else "center")

    # Sheet 2: Product Performance
    ws_prod = wb.create_sheet(title="Product_Performance")
    ws_prod.views.sheetView[0].showGridLines = True
    prod_summary = analysis_results["product_summary"]
    
    headers = ["Product ID", "Product Name", "Category", "Orders", "Quantity Sold", "Revenue (INR)", "Total Cost (INR)", "Profit (INR)", "Profit Margin %"]
    ws_prod.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws_prod.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for _, row in prod_summary.iterrows():
        ws_prod.append([
            row["Product_ID"], row["Product_Name"], row["Category"],
            int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"])
        ])

    for row in ws_prod.iter_rows(min_row=2, max_row=len(prod_summary) + 1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [6, 7, 8]:
                cell.number_format = "#,##0.00"
            elif cell.column == 9:
                cell.number_format = "0.00%"
                cell.value = cell.value / 100.0

    # Sheet 3: Category Performance
    ws_cat = wb.create_sheet(title="Category_Performance")
    ws_cat.views.sheetView[0].showGridLines = True
    cat_summary = analysis_results["category_summary"]
    cat_headers = ["Category", "Orders", "Quantity Sold", "Revenue (INR)", "Total Cost (INR)", "Profit (INR)", "Profit Margin %", "Revenue Share %"]
    ws_cat.append(cat_headers)
    for c_idx in range(1, len(cat_headers) + 1):
        cell = ws_cat.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for _, row in cat_summary.iterrows():
        ws_cat.append([
            row["Category"], int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0, float(row["Revenue_Share_Pct"]) / 100.0
        ])

    for row in ws_cat.iter_rows(min_row=2, max_row=len(cat_summary) + 1, min_col=1, max_col=len(cat_headers)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [4, 5, 6]:
                cell.number_format = "#,##0.00"
            elif cell.column in [7, 8]:
                cell.number_format = "0.00%"

    # Sheet 4: Monthly Trends
    ws_mon = wb.create_sheet(title="Monthly_Trends")
    ws_mon.views.sheetView[0].showGridLines = True
    mon_summary = analysis_results["monthly_summary"]
    mon_headers = ["Year-Month", "Month", "Orders", "Units Sold", "Revenue (INR)", "Total Cost (INR)", "Profit (INR)", "Profit Margin %"]
    ws_mon.append(mon_headers)
    for c_idx in range(1, len(mon_headers) + 1):
        cell = ws_mon.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for _, row in mon_summary.iterrows():
        ws_mon.append([
            row["Year_Month"], row["Month_Name"], int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0
        ])

    for row in ws_mon.iter_rows(min_row=2, max_row=len(mon_summary) + 1, min_col=1, max_col=len(mon_headers)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [5, 6, 7]:
                cell.number_format = "#,##0.00"
            elif cell.column == 8:
                cell.number_format = "0.00%"

    # Sheet 5: Regional Breakdown
    ws_reg = wb.create_sheet(title="Regional_Breakdown")
    ws_reg.views.sheetView[0].showGridLines = True
    reg_summary = analysis_results["region_summary"]
    reg_headers = ["Region", "Orders", "Units Sold", "Revenue (INR)", "Total Cost (INR)", "Profit (INR)", "Profit Margin %"]
    ws_reg.append(reg_headers)
    for c_idx in range(1, len(reg_headers) + 1):
        cell = ws_reg.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for _, row in reg_summary.iterrows():
        ws_reg.append([
            row["Region"], int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0
        ])

    for row in ws_reg.iter_rows(min_row=2, max_row=len(reg_summary) + 1, min_col=1, max_col=len(reg_headers)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [4, 5, 6]:
                cell.number_format = "#,##0.00"
            elif cell.column == 7:
                cell.number_format = "0.00%"

    # Sheet 6: Cleaned Dataset
    ws_data = wb.create_sheet(title="Cleaned_Data")
    ws_data.views.sheetView[0].showGridLines = True
    data_headers = ["Order_ID", "Order_Date", "Product_ID", "Product_Name", "Category", "Region", "Quantity", "Selling_Price", "Cost_Price", "Revenue", "Total_Cost", "Profit", "Profit_Margin"]
    ws_data.append(data_headers)
    for c_idx in range(1, len(data_headers) + 1):
        cell = ws_data.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for _, row in cleaned_df.iterrows():
        ws_data.append([
            row["Order_ID"], row["Order_Date"].strftime("%Y-%m-%d"), row["Product_ID"],
            row["Product_Name"], row["Category"], row["Region"],
            int(row["Quantity"]), float(row["Selling_Price"]), float(row["Cost_Price"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0
        ])

    for row in ws_data.iter_rows(min_row=2, max_row=len(cleaned_df) + 1, min_col=1, max_col=len(data_headers)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [8, 9, 10, 11, 12]:
                cell.number_format = "#,##0.00"
            elif cell.column == 13:
                cell.number_format = "0.00%"

    # Auto-adjust column widths on all sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(target_path)
    print(f"[+] Saved Analytical Workbook: {target_path}")

# --------------------------------------------------------------------------------------
# 7. PROFESSIONAL EXCEL DASHBOARD WORKBOOK (dashboard/Product_Sales_Dashboard.xlsx)
# --------------------------------------------------------------------------------------
def generate_interactive_dashboard(raw_df, clean_df, analysis_results):
    print("=" * 70)
    print("STEP 6: BUILDING INTERACTIVE EXCEL DASHBOARD")
    print("=" * 70)

    target_path = os.path.join(DASHBOARD_DIR, "Product_Sales_Dashboard.xlsx")
    wb = openpyxl.Workbook()

    # 1. Colors & Design Palette
    brand_dark = "1B365D"       # Deep Navy
    brand_primary = "2A5C9A"    # Corporate Blue
    accent_green = "107C41"     # Excel Green
    kpi_card_fill = "F0F4F8"    # Soft Gray-Blue
    card_border_color = "B0C4DE"
    white = "FFFFFF"

    title_font = Font(name="Segoe UI", size=18, bold=True, color=brand_dark)
    subtitle_font = Font(name="Segoe UI", size=10, italic=True, color="555555")
    kpi_title_font = Font(name="Segoe UI", size=9, bold=True, color="666666")
    kpi_value_font = Font(name="Segoe UI", size=16, bold=True, color=brand_dark)
    table_head_font = Font(name="Segoe UI", size=10, bold=True, color=white)
    cell_font = Font(name="Segoe UI", size=9)
    bold_cell_font = Font(name="Segoe UI", size=9, bold=True)

    border_thin = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )
    border_card = Border(
        left=Side(style='medium', color=card_border_color),
        right=Side(style='medium', color=card_border_color),
        top=Side(style='medium', color=card_border_color),
        bottom=Side(style='medium', color=card_border_color)
    )

    # -------------------------------------------------------------------------
    # Sheet 1: Dashboard
    # -------------------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_dash.merge_cells("A1:K1")
    t_cell = ws_dash["A1"]
    t_cell.value = "PRODUCT SALES AND PROFIT PERFORMANCE DASHBOARD"
    t_cell.font = title_font
    t_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws_dash.row_dimensions[1].height = 32

    ws_dash.merge_cells("A2:K2")
    sub_cell = ws_dash["A2"]
    sub_cell.value = "B.Sc. Data Science Project | Real-time Business Metrics, Group Summaries & Visual Trends"
    sub_cell.font = subtitle_font
    sub_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws_dash.row_dimensions[2].height = 18

    # 6 KPI Cards Setup
    # Layout across Columns A to L, Rows 4 to 6
    kpis = [
        {"title": "TOTAL SALES REVENUE", "formula": "=SUM(Cleaned_Data!J2:J301)", "format": "INR #,##0.00", "cols": ("B4", "C5")},
        {"title": "TOTAL COST OF GOODS", "formula": "=SUM(Cleaned_Data!K2:K301)", "format": "INR #,##0.00", "cols": ("D4", "E5")},
        {"title": "NET OPERATING PROFIT", "formula": "=SUM(Cleaned_Data!L2:L301)", "format": "INR #,##0.00", "cols": ("F4", "G5")},
        {"title": "OVERALL PROFIT MARGIN", "formula": "=(F5/B5)*100", "format": "0.00\"%\"", "cols": ("H4", "I5")},
        {"title": "TOTAL UNITS SOLD", "formula": "=SUM(Cleaned_Data!G2:G301)", "format": "#,##0\" units\"", "cols": ("J4", "K5")},
        {"title": "TOTAL ORDERS", "formula": "=COUNTA(Cleaned_Data!A2:A301)", "format": "#,##0\" orders\"", "cols": ("L4", "M5")},
    ]

    card_positions = [
        ("A4", "B5", "TOTAL REVENUE", "=SUM(Cleaned_Data!J2:J301)", "INR #,##0.00"),
        ("C4", "D5", "TOTAL COST", "=SUM(Cleaned_Data!K2:K301)", "INR #,##0.00"),
        ("E4", "F5", "TOTAL PROFIT", "=SUM(Cleaned_Data!L2:L301)", "INR #,##0.00"),
        ("G4", "H5", "PROFIT MARGIN", "=(E5/A5)", "0.00%"),
        ("I4", "J5", "TOTAL UNITS SOLD", "=SUM(Cleaned_Data!G2:G301)", "#,##0"),
        ("K4", "L5", "TOTAL ORDERS", "=COUNTA(Cleaned_Data!A2:A301)", "#,##0"),
    ]

    for top_left, bottom_right, title_text, form_str, num_fmt in card_positions:
        ws_dash.merge_cells(f"{top_left[0]}4:{bottom_right[0]}4")
        t_c = ws_dash[f"{top_left[0]}4"]
        t_c.value = title_text
        t_c.font = kpi_title_font
        t_c.fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
        t_c.alignment = Alignment(horizontal="center", vertical="center")

        ws_dash.merge_cells(f"{top_left[0]}5:{bottom_right[0]}5")
        v_c = ws_dash[f"{top_left[0]}5"]
        v_c.value = form_str
        v_c.font = kpi_value_font
        v_c.fill = PatternFill(start_color=kpi_card_fill, end_color=kpi_card_fill, fill_type="solid")
        v_c.alignment = Alignment(horizontal="center", vertical="center")
        v_c.number_format = num_fmt

    ws_dash.row_dimensions[4].height = 20
    ws_dash.row_dimensions[5].height = 30

    # Section 1: Category Summary Table on Dashboard (Row 8)
    cat_start_row = 8
    ws_dash.cell(row=cat_start_row, column=1, value="CATEGORY-WISE PERFORMANCE SUMMARY").font = Font(name="Segoe UI", size=11, bold=True, color=brand_dark)
    
    cat_headers = ["Category", "Orders", "Units Sold", "Revenue (INR)", "Profit (INR)", "Profit Margin %"]
    for c_idx, h in enumerate(cat_headers, start=1):
        c = ws_dash.cell(row=cat_start_row + 1, column=c_idx, value=h)
        c.fill = PatternFill(start_color=brand_primary, end_color=brand_primary, fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    cat_df = analysis_results["category_summary"]
    for r_idx, row in enumerate(cat_df.itertuples(), start=cat_start_row + 2):
        ws_dash.cell(row=r_idx, column=1, value=row.Category).font = bold_cell_font
        ws_dash.cell(row=r_idx, column=2, value=row.Orders).font = cell_font
        ws_dash.cell(row=r_idx, column=3, value=row.Quantity_Sold).font = cell_font
        
        c_rev = ws_dash.cell(row=r_idx, column=4, value=row.Revenue)
        c_rev.font = cell_font
        c_rev.number_format = "#,##0.00"

        c_prof = ws_dash.cell(row=r_idx, column=5, value=row.Profit)
        c_prof.font = cell_font
        c_prof.number_format = "#,##0.00"

        c_mrg = ws_dash.cell(row=r_idx, column=6, value=row.Profit_Margin / 100.0)
        c_mrg.font = cell_font
        c_mrg.number_format = "0.00%"

        for c_idx in range(1, 7):
            ws_dash.cell(row=r_idx, column=c_idx).border = border_thin

    # Section 2: Regional Performance Table on Dashboard (Row 16)
    reg_start_row = 16
    ws_dash.cell(row=reg_start_row, column=1, value="REGIONAL PERFORMANCE SUMMARY").font = Font(name="Segoe UI", size=11, bold=True, color=brand_dark)
    reg_headers = ["Region", "Orders", "Units Sold", "Revenue (INR)", "Profit (INR)", "Profit Margin %"]
    for c_idx, h in enumerate(reg_headers, start=1):
        c = ws_dash.cell(row=reg_start_row + 1, column=c_idx, value=h)
        c.fill = PatternFill(start_color=brand_primary, end_color=brand_primary, fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    reg_df = analysis_results["region_summary"]
    for r_idx, row in enumerate(reg_df.itertuples(), start=reg_start_row + 2):
        ws_dash.cell(row=r_idx, column=1, value=row.Region).font = bold_cell_font
        ws_dash.cell(row=r_idx, column=2, value=row.Orders).font = cell_font
        ws_dash.cell(row=r_idx, column=3, value=row.Quantity_Sold).font = cell_font
        
        c_rev = ws_dash.cell(row=r_idx, column=4, value=row.Revenue)
        c_rev.font = cell_font
        c_rev.number_format = "#,##0.00"

        c_prof = ws_dash.cell(row=r_idx, column=5, value=row.Profit)
        c_prof.font = cell_font
        c_prof.number_format = "#,##0.00"

        c_mrg = ws_dash.cell(row=r_idx, column=6, value=row.Profit_Margin / 100.0)
        c_mrg.font = cell_font
        c_mrg.number_format = "0.00%"

        for c_idx in range(1, 7):
            ws_dash.cell(row=r_idx, column=c_idx).border = border_thin

    # Section 3: Professional Native Charts embedded onto the Dashboard
    # Chart A: Category Revenue & Profit Bar Chart (Placed at H8)
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Category Revenue & Net Profit"
    chart1.y_axis.title = "Amount (INR)"
    chart1.x_axis.title = "Category"
    chart1.height = 10
    chart1.width = 16

    data1 = Reference(ws_dash, min_col=4, min_row=cat_start_row + 1, max_col=5, max_row=cat_start_row + 1 + len(cat_df))
    cats1 = Reference(ws_dash, min_col=1, min_row=cat_start_row + 2, max_row=cat_start_row + 1 + len(cat_df))
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    ws_dash.add_chart(chart1, "H8")

    # Chart B: Regional Profit Comparison Bar Chart (Placed at H22)
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 11
    chart2.title = "Regional Sales Revenue"
    chart2.y_axis.title = "Revenue (INR)"
    chart2.x_axis.title = "Region"
    chart2.height = 10
    chart2.width = 16

    data2 = Reference(ws_dash, min_col=4, min_row=reg_start_row + 1, max_row=reg_start_row + 1 + len(reg_df))
    cats2 = Reference(ws_dash, min_col=1, min_row=reg_start_row + 2, max_row=reg_start_row + 1 + len(reg_df))
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    ws_dash.add_chart(chart2, "H22")

    # Instructions box for Excel Slicers & PivotTables
    inst_row = 26
    ws_dash.cell(row=inst_row, column=1, value="HOW TO INTERACT & ADD EXCEL PIVOTTABLES / SLICERS:").font = Font(name="Segoe UI", size=10, bold=True, color=brand_dark)
    instructions = [
        "1. Interactive Slicers: Open this file in Microsoft Excel -> Go to 'Cleaned_Data' sheet -> Select any cell in the table.",
        "2. Click 'Insert' ribbon -> Click 'PivotTable' -> Select New or Existing Worksheet.",
        "3. Drag 'Category' or 'Region' to Rows and 'Revenue' & 'Profit' to Values.",
        "4. Click 'Insert Slicer' in the PivotTable Analyze menu and check 'Category', 'Region', and 'Month_Name' to filter dynamically."
    ]
    for i_offset, text in enumerate(instructions, start=1):
        ws_dash.cell(row=inst_row + i_offset, column=1, value=text).font = Font(name="Segoe UI", size=8.5, italic=True, color="444444")

    # -------------------------------------------------------------------------
    # Sheet 2: Raw_Data
    # -------------------------------------------------------------------------
    ws_raw = wb.create_sheet(title="Raw_Data")
    ws_raw.views.sheetView[0].showGridLines = True
    raw_headers = list(raw_df.columns)
    ws_raw.append(raw_headers)
    for c_idx in range(1, len(raw_headers) + 1):
        c = ws_raw.cell(row=1, column=c_idx)
        c.fill = PatternFill(start_color="555555", end_color="555555", fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center")

    for _, row in raw_df.iterrows():
        ws_raw.append([row[col] if pd.notna(row[col]) else "" for col in raw_headers])

    for row in ws_raw.iter_rows(min_row=2, max_row=len(raw_df) + 1, min_col=1, max_col=len(raw_headers)):
        for cell in row:
            cell.font = cell_font
            cell.border = border_thin

    # -------------------------------------------------------------------------
    # Sheet 3: Cleaned_Data
    # -------------------------------------------------------------------------
    ws_clean = wb.create_sheet(title="Cleaned_Data")
    ws_clean.views.sheetView[0].showGridLines = True
    clean_cols = ["Order_ID", "Order_Date", "Product_ID", "Product_Name", "Category", "Region", "Quantity", "Selling_Price", "Cost_Price", "Revenue", "Total_Cost", "Profit", "Profit_Margin"]
    ws_clean.append(clean_cols)
    for c_idx in range(1, len(clean_cols) + 1):
        c = ws_clean.cell(row=1, column=c_idx)
        c.fill = PatternFill(start_color=brand_primary, end_color=brand_primary, fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center")

    # Append rows with actual Excel formulas for Revenue, Total_Cost, Profit, Margin
    for r_idx, row in enumerate(clean_df.itertuples(), start=2):
        ws_clean.cell(row=r_idx, column=1, value=row.Order_ID)
        ws_clean.cell(row=r_idx, column=2, value=row.Order_Date.strftime("%Y-%m-%d"))
        ws_clean.cell(row=r_idx, column=3, value=row.Product_ID)
        ws_clean.cell(row=r_idx, column=4, value=row.Product_Name)
        ws_clean.cell(row=r_idx, column=5, value=row.Category)
        ws_clean.cell(row=r_idx, column=6, value=row.Region)
        ws_clean.cell(row=r_idx, column=7, value=int(row.Quantity))
        ws_clean.cell(row=r_idx, column=8, value=float(row.Selling_Price))
        ws_clean.cell(row=r_idx, column=9, value=float(row.Cost_Price))
        
        # Formulas
        ws_clean.cell(row=r_idx, column=10, value=f"=G{r_idx}*H{r_idx}")   # Revenue
        ws_clean.cell(row=r_idx, column=11, value=f"=G{r_idx}*I{r_idx}")   # Cost
        ws_clean.cell(row=r_idx, column=12, value=f"=J{r_idx}-K{r_idx}")   # Profit
        ws_clean.cell(row=r_idx, column=13, value=f"=(L{r_idx}/J{r_idx})") # Margin %

        for c_idx in range(1, 14):
            c = ws_clean.cell(row=r_idx, column=c_idx)
            c.font = cell_font
            c.border = border_thin
            if c_idx in [8, 9, 10, 11, 12]:
                c.number_format = "#,##0.00"
            elif c_idx == 13:
                c.number_format = "0.00%"

    # -------------------------------------------------------------------------
    # Sheet 4: Product_Analysis
    # -------------------------------------------------------------------------
    ws_pa = wb.create_sheet(title="Product_Analysis")
    ws_pa.views.sheetView[0].showGridLines = True
    prod_headers = ["Product ID", "Product Name", "Category", "Orders", "Quantity Sold", "Revenue", "Total Cost", "Profit", "Profit Margin"]
    ws_pa.append(prod_headers)
    for c_idx in range(1, len(prod_headers) + 1):
        c = ws_pa.cell(row=1, column=c_idx)
        c.fill = PatternFill(start_color=brand_dark, end_color=brand_dark, fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center")

    prod_summary = analysis_results["product_summary"]
    for _, row in prod_summary.iterrows():
        ws_pa.append([
            row["Product_ID"], row["Product_Name"], row["Category"],
            int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0
        ])
    for row in ws_pa.iter_rows(min_row=2, max_row=len(prod_summary) + 1, min_col=1, max_col=len(prod_headers)):
        for cell in row:
            cell.font = cell_font
            cell.border = border_thin
            if cell.column in [6, 7, 8]:
                cell.number_format = "#,##0.00"
            elif cell.column == 9:
                cell.number_format = "0.00%"

    # -------------------------------------------------------------------------
    # Sheet 5: Category_Analysis
    # -------------------------------------------------------------------------
    ws_ca = wb.create_sheet(title="Category_Analysis")
    ws_ca.views.sheetView[0].showGridLines = True
    ws_ca.append(cat_headers)
    for c_idx in range(1, len(cat_headers) + 1):
        c = ws_ca.cell(row=1, column=c_idx)
        c.fill = PatternFill(start_color=brand_dark, end_color=brand_dark, fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center")

    for _, row in cat_df.iterrows():
        ws_ca.append([
            row["Category"], int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0
        ])
    for row in ws_ca.iter_rows(min_row=2, max_row=len(cat_df) + 1, min_col=1, max_col=len(cat_headers)):
        for cell in row:
            cell.font = cell_font
            cell.border = border_thin
            if cell.column in [4, 5, 6]:
                cell.number_format = "#,##0.00"
            elif cell.column == 6:
                cell.number_format = "0.00%"

    # -------------------------------------------------------------------------
    # Sheet 6: Monthly_Analysis
    # -------------------------------------------------------------------------
    ws_ma = wb.create_sheet(title="Monthly_Analysis")
    ws_ma.views.sheetView[0].showGridLines = True
    mon_headers = ["Year-Month", "Month", "Orders", "Units Sold", "Revenue", "Total Cost", "Profit", "Profit Margin"]
    ws_ma.append(mon_headers)
    for c_idx in range(1, len(mon_headers) + 1):
        c = ws_ma.cell(row=1, column=c_idx)
        c.fill = PatternFill(start_color=brand_dark, end_color=brand_dark, fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center")

    mon_df = analysis_results["monthly_summary"]
    for _, row in mon_df.iterrows():
        ws_ma.append([
            row["Year_Month"], row["Month_Name"], int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0
        ])
    for row in ws_ma.iter_rows(min_row=2, max_row=len(mon_df) + 1, min_col=1, max_col=len(mon_headers)):
        for cell in row:
            cell.font = cell_font
            cell.border = border_thin
            if cell.column in [5, 6, 7]:
                cell.number_format = "#,##0.00"
            elif cell.column == 8:
                cell.number_format = "0.00%"

    # -------------------------------------------------------------------------
    # Sheet 7: Region_Analysis
    # -------------------------------------------------------------------------
    ws_ra = wb.create_sheet(title="Region_Analysis")
    ws_ra.views.sheetView[0].showGridLines = True
    ws_ra.append(reg_headers)
    for c_idx in range(1, len(reg_headers) + 1):
        c = ws_ra.cell(row=1, column=c_idx)
        c.fill = PatternFill(start_color=brand_dark, end_color=brand_dark, fill_type="solid")
        c.font = table_head_font
        c.alignment = Alignment(horizontal="center")

    for _, row in reg_df.iterrows():
        ws_ra.append([
            row["Region"], int(row["Orders"]), int(row["Quantity_Sold"]),
            float(row["Revenue"]), float(row["Total_Cost"]), float(row["Profit"]),
            float(row["Profit_Margin"]) / 100.0
        ])
    for row in ws_ra.iter_rows(min_row=2, max_row=len(reg_df) + 1, min_col=1, max_col=len(reg_headers)):
        for cell in row:
            cell.font = cell_font
            cell.border = border_thin
            if cell.column in [4, 5, 6]:
                cell.number_format = "#,##0.00"
            elif cell.column == 6:
                cell.number_format = "0.00%"

    # Auto-fit columns
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    # Save Dashboard
    wb.save(target_path)
    print(f"[+] Saved Interactive Excel Dashboard: {target_path}")

# --------------------------------------------------------------------------------------
# 8. COMPREHENSIVE PROJECT REPORT (.DOCX GENERATION)
# --------------------------------------------------------------------------------------
def generate_project_report(analysis_results, cleaning_audit, chart_paths):
    print("=" * 70)
    print("STEP 7: COMPILING FORMAL ACADEMIC PROJECT REPORT (Project_Report.docx)")
    print("=" * 70)

    target_path = os.path.join(REPORT_DIR, "Project_Report.docx")
    doc = docx.Document()

    # Page Margins: Standard 1 inch
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Helper styling functions
    def set_cell_background(cell, fill_hex):
        tcPr = cell._element.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_hex)
        tcPr.append(shd)

    def add_custom_heading(text, level=1):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        if level == 1:
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(6)
        elif level == 2:
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0x2A, 0x5C, 0x9A)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
        elif level == 3:
            run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
        return p

    def add_body_p(text, bold_prefix=None, space_after=5):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(space_after)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.name = "Calibri"
            r_pre.font.size = Pt(11)
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        return p

    def add_bullet_p(text, bold_prefix=None):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.name = "Calibri"
            r_pre.font.size = Pt(11)
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        return p

    kpis = analysis_results["kpis"]
    hl = analysis_results["highlights"]

    # -------------------------------------------------------------------------
    # 1. COVER PAGE
    # -------------------------------------------------------------------------
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(60)
    p_title.paragraph_format.space_after = Pt(10)
    r_title = p_title.add_run("PRODUCT SALES AND PROFIT ANALYSIS\nUSING PYTHON AND EXCEL")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(40)
    r_sub = p_sub.add_run("A Mini Project Submitted in Partial Fulfillment of the Requirements for\nBachelor of Science in Data Science (B.Sc. Data Science)")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    p_cand = doc.add_paragraph()
    p_cand.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cand.paragraph_format.space_after = Pt(40)
    r_by = p_cand.add_run("Submitted By:\n")
    r_by.font.size = Pt(11)
    r_name = p_cand.add_run("[STUDENT NAME]\n")
    r_name.bold = True
    r_name.font.size = Pt(13)
    r_roll = p_cand.add_run("Roll No / Seat No: [ROLL NUMBER]\nClass: B.Sc. Data Science\nAcademic Year: [2024 - 2025]")
    r_roll.font.size = Pt(11)

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(40)
    r_guide = p_inst.add_run("Under the Guidance of:\n")
    r_guide.font.size = Pt(11)
    r_gname = p_inst.add_run("[PROJECT GUIDE / PROFESSOR NAME]\n\n")
    r_gname.bold = True
    r_gname.font.size = Pt(12)
    r_dept = p_inst.add_run("Department of Data Science and Computer Applications\n[COLLEGE / INSTITUTION NAME]\n[AFFILIATED UNIVERSITY NAME]")
    r_dept.font.size = Pt(11)

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 2. CERTIFICATE TEMPLATE
    # -------------------------------------------------------------------------
    add_custom_heading("CERTIFICATE", level=1)
    p_cert = doc.add_paragraph()
    p_cert.paragraph_format.space_before = Pt(15)
    p_cert.paragraph_format.line_spacing = 1.3
    r_cert = p_cert.add_run(
        "This is to certify that the mini project entitled \"Product Sales and Profit Analysis Using Python and Excel\" is a bonafide work carried out by [STUDENT NAME] (Roll No: [ROLL NUMBER]), student of B.Sc. Data Science at [COLLEGE / INSTITUTION NAME], in partial fulfillment of the requirements for the degree of Bachelor of Science in Data Science during the academic year [2024 - 2025].\n\n"
        "The project has been completed under our supervision and guidance, and the results presented in this report have not been submitted to any other university or institute for the award of any degree or diploma."
    )
    r_cert.font.name = "Calibri"
    r_cert.font.size = Pt(11)

    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(80)
    r_sign = p_sign.add_run(
        "________________________                     ________________________\n"
        "   [Project Guide]                                  [Head of Department]\n"
        "Department of Data Science                      Department of Data Science\n\n\n"
        "Date: __________________\n"
        "Place: _________________"
    )
    r_sign.font.name = "Calibri"
    r_sign.font.size = Pt(11)

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 3. ACKNOWLEDGEMENT
    # -------------------------------------------------------------------------
    add_custom_heading("ACKNOWLEDGEMENT", level=1)
    add_body_p(
        "I would like to express my sincere gratitude and heartfelt indebtedness to my project guide, [PROJECT GUIDE NAME], for their continuous encouragement, valuable guidance, constructive critique, and insightful suggestions throughout the development of this mini project."
    )
    add_body_p(
        "I also extend my deep regards to [HEAD OF DEPARTMENT NAME], Head of the Department of Data Science, and the Principal of [COLLEGE / INSTITUTION NAME] for providing the essential infrastructure, computational laboratories, and academic support."
    )
    add_body_p(
        "Finally, I express my appreciation to my family, peers, and faculty members whose motivation and unwavering encouragement helped me successfully conclude this project."
    )
    p_ack_sign = doc.add_paragraph()
    p_ack_sign.paragraph_format.space_before = Pt(30)
    r_as = p_ack_sign.add_run("[STUDENT NAME]\nB.Sc. Data Science")
    r_as.bold = True

    # -------------------------------------------------------------------------
    # 4. ABSTRACT
    # -------------------------------------------------------------------------
    add_custom_heading("ABSTRACT", level=1)
    add_body_p(
        "In modern commercial enterprises, transaction datasets accumulate rapidly. However, raw transactional records do not directly yield actionable managerial insights without systematic extraction, computational processing, and visual modeling. This mini project, titled 'Product Sales and Profit Analysis Using Python and Excel', bridges programmatic computational data science with intuitive business intelligence spreadsheet workflows."
    )
    add_body_p(
        "Using a reproducible 2024 sales transaction dataset, the project utilizes Python (Pandas, NumPy, Matplotlib) to perform rigorous data hygiene, automated missing-value and duplicate resolution, boundary validation, and derived financial indicator calculations (Revenue, Total Cost of Goods, Net Profit, and Operating Profit Margin). Aggregations are conducted across products, categories, geographical territories, and calendar months. In parallel, Microsoft Excel is integrated to engineer an executive dashboard featuring dynamic KPI metric cards, grouped summary pivot tables, native visualization charts, and conditional formatting. Key outcomes include pinpointing highest-volume products, peak profit generators, low-margin products, and seasonal demand trajectories, demonstrating how open-source data science combined with Excel delivers actionable business intelligence."
    )

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 5. INTRODUCTION
    # -------------------------------------------------------------------------
    add_custom_heading("1. INTRODUCTION", level=1)
    add_body_p(
        "Sales analysis is an indispensable operational and strategic practice in contemporary commerce. Retailers, manufacturers, and e-commerce platforms handle thousands of discrete transactions across numerous distribution channels. Without systematic data analysis, organizations face severe challenges such as blind inventory procurement, ineffective discounting schemes, suboptimal warehouse allocation, and compressed profit margins."
    )
    add_body_p(
        "Data science provides structured methodologies to clean, transform, aggregate, and visualize raw transactional logs. Python has emerged as the premier language for statistical data science due to its rich ecosystem of open-source libraries like Pandas and NumPy. Concurrently, Microsoft Excel remains the ubiquitous standard in commercial executive decision-making. By uniting Python's robust analytical capabilities with Excel's accessible spreadsheet modeling, this project provides a dual-tool data analysis architecture tailored for real-world enterprise reporting."
    )

    # -------------------------------------------------------------------------
    # 6. PROBLEM STATEMENT
    # -------------------------------------------------------------------------
    add_custom_heading("2. PROBLEM STATEMENT", level=1)
    add_body_p(
        "Modern business enterprises generate voluminous transactional records across disparate product categories and territories. However, raw sales tables in their uncleaned form present several critical hurdles:"
    )
    add_bullet_p("Raw data contains noise, missing attributes, duplicated records, and invalid values that distort business conclusions.", "Data Quality Obstacles: ")
    add_bullet_p("Manual recalculations of total costs, revenues, and profit margins are prone to human calculation error and lack scalability.", "Computational Bottlenecks: ")
    add_bullet_p("Without hierarchical grouping, management cannot distinguish between products generating high revenue versus products delivering high net profit margin.", "Lack of Granular Profit Visibility: ")
    add_bullet_p("Absence of monthly longitudinal and regional territory tracking prevents managers from detecting seasonal trends or geographical underperformance.", "Strategic Blindspots: ")
    add_body_p(
        "Therefore, there is an urgent need to build an automated, reliable, and auditable pipeline that ingests raw sales data, cleans and validates records, computes exact financial indicators, and synthesizes visual executive dashboards."
    )

    # -------------------------------------------------------------------------
    # 7. OBJECTIVES
    # -------------------------------------------------------------------------
    add_custom_heading("3. OBJECTIVES", level=1)
    add_body_p("The primary objectives of this B.Sc. Data Science mini project are:")
    add_bullet_p("Synthesize a realistic, reproducible multi-category sales dataset spanning 12 calendar months with controlled anomalies.", "Dataset Curation: ")
    add_bullet_p("Implement an automated data cleaning script in Python to identify nulls, duplicates, and boundary violations without silently dropping valid records.", "Data Cleaning & Hygiene: ")
    add_bullet_p("Accurately calculate Revenue, Total Cost, Net Profit, and Profit Margin with robust mathematical zero-division protection.", "Financial Computation: ")
    add_bullet_p("Analyze sales performance across product lines, product categories, geographic regions, and monthly time horizons.", "Multi-Dimensional Aggregation: ")
    add_bullet_p("Identify high-performing and low-performing products based on explicit business criteria.", "Benchmarking: ")
    add_bullet_p("Generate publication-ready data visualizations using Matplotlib.", "Statistical Visualization: ")
    add_bullet_p("Design a comprehensive, professional Excel dashboard featuring KPI scorecards, summary tables, native charts, and structured tabs.", "Executive Dashboarding: ")

    # -------------------------------------------------------------------------
    # 8. TOOLS AND TECHNOLOGIES
    # -------------------------------------------------------------------------
    add_custom_heading("4. TOOLS AND TECHNOLOGIES", level=1)
    
    tools_table = doc.add_table(rows=1, cols=3)
    tools_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = tools_table.rows[0].cells
    hdr_cells[0].text = "Technology / Tool"
    hdr_cells[1].text = "Role in Project"
    hdr_cells[2].text = "Primary Library / Features Used"
    for c in hdr_cells:
        set_cell_background(c, "1F4E79")
        for p in c.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.bold = True
                r.font.size = Pt(10)

    tech_list = [
        ("Python 3.12", "Core programming environment", "File handling, functional orchestration, automation"),
        ("Pandas", "Tabular data manipulation", "DataFrames, grouping, aggregation, sorting, CSV ingestion"),
        ("NumPy", "Vectorized numerical computing", "Conditional handling (np.where), array math, random seeding"),
        ("Matplotlib", "Statistical data visualization", "Horizontal bars, grouped charts, line trends, formatting"),
        ("Seaborn", "Statistical aesthetic theming", "Plot formatting, color styling, grid aesthetics"),
        ("Openpyxl", "Spreadsheet engine", "Workbook creation, cell formulas, formatting, native Excel charts"),
        ("Microsoft Excel", "Business dashboard & reporting", "Dynamic KPI cards, PivotTables, slicers, conditional styling"),
        ("Python-Docx", "Automated report compilation", "Document generation, academic layout, image embedding")
    ]
    for tech, role, feat in tech_list:
        row_cells = tools_table.add_row().cells
        row_cells[0].text = tech
        row_cells[1].text = role
        row_cells[2].text = feat
        for c in row_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------------------
    # 9. DATASET DESCRIPTION
    # -------------------------------------------------------------------------
    add_custom_heading("5. DATASET DESCRIPTION", level=1)
    add_body_p(
        "To ensure complete academic reproducibility and eliminate reliance on proprietary business secrets, a synthetic sales dataset was generated using Python with a fixed random seed of 42. The dataset encompasses 12 calendar months (January 2024 through December 2024), simulating an omni-channel retail distributor."
    )
    add_body_p(
        "Disclaimer: This dataset is entirely synthetic and created solely for academic and educational demonstration within this B.Sc. Data Science mini project.",
        bold_prefix="Academic Dataset Disclaimer: "
    )
    
    ds_table = doc.add_table(rows=1, cols=4)
    ds_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    dhdr = ds_table.rows[0].cells
    dhdr[0].text = "Field Name"
    dhdr[1].text = "Data Type"
    dhdr[2].text = "Sample Value"
    dhdr[3].text = "Business Description"
    for c in dhdr:
        set_cell_background(c, "1F4E79")
        for p in c.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.bold = True
                r.font.size = Pt(10)

    schema = [
        ("Order_ID", "String / Alphanumeric", "ORD-2024012", "Unique alphanumeric transaction identifier"),
        ("Order_Date", "Date (YYYY-MM-DD)", "2024-03-15", "Transaction timestamp spanning 2024"),
        ("Product_ID", "String", "P101", "Unique SKU / product identifier"),
        ("Product_Name", "String", "Laptop Pro 15", "Descriptive commercial product label"),
        ("Category", "String / Categorical", "Electronics", "Broad inventory classification"),
        ("Region", "String / Categorical", "North", "Geographical sales territory"),
        ("Quantity", "Integer", "3", "Number of units purchased per order"),
        ("Selling_Price", "Float / Currency", "55000.00", "Gross unit selling price to customer (INR)"),
        ("Cost_Price", "Float / Currency", "45000.00", "Unit procurement / manufacturing cost (INR)")
    ]
    for fld, dtype, samp, desc in schema:
        row_cells = ds_table.add_row().cells
        row_cells[0].text = fld
        row_cells[1].text = dtype
        row_cells[2].text = samp
        row_cells[3].text = desc
        for c in row_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------------------
    # 10. METHODOLOGY
    # -------------------------------------------------------------------------
    add_custom_heading("6. METHODOLOGY", level=1)
    add_body_p(
        "The project follows a rigorous 6-stage Data Science methodology structured to mirror real-world industrial data pipelines:"
    )
    add_bullet_p("Synthesizing raw transaction logs with deliberate dirty entries to evaluate pipeline resilience.", "Phase 1 - Data Ingestion: ")
    add_bullet_p("Executing programmatic cleaning rules in Pandas to detect null values, duplicates, and non-positive prices/quantities.", "Phase 2 - Data Hygiene: ")
    add_bullet_p("Computing Revenue, Cost of Goods Sold, Profit, and Profit Margin vectors using vectorized arithmetic.", "Phase 3 - Mathematical Transformation: ")
    add_bullet_p("Aggregating multi-dimensional metrics across SKUs, categories, regions, and time series.", "Phase 4 - Exploratory Aggregation: ")
    add_bullet_p("Exporting high-resolution visualization charts with data labels and unified palettes.", "Phase 5 - Visual Storytelling: ")
    add_bullet_p("Generating structured Excel workbooks with formula references and an interactive Dashboard sheet.", "Phase 6 - Executive Dashboard Delivery: ")

    # -------------------------------------------------------------------------
    # 11. DATA CLEANING & AUDIT
    # -------------------------------------------------------------------------
    add_custom_heading("7. DATA CLEANING AND QUALITY AUDIT", level=1)
    add_body_p(
        "A hallmark of professional data science is ensuring data integrity prior to analytical inference. Raw transactional systems frequently harbor malformed inputs caused by input slips, network dropouts, or database glitches. To test our data cleaning engine, intentional anomalies were introduced into the raw dataset."
    )
    
    cl_table = doc.add_table(rows=1, cols=3)
    cl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    chdr = cl_table.rows[0].cells
    chdr[0].text = "Cleaning Verification Step"
    chdr[1].text = "Observed Condition in Raw Data"
    chdr[2].text = "Action Taken & Records Affected"
    for c in chdr:
        set_cell_background(c, "1F4E79")
        for p in c.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.bold = True
                r.font.size = Pt(10)

    clean_audit_rows = [
        ("Duplicate Identification", f"{cleaning_audit['duplicates_removed']} duplicate records detected", "Removed duplicate rows using drop_duplicates()"),
        ("Missing Categorical Check", f"{cleaning_audit['missing_region']} record missing 'Region'", "Identified NaN value; filtered out record"),
        ("Missing Numeric Check", "1 record missing 'Quantity'", "Identified NaN value; filtered out record"),
        ("Non-Positive Quantity Check", f"{cleaning_audit['invalid_quantity']} record with Quantity <= 0", "Validated Quantity > 0; removed erroneous record"),
        ("Zero / Negative Price Check", f"{cleaning_audit['invalid_price']} record with Price <= 0", "Validated Selling_Price > 0; removed entry"),
        ("Date Format Standardization", "Mixed ISO strings", "Converted via pd.to_datetime(), sorted chronologically")
    ]
    for step, obs, act in clean_audit_rows:
        row_cells = cl_table.add_row().cells
        row_cells[0].text = step
        row_cells[1].text = obs
        row_cells[2].text = act
        for c in row_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    add_body_p(
        f"Summary Audit Result: The raw dataset comprised {cleaning_audit['total_raw']} records. After applying strict validation rules, {cleaning_audit['total_clean']} high-integrity records were retained ({cleaning_audit['records_retained_pct']}% retention rate). No valid commercial transactions were silently dropped.",
        bold_prefix="Audit Summary: "
    )

    # -------------------------------------------------------------------------
    # 12. FORMULAS AND CALCULATIONS
    # -------------------------------------------------------------------------
    add_custom_heading("8. FORMULAS AND MATHEMATICAL FOUNDATION", level=1)
    add_body_p("The project implements standard commercial accounting formulas across both Python and Excel:")
    
    add_body_p("Revenue represents the total monetary inflow generated from merchandise sales:", bold_prefix="1. Sales Revenue: ")
    add_body_p("        Revenue = Quantity × Selling_Price\n        Excel Formula: =G2*H2")

    add_body_p("Total Cost (Cost of Goods Sold - COGS) reflects direct expenditure incurred to acquire inventory:", bold_prefix="2. Total Cost: ")
    add_body_p("        Total Cost = Quantity × Cost_Price\n        Excel Formula: =G2*I2")

    add_body_p("Net Profit signifies the surplus earnings remaining after deducting procurement expenses from revenue:", bold_prefix="3. Net Profit: ")
    add_body_p("        Profit = Revenue − Total_Cost\n        Excel Formula: =J2-K2")

    add_body_p("Profit Margin indicates operational profitability as a percentage of top-line revenue:", bold_prefix="4. Profit Margin (%): ")
    add_body_p("        Profit Margin (%) = (Profit / Revenue) × 100\n        Excel Formula: =L2/J2 (formatted as %)")
    add_body_p(
        "Zero-Division Protection: In Python, edge cases where Revenue = 0 are safely guarded using np.where(Revenue > 0, (Profit / Revenue) * 100.0, 0.0), avoiding ZeroDivisionError or NaN propagation.",
        bold_prefix="Robust Zero Handling: "
    )

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 13. PYTHON IMPLEMENTATION
    # -------------------------------------------------------------------------
    add_custom_heading("9. PYTHON IMPLEMENTATION", level=1)
    add_body_p(
        "The Python codebase is designed according to clean software engineering practices with modular separation between data ingestion, cleaning, business calculation, statistical aggregation, chart visualization, and automated artifact generation."
    )
    add_bullet_p("main.py: Automated standalone Python execution script orchestrating the end-to-end pipeline.", "Primary Script: ")
    add_bullet_p("analysis.ipynb: Interactive Jupyter Notebook providing exploratory data analysis, step-by-step DataFrame inspections, and inline Matplotlib charts.", "Interactive Notebook: ")
    add_bullet_p("Openpyxl Engine: Programmatically builds Excel workbooks with custom typography, fill colors, thin cell borders, number formatting, and embedded Excel charts.", "Spreadsheet Automation: ")
    add_bullet_p("Python-Docx Engine: Compiles this formal academic project report with embedded charts and formatted data tables.", "Document Compiler: ")

    # -------------------------------------------------------------------------
    # 14. EXCEL DASHBOARD DESIGN
    # -------------------------------------------------------------------------
    add_custom_heading("10. EXCEL DASHBOARD ARCHITECTURE", level=1)
    add_body_p(
        "The workbook 'Product_Sales_Dashboard.xlsx' is engineered into 7 organized, dedicated worksheets:"
    )
    add_bullet_p("Contains 6 prominent executive KPI cards, Category and Regional performance summary tables, two native Excel charts, and step-by-step PivotTable/Slicer instructions.", "1. Dashboard: ")
    add_bullet_p("Unmodified initial dataset including deliberate anomalies for auditability.", "2. Raw_Data: ")
    add_bullet_p("The 300 validated transactions with live Excel formulas for Revenue, Cost, Profit, and Profit Margin.", "3. Cleaned_Data: ")
    add_bullet_p("SKU-level performance table detailing units sold, revenue, cost, profit, and profit margin.", "4. Product_Analysis: ")
    add_bullet_p("Aggregated category-level summary with revenue share percentages.", "5. Category_Analysis: ")
    add_bullet_p("Chronological monthly sales progression over the 12 calendar months.", "6. Monthly_Analysis: ")
    add_bullet_p("Territory-level performance summary comparing East, West, North, South, and Central zones.", "7. Region_Analysis: ")

    # -------------------------------------------------------------------------
    # 15. RESULTS AND FINDINGS
    # -------------------------------------------------------------------------
    add_custom_heading("11. RESULTS AND EMPIRICAL FINDINGS", level=1)
    add_body_p("Execution of the computational analysis yielded the following key performance metrics:")
    add_body_p(
        f"Across the full calendar year 2024, the enterprise recorded Total Sales Revenue of INR {kpis['Total_Revenue']:,.2f} against Total Cost of Goods Sold of INR {kpis['Total_Cost']:,.2f}, realizing a Net Operating Profit of INR {kpis['Total_Profit']:,.2f} with an Overall Profit Margin of {kpis['Overall_Profit_Margin']}%. Total merchandise volume reached {kpis['Total_Quantity_Sold']:,} units across {kpis['Total_Orders']:,} validated customer transactions with an Average Order Value (AOV) of INR {kpis['Avg_Order_Value']:,.2f}."
    )

    # KPI Table
    kpi_tbl = doc.add_table(rows=1, cols=3)
    kpi_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    khdr = kpi_tbl.rows[0].cells
    khdr[0].text = "Executive Key Performance Indicator (KPI)"
    khdr[1].text = "Computed Value"
    khdr[2].text = "Managerial Implication"
    for c in khdr:
        set_cell_background(c, "1F4E79")
        for p in c.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.bold = True
                r.font.size = Pt(10)

    kpi_summary_rows = [
        ("Total Sales Revenue", f"INR {kpis['Total_Revenue']:,.2f}", "Total gross turnover achieved across all products in 2024"),
        ("Total Cost of Goods Sold", f"INR {kpis['Total_Cost']:,.2f}", "Total procurement / manufacturing expenditure incurred"),
        ("Net Operating Profit", f"INR {kpis['Total_Profit']:,.2f}", "Net bottom-line earnings retained by the enterprise"),
        ("Overall Profit Margin", f"{kpis['Overall_Profit_Margin']}%", "Healthy operational return on aggregate turnover"),
        ("Total Units Sold", f"{kpis['Total_Quantity_Sold']:,} units", "Aggregate physical merchandise volume moved"),
        ("Total Processed Orders", f"{kpis['Total_Orders']:,} orders", "Total valid commercial customer transactions"),
        ("Average Order Value (AOV)", f"INR {kpis['Avg_Order_Value']:,.2f}", "Average gross revenue realized per customer order")
    ]
    for kpi_name, kpi_val, kpi_imp in kpi_summary_rows:
        row_cells = kpi_tbl.add_row().cells
        row_cells[0].text = kpi_name
        row_cells[1].text = kpi_val
        row_cells[2].text = kpi_imp
        for c in row_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Performance Highlights
    add_custom_heading("Key Product Highlights", level=2)
    add_bullet_p(f"{hl['best_selling_by_qty']['Product_Name']} with {hl['best_selling_by_qty']['Quantity_Sold']} units sold.", "Best-Selling Product (by Volume): ")
    add_bullet_p(f"{hl['best_selling_by_rev']['Product_Name']} generating INR {hl['best_selling_by_rev']['Revenue']:,.2f}.", "Best-Selling Product (by Revenue): ")
    add_bullet_p(f"{hl['most_profitable']['Product_Name']} delivering net profit of INR {hl['most_profitable']['Profit']:,.2f} ({hl['most_profitable']['Profit_Margin']}% margin).", "Most Profitable Product: ")
    add_bullet_p(f"{hl['lowest_performing']['Product_Name']} generating net profit of INR {hl['lowest_performing']['Profit']:,.2f} ({hl['lowest_performing']['Profit_Margin']}% margin).", "Lowest-Performing Product: ")
    add_body_p(
        "Criteria for Lowest-Performing Product: Evaluated primarily by net profit contribution. While low-priced stationery items (such as Gel Pens) achieve respectable profit margins (~50%), their minute unit price results in minimal absolute profit contribution compared to high-ticket consumer technology items.",
        bold_prefix="Benchmark Criteria Note: "
    )

    doc.add_page_break()

    # Visualizations Insertion
    add_custom_heading("12. GRAPHICAL VISUALIZATIONS AND TREND ANALYSIS", level=1)
    
    # Chart 1
    if "product_profit" in chart_paths and os.path.exists(chart_paths["product_profit"]):
        add_custom_heading("Figure 1: Product-Wise Net Profit Breakdown", level=2)
        doc.add_picture(chart_paths["product_profit"], width=Inches(6.0))
        p_cap = doc.add_paragraph("Figure 1: Net profit realized across individual product catalog items (INR in Thousands).")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].font.size = Pt(9)
        p_cap.runs[0].font.italic = True
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Chart 2
    if "category_sales" in chart_paths and os.path.exists(chart_paths["category_sales"]):
        add_custom_heading("Figure 2: Category-Wise Revenue vs. Profit", level=2)
        doc.add_picture(chart_paths["category_sales"], width=Inches(5.8))
        p_cap = doc.add_paragraph("Figure 2: Comparative performance across commercial product categories (INR in Lakhs).")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].font.size = Pt(9)
        p_cap.runs[0].font.italic = True
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Chart 3
    if "monthly_sales" in chart_paths and os.path.exists(chart_paths["monthly_sales"]):
        add_custom_heading("Figure 3: Monthly Sales, Cost and Profit Trend", level=2)
        doc.add_picture(chart_paths["monthly_sales"], width=Inches(6.0))
        p_cap = doc.add_paragraph("Figure 3: 12-Month longitudinal financial trend illustrating revenue, cost, and profit trajectory.")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].font.size = Pt(9)
        p_cap.runs[0].font.italic = True
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Chart 4
    if "region_sales" in chart_paths and os.path.exists(chart_paths["region_sales"]):
        add_custom_heading("Figure 4: Regional Territory Sales Distribution", level=2)
        doc.add_picture(chart_paths["region_sales"], width=Inches(5.5))
        p_cap = doc.add_paragraph("Figure 4: Geographic distribution of turnover across North, South, East, West, and Central territories.")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].font.size = Pt(9)
        p_cap.runs[0].font.italic = True
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Chart 5
    if "profit_comparison" in chart_paths and os.path.exists(chart_paths["profit_comparison"]):
        add_custom_heading("Figure 5: Overall Financial Component Comparison", level=2)
        doc.add_picture(chart_paths["profit_comparison"], width=Inches(5.5))
        p_cap = doc.add_paragraph("Figure 5: High-level financial breakdown of Revenue, Cost of Goods Sold, and Net Profit.")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].font.size = Pt(9)
        p_cap.runs[0].font.italic = True
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 16. LIMITATIONS
    # -------------------------------------------------------------------------
    add_custom_heading("13. PROJECT LIMITATIONS", level=1)
    add_body_p("While this project successfully demonstrates a rigorous data analysis solution, certain inherent academic limitations should be noted:")
    add_bullet_p("The dataset is synthetically generated with fixed random seeds for academic reproducibility rather than reflecting live multi-tenant ERP databases.", "Synthetic Data Origin: ")
    add_bullet_p("Cost prices and selling prices are treated as relatively static per order without factoring in fluctuating vendor purchase inflation or multi-tiered corporate volume rebates.", "Static Cost Structures: ")
    add_bullet_p("The current pipeline operates on batch CSV exports rather than real-time transactional event streaming (e.g., Apache Kafka).", "Batch Processing Paradigm: ")
    add_bullet_p("The model assumes successful delivery and excludes returned merchandise, customer warranty claims, and shipping freight logistics.", "Excluded Secondary Overhead: ")

    # -------------------------------------------------------------------------
    # 17. FUTURE SCOPE
    # -------------------------------------------------------------------------
    add_custom_heading("14. FUTURE SCOPE", level=1)
    add_body_p("This project forms a robust foundation for numerous advanced data science enhancements:")
    add_bullet_p("Incorporate ARIMA, Prophet, or LSTM machine learning models to forecast monthly demand and automate re-order points.", "Predictive Demand Forecasting: ")
    add_bullet_p("Apply K-Means clustering and RFM (Recency, Frequency, Monetary) modeling to segment customers by lifetime value.", "Customer Segmentation: ")
    add_bullet_p("Migrate the local Excel dashboard into enterprise cloud business intelligence tools such as Power BI or Tableau with automated daily refreshes.", "Cloud BI Integration: ")
    add_bullet_p("Build a lightweight web front-end using Streamlit or FastAPI to allow non-technical store managers to upload CSV files and instantly receive downloadable analysis reports.", "Web Application Deployment: ")

    # -------------------------------------------------------------------------
    # 18. CONCLUSION
    # -------------------------------------------------------------------------
    add_custom_heading("15. CONCLUSION", level=1)
    add_body_p(
        "The mini project 'Product Sales and Profit Analysis Using Python and Excel' successfully achieves its objective of delivering a complete, robust, and beginner-friendly data science solution for commercial performance analysis. By harnessing the programmatic data cleaning, vectorized computation, and visualization power of Python alongside the accessible executive reporting capabilities of Microsoft Excel, the project bridges the gap between raw data manipulation and managerial decision-making."
    )
    add_body_p(
        "The automated pipeline safely resolves dirty records, calculates exact revenue, cost, profit, and profit margin metrics, and extracts meaningful performance highlights. The resulting interactive Excel dashboard and multi-dimensional charts empower stakeholders to quickly pinpoint lucrative inventory, address underperforming territories, and optimize strategic profitability."
    )

    # -------------------------------------------------------------------------
    # 19. REFERENCES
    # -------------------------------------------------------------------------
    add_custom_heading("16. REFERENCES", level=1)
    refs = [
        "McKinney, Wes. Python for Data Analysis: Data Wrangling with Pandas, NumPy, and Jupyter. 3rd ed., O'Reilly Media, 2022.",
        "VanderPlas, Jake. Python Data Science Handbook: Essential Tools for Working with Data. 2nd ed., O'Reilly Media, 2023.",
        "Alexander, Michael, and Dick Kusleika. Excel Dashboards and Reports. 3rd ed., John Wiley & Sons, 2016.",
        "Pandas Development Team. 'Pandas Documentation.' https://pandas.pydata.org/docs/ (Accessed 2024).",
        "Matplotlib Development Team. 'Matplotlib: Visualization with Python.' https://matplotlib.org/stable/ (Accessed 2024).",
        "Openpyxl Community. 'openpyxl - A Python library to read/write Excel 2010 xlsx/xlsm files.' https://openpyxl.readthedocs.io/ (Accessed 2024).",
        "Microsoft Corporation. 'Create a PivotTable to analyze worksheet data.' Microsoft Support Documentation, 2024."
    ]
    for r_text in refs:
        add_bullet_p(r_text)

    # Save Word document
    doc.save(target_path)
    print(f"[+] Saved Formal Academic Project Report: {target_path}")

# --------------------------------------------------------------------------------------
# 9. MAIN PIPELINE EXECUTION
# --------------------------------------------------------------------------------------
def main():
    print("=" * 80)
    print("PRODUCT SALES AND PROFIT ANALYSIS USING PYTHON AND EXCEL")
    print("B.Sc. Data Science Mini Project Pipeline Execution")
    print("=" * 80)

    # 1. Clean Data
    raw_df, clean_df, cleaning_audit = clean_and_prepare_data(CSV_FILE_PATH)

    # 2. Perform Calculations
    calculated_df = perform_calculations(clean_df)

    # 3. Analyze Data
    analysis_results = analyze_data(calculated_df)

    # 4. Generate Visualizations
    chart_paths = generate_visualizations(analysis_results)

    # 5. Generate Sales_Analysis.xlsx
    generate_sales_analysis_workbook(calculated_df, analysis_results)

    # 6. Generate Product_Sales_Dashboard.xlsx
    generate_interactive_dashboard(raw_df, calculated_df, analysis_results)

    # 7. Generate Project_Report.docx
    generate_project_report(analysis_results, cleaning_audit, chart_paths)

    print("=" * 80)
    print("[SUCCESS] ALL PROJECT WORKFLOWS COMPLETED SUCCESSFULLY!")
    print(f"[*] Visualizations saved to:       {OUTPUT_DIR}")
    print(f"[*] Analytical Workbook saved to:  {os.path.join(OUTPUT_DIR, 'Sales_Analysis.xlsx')}")
    print(f"[*] Interactive Dashboard saved to: {os.path.join(DASHBOARD_DIR, 'Product_Sales_Dashboard.xlsx')}")
    print(f"[*] Formal Project Report saved to: {os.path.join(REPORT_DIR, 'Project_Report.docx')}")
    print("=" * 80)

if __name__ == "__main__":
    main()
