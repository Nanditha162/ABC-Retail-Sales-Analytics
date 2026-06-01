# ABC Retail Sales Analytics

## Project Overview

This project implements an end-to-end Data Engineering and Business Intelligence solution for ABC Retail Solutions. The objective is to ingest, clean, transform, validate, and analyze retail transaction data collected from multiple source systems and generate business insights through an interactive Power BI dashboard.

The solution addresses common data quality challenges such as duplicate transactions, missing values, inconsistent product names, invalid quantities, varying date formats, and protection of sensitive customer information through PII masking.

---

## Problem Statement

ABC Retail Solutions operates across multiple cities through both online and offline sales channels. Due to inconsistencies in the transactional datasets, the company faces challenges in generating accurate business insights and reporting.

This project builds a complete data pipeline to:

* Ingest retail transaction data from multiple sources
* Perform data cleaning and transformation
* Apply data quality checks and validations
* Mask sensitive customer information
* Generate business KPIs and aggregated datasets
* Create interactive Power BI dashboards for business reporting

---


## Data Engineering Pipeline

### Data Ingestion

* Read retail_data1.xlsx
* Read retail_data2.xlsx
* Read product_details.xlsx

### Data Cleaning

* Removed duplicate records
* Handled missing values
* Standardized product names
* Standardized categories
* Corrected date formats
* Validated quantities and prices

### PII Masking

* Masked customer email addresses
* Masked customer phone numbers

### Data Enrichment

* Joined transaction data with product dimension table
* Added product category and standardized pricing information

### Data Aggregation

Generated:

* KPI Summary
* Revenue by Category
* Revenue by City
* Revenue by Product
* Revenue by Channel
* Revenue by Payment Method
* Monthly Revenue

---

## Power BI Dashboard

The dashboard consists of four business-focused pages:

### 1. Executive Overview

* Total Revenue
* Total Orders
* Total Customers
* Average Order Value
* Revenue Trend Analysis
* Revenue by Category
* Revenue by City

### 2. Product Performance

* Top Products by Revenue
* Units Sold Analysis
* Product Revenue Contribution
* Product Performance Summary

### 3. Category Trends

* Revenue by Category
* Orders by Category
* Category Contribution Analysis
* Revenue Share by Category

### 4. Regional Insights

* Revenue by City
* Revenue Distribution by Region
* Category Performance Across Cities
* Sales Channel Analysis

---

## Technology Stack

* Python
* Pandas
* NumPy
* Google Colab
* Excel
* Power BI
* GitHub

---

## Repository Structure

```text
ABC-Retail-Sales-Analytics
│
├── Code
│   └── ABC_Retail_Pipeline_code.py
│
├── Data
│   └── curated_retail_data_final.xlsx
│
├── Documentation
│   ├── ABC_Retail_Documentation.docx
│
├── PowerBI
│   ├── ABC_Business_Insights_Dashboard.pbix
│
└── README.md
```

---

## Key Business KPIs

* Total Revenue
* Revenue by Category
* Revenue by City
* Revenue by Product
* Revenue by Payment Method
* Monthly Revenue Trend
* Average Order Value
* Units Sold

---

## Business Outcomes

This solution improves data quality, reporting accuracy, and business decision-making by providing a centralized and analytics-ready retail dataset along with interactive Power BI dashboards.

---

