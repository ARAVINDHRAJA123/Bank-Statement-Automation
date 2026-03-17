# 💳 Bank Statement Analyzer

A Python-based data engineering project that transforms unstructured bank statement PDFs into structured, analytics-ready Excel reports with automated insights and visual dashboards.

---

## 🚀 Overview

Bank statements are semi-structured documents with inconsistent layouts, multi-line entries, and embedded metadata. Manually analyzing them is time-consuming and error-prone.

This project solves that by building an end-to-end pipeline that:

- Extracts transaction data from PDF statements
- Cleans and reconstructs multi-line records
- Detects merchants automatically
- Categorizes spending patterns
- Generates Excel reports with charts and insights

---

## ⚙️ Features

### 📄 Data Extraction
- Parses multi-page bank statement PDFs
- Handles inconsistent formats across pages
- Reconstructs multi-line transaction records

### 🧠 Data Processing
- Cleans and normalizes transaction data
- Extracts merchant names from narration
- Categorizes spending into logical buckets

### 📊 Analytics & Reporting
- Monthly spending summary
- Category-wise distribution
- Top merchants analysis
- Key financial metrics

### 📈 Excel Dashboard
- Auto-formatted output (date + currency)
- Auto-fit column widths
- Pie chart for spending distribution
- Bar chart for top merchants

---

## 📊 Output Structure

### 1. Transactions Sheet
| Date | Merchant | Narration | Debit | Balance | Category |
|------|----------|----------|------|---------|----------|

---

### 2. Monthly Summary
| Month | Total Spend |

---

### 3. Category Breakdown
- Spending grouped by amount ranges  
- Pie chart visualization  

---

### 4. Top Merchants (NEW)
| Merchant | Total Spend |
- Bar chart of top 10 merchants  

---

### 5. Spending Stats
| Metric | Value |
|--------|------|
| Total Spending |
| Average Transaction |
| Largest Transaction |

---

## 📷 Sample Output

### Transactions & Analytics Dashboard
<img width="662" height="240" alt="image" src="https://github.com/user-attachments/assets/c05a81e9-431b-48ec-b570-8e868f099556" />


## 🧱 Architecture
