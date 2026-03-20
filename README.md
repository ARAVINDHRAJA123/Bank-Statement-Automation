# 💳 Bank Statement Analyser

A Python tool that converts HDFC bank statement PDFs into a structured Excel report — with smart spending categories, anomaly detection, income vs expense charts, and merchant analysis. No API key or external tools needed.

---

## 🚀 Quick Start

```bash
pip install -r requirements.txt
```

Rename your HDFC PDF to `Account Statement.pdf` and place it in the same folder, then:

```bash
python bank_statement_analyser.py
```

Open `Bank_Statement_Report.xlsx` to see your report.

---

## 📦 Requirements

- Python 3.10 or higher
- See `requirements.txt`

```
pdfplumber>=0.11.0
openpyxl>=3.1.0
```

---

## ⚙️ How It Works

```
Account Statement.pdf
        │
        ▼
┌──────────────────────────┐
│  Spatial PDF Extraction  │  Reads every word with its x/y position on the page
│  (pdfplumber)            │  Assigns each word to the correct column by coordinate
│                          │  Reconstructs multi-line narrations automatically
└───────────┬──────────────┘
            │
            ▼
┌──────────────────────────┐
│  Clean & Enrich          │  Removes duplicates
│                          │  Extracts clean merchant names from UPI/POS/NEFT strings
│                          │  Assigns a spending category to each transaction
└───────────┬──────────────┘
            │
            ▼
┌──────────────────────────┐
│  Analytics               │  Monthly income vs expense breakdown
│                          │  Category-wise spending summary
│                          │  Top 10 merchants by spend
│                          │  Anomaly detection — flags unusually large transactions
└───────────┬──────────────┘
            │
            ▼
   Bank_Statement_Report.xlsx
```

---

## 📊 Output — 6 Sheets

### 1. Summary
Total income, total expenses, net cash flow, transaction counts, largest credit, largest debit — all in one place.

### 2. Transactions
Full transaction list with date, merchant, narration, ref number, debit, credit, balance, category, and an anomaly flag. Debits in red, credits in green. Flagged rows highlighted.

| Date | Merchant | Narration | Ref No | Value Date | Debit (₹) | Credit (₹) | Balance (₹) | Category | Flag |
|------|----------|-----------|--------|------------|-----------|------------|-------------|----------|------|

### 3. Monthly Summary
Income, expense, and net per month with a clustered bar chart comparing income vs expense side by side. Value labels shown on every bar.

### 4. Categories
Spending grouped into real categories — Food & Dining, Transport, Shopping, Bills & Utilities, Health, Insurance, Entertainment, Finance & EMI, Salary / Income. Includes a pie chart.

### 5. Top Merchants
Top 10 merchants by total spend with a horizontal bar chart. Value labels on every bar.

### 6. Anomalies ⚠
Transactions that are statistically much larger than your usual spend, with a plain-English explanation of why each one was flagged — *"This is 9.8x your average spend of Rs. 870. Anything above Rs. 2,609 is flagged."*

---

## 🛠 Configuration

All settings are at the top of the script:

```python
INPUT_PDF   = "Account Statement.pdf"   # your PDF filename
OUTPUT_XLSX = "Bank_Statement_Report.xlsx"
ANOMALY_Z   = 2.0   # sensitivity — lower = more flags, higher = fewer
```

To add or edit spending categories, update `CATEGORY_KEYWORDS`:

```python
CATEGORY_KEYWORDS = {
    "Food & Dining": ["swiggy", "zomato", "your_restaurant", ...],
    "My Category":   ["keyword1", "keyword2"],
}
```

---

## 🏦 Compatibility

Tested on HDFC Bank savings account statements (text-based PDF).

> ⚠️ Scanned PDFs will not work. The PDF must be text-based — if you can select and copy text from it, it will work. If not, it is a scanned image and needs OCR first.

For other banks (SBI, ICICI, Axis), the `HDFC_COLS` coordinate boundaries at the top of the script need to be adjusted to match that bank's column layout.

---

## 📁 Project Structure

```
Bank-Statement-Automation/
├── bank_statement_analyser.py   ← main script
├── Account Statement.pdf             ← your input PDF (rename to this)
├── Bank_Statement_Report.xlsx        ← generated output
├── requirements.txt
└── README.md
```

## 📷 Sample Output

### Transactions & Analytics Dashboard

<img width="1339" height="659" alt="image" src="https://github.com/user-attachments/assets/d52e77f7-1d92-43dd-8ab8-a3c9fea39130" />
