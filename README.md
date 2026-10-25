# 🧹 Local Auto Data Cleaner & Power BI-Style Analytics Dashboard

An interactive, user-friendly Web Application built with **Streamlit**, **Pandas**, **Plotly**, and **Scikit-Learn** that transforms messy CSV/Excel datasets into clean, standardized, and business-ready data. It also features a built-in **Power BI-Style Analytics Dashboard** for live visual exploration and automated AI insights.

---

## ✨ Key Features

### 🧱 1. Smart Data Cleaning & Column Operations
* **Column Management:** Quick rename, multi-column drop, and custom delimiter column splitting with real-time previews.
* **Find & Replace & Filtering:** Exact match, substring, or Regex find-and-replace across multiple columns. Combine conditions with custom `AND`/`OR` filtering logic.
* **Text Normalization:** Trim excess whitespace, standardize text cases (*UPPER, lower, Title*), and automatically remove accents, emojis, or special characters.

### 🧼 2. Intelligent Data Formatting & Missing Value Handling
* **Column-First Missing Values Breakdown:** Auto-detects empty cells with missing percentage metrics and allows column-specific actions (*Drop Rows, Mean/Mode Imputation, or Custom Value Filling*).
* **Duplicate Rows Preview & Removal:** Interactive check to preview duplicate rows before dropping them.
* **Header-Specific Formatting:** Convert messy inputs into standardized Currency, Date presets, Custom Date strings, Integers, or Categories.

### 🚀 3. Advanced Transformations
* **Calculated Columns:** Perform mathematical operations between columns or constants (`+`, `-`, `*`, `/`) with custom decimal rounding.
* **Date Duration Calculator:** Calculate elapsed time (*Days, Months, Years, Hours*) between two date columns.
* **Fuzzy Value Grouping:** Powered by `rapidfuzz` / `difflib` to cluster and unify misspelled text entries automatically.
* **Outlier Handling:** Detect outliers using **IQR** or **Z-Score** methods and choose to clip values to boundaries or drop rows.

### 📈 4. Power BI-Style Interactive Canvas
* **Dynamic Slicers:** Filter data live across categories with immediate cross-chart updates.
* **Decomposition Treemap:** Explore categorical hierarchies and metric breakdowns dynamically.
* **🤖 AI Key Influencers:** Uses Machine Learning (*Random Forest*) to identify key factors driving target KPIs.
* **📝 Smart Narrative:** Automatically generates executive text summaries highlighting key statistical insights, peaks, and domain totals.
* **🔢 Quick AutoSum:** Instant KPI metrics for `SUM`, `AVERAGE`, `MIN`, `MAX`, and `COUNT`.

### 📋 5. Data Audit & Export
* **Quality Audit Report:** Before-and-after comparisons of missing values, row/column counts, memory usage, and data types.
* **Multi-Format Export:** Export clean data directly to **CSV** or styled **Excel (.xlsx)** format.
* **💾 Recipe Presets:** Export or import cleaning rules in `.json` format for automated weekly workflows.

---

## 🛠️ Tech Stack & Dependencies

* **Frontend & Framework:** [Streamlit](https://streamlit.io/)
* **Data Processing:** [Pandas](https://pandas.pydata.org/), [NumPy](https://numpy.org/)
* **Data Visualization:** [Plotly Express & Graph Objects](https://plotly.com/python/)
* **Machine Learning:** [Scikit-Learn](https://scikit-learn.org/)
* **Fuzzy Matching:** [RapidFuzz](https://github.com/maxbachmann/RapidFuzz)
* **Excel Engine:** OpenPyXL, XlsxWriter

---

## 🚀 Quick Start & Installation

### 1. Clone the repository
```bash
git clone [https://github.com/fawwazredha/datacleaner.git](https://github.com/fawwazredha/datacleaner.git)
cd datacleaner
