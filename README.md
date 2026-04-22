# 📊 Excel Analytics Project (Pivot-Based Dashboard)

![Excel](https://img.shields.io/badge/Microsoft%20Excel-Data%20Analysis-green?logo=microsoft-excel)
![Status](https://img.shields.io/badge/status-complete-success)
![Type](https://img.shields.io/badge/type-Analytics-blue)

This repository contains a **deconstructed Microsoft Excel workbook** designed for **data analysis using Pivot Tables, structured sheets, and formatted reporting**.

---

## 📌 Project Overview

This project represents an **Excel-based analytical solution**, including:

📊 Multiple worksheets for data organization
🔄 Pivot tables for dynamic analysis
🎨 Custom styling and themes
⚙️ Workbook-level configurations

---

## 🧱 File Structure (Deconstructed Excel)

```id="excelstruct01"
.
├── 📄 workbook.xml                # Core workbook structure
├── 📄 styles.xml                  # Formatting & styles
├── 📄 sharedStrings.xml           # Text data storage
├── 📄 theme1.xml                  # Theme & design
├── 📄 calcChain.xml               # Calculation dependencies
│
├── 📊 sheet1.xml → sheet5.xml     # Data & report sheets
├── 🔗 sheet*.rels                 # Sheet relationships
│
├── 📊 pivotTable1.xml
├── 📊 pivotTable2.xml             # Pivot table definitions
├── 📦 pivotCacheDefinition*.xml   # Pivot structure
├── 📦 pivotCacheRecords*.xml      # Pivot data cache
│
├── ⚙️ app.xml / core.xml          # Metadata
├── ⚙️ workbook.xml.rels           # Workbook relationships
├── 🖨️ printerSettings*.bin        # Print settings
└── 📄 [Content_Types].xml         # Package definition
```

---

## 📊 Key Features

### 🔹 Multi-Sheet Data Organization

* 5 structured worksheets (`sheet1` → `sheet5`)
* Separation of raw data, processed data, and reports

### 🔹 Pivot Table Analytics

* Pre-built pivot tables:

  * `pivotTable1.xml`
  * `pivotTable2.xml`
* Backed by cache definitions and records
* Enables:

  * Aggregation
  * Filtering
  * Drill-down analysis

### 🔹 Data Processing Engine

* `calcChain.xml` manages formula dependencies
* Ensures efficient recalculation

### 🔹 Styling & UI

* `styles.xml` → formatting rules
* `theme1.xml` → consistent visual design
* `sharedStrings.xml` → optimized text handling

---

## ⚙️ How It Works

```id="workflow01"
Raw Data (Sheets)
      ↓
Data Processing (Formulas + Calc Chain)
      ↓
Pivot Cache (Definitions + Records)
      ↓
Pivot Tables (Analysis Layer)
      ↓
Formatted Output (Styled Sheets)
```

---

## 🚀 How to Use

### 🔹 Option 1: Rebuild Excel File

1. Zip all files into `.zip`
2. Rename to `.xlsx`
3. Open in **Microsoft Excel**

---

### 🔹 Option 2: Analyze Internals

* Open `.xml` files in:

  * VS Code
  * Notepad++
* Inspect:

  * Pivot logic
  * Data structure
  * Relationships

---

## 🛠️ Use Cases

| Use Case               | Description                             |
| ---------------------- | --------------------------------------- |
| 📊 Data Analysis       | Perform aggregations using pivot tables |
| 📈 Reporting           | Generate structured reports             |
| 🔍 Reverse Engineering | Understand Excel internals              |
| ⚙️ Optimization        | Improve formulas and structure          |
| 🎓 Learning            | Study how Excel stores data             |

---

## ⚠️ Important Notes

⚡ This is a **deconstructed Excel file**, not directly editable as-is
⚡ Incorrect modifications may corrupt the workbook
⚡ Always keep a backup before editing

---

## 🤝 Contributing

You can contribute by:

* Improving pivot logic 📊
* Optimizing structure ⚡
* Enhancing formatting 🎨
* Adding documentation 📄


---

## ⭐ Support

If this helped you understand Excel internals or analytics workflows, consider giving it a ⭐!
