# FY26 Target vs. Achievement Automated Reconciliation Workflow

## Executive Summary
This project provides an automated data processing and reporting pipeline for Sun King expansion product lines (**General Accounts**, **Inverters**, and **PayG Phones**). 

The Python engine ingests raw target and achievement data from `FY27_Targets_and_Achievements.xlsx`, cleans and standardizes geographic mappings, filters out irrelevant metric columns, and exports a standardized, executive-ready Excel workbook: **`FY27 Half AOP Target vs. Achievement.xlsx`**.

---

## Technical Architecture & Requirements

* **Python Version**: Python 3.14+
* **Dependencies**: `pandas`, `openpyxl`, `numpy`
* **Virtual Environment**: `ai-env`
* **IDE**: Visual Studio Code (VS Code)

---

## File Structure

```text
/Code
│
├── FY27_Targets_and_Achievements.xlsx    # Raw Source Workbook
├── FY27 Half AOP Target vs. Achievement.xlsx # Automated Output Deliverable
├── reconcile.py                         # Core Processing Engine
├── README.md                            # Project Documentation
└── ai-env/                              # Virtual Environment
```

---

## Business Logic & Execution Rules

### 1. Ingestion & Column Filtering
* Reads source sheets using header offset configuration (`header_offset=1`).
* Automatically drops non-primary unit metrics such as **Unit Sales**, **WP** (Watt-peak / Wholesale Price), and related columns.
* Filters out summary/variance rows (e.g., `Total`, `Grand Total`, `Subtotal`, `Difference`, `Variance`, `Check`).

### 2. Geographic Standardization & Alias Resolution
Applies automatic remapping across all product datasets prior to merging:
* **Ihala** $\rightarrow$ Re-mapped to `Ihiala`.
* **`*Igabi*`** $\rightarrow$ Re-mapped to `Kaduna North`.
* **Abeokuta North** $\rightarrow$ Reclassified under `Zone = Lagos`, `Region = Ogun West`.
* **Ijebu East** $\rightarrow$ Reclassified under `Zone = Lagos`, `Region = Ogun East`.

### 3. Merging & Status Classification
Performs an **Outer Join** on `Area` to preserve 100% of reported achievements, coalescing missing Zone/Region references and assigning status flags:
* **`Target Achieved`**: Target $> 0$ and Achievement $\ge$ Target. *(Soft Green Fill)*
* **`Target Not Achieved`**: Target $> 0$ and Achievement $<$ Target. *(Soft Red Fill)*
* **`AHQ Not Launched`**: Target $> 0$ and Achievement $= 0$. *(Soft Red Fill)*
* **`Missing from Target Sheet`**: Target $= 0$ and Achievement $> 0$.

---

## Output Workbook Structure

The output file **`FY27 Half AOP Target vs. Achievement.xlsx`** contains three tabs:
1. `Reconciliation Analysis` (General Accounts)
2. `Inverter AOP vs Target` (Inverters)
3. `PayG Phone AOP vs Target` (PayG Phones)

Each tab features a dual-layout structure:
* **Left Table (Cols A–H)**: Main Area-level detail table (`Zone`, `Region`, `Area`, `Target`, `Achievement`, `Deficit`, `% Achieved`, `Status`).
* **Right Table (Cols J–N)**: Zone Summary pivot table at top right, stacked above a Region Summary pivot table.

---

## Quick Start Guide

### Step 1: Open Terminal in VS Code
Open your project folder in VS Code and press `Ctrl + ~` to launch the terminal.

### Step 2: Activate Virtual Environment
```bash
source ai-env/bin/activate
```

### Step 3: Run the Reconciliation Script
```bash
python reconcile.py
```

Upon execution, check the terminal output for summary total validations and open **`FY27 Half AOP Target vs. Achievement.xlsx`** to view the completed report.
