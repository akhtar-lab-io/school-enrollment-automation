# 📊 Automated Student Data Recapitulation System

A simple yet powerful Command Line Interface (CLI) automation tool built with **Python** and **Pandas**. This project was specifically designed to tackle the real-world administrative challenges of managing rapidly growing student enrollment data, replacing time-consuming manual entry with structured automation that exports directly to Microsoft Excel.

---

## 💡 The Problem & Motivation

With the school's growing reputation, student enrollment numbers surged significantly. Processing and typing extensive detailed columns for each student (ranging from personal credentials, parents' info, to background school history) one-by-one directly into Excel became a tedious, error-prone, and exhausting task.

To solve this, this automated CLI script was developed. Instead of dealing with clunky spreadsheets and worrying about accidentally overwriting existing cells, the operator can seamlessly input data sequentially directly from the terminal.

## ✨ Key Features

* **Terminal-Based Guided Input:** User-friendly command-line prompts ensuring smooth data entry for comprehensive student records.
* **Smart Auto-Increment Indexing (`No`):** The system dynamically scans the existing Excel sheet, tracks the last entry, and automatically increments the index number for new entries without overlapping.
* **Automated Excel Compiling:** Merges old data seamlessly with newly inputted records using `pandas.concat` upon program exit.
* **Safe-Stop Command:** Simply type `exit` or answer the prompt to safely stop and trigger the database export mechanism instantly.

## 🛠️ Tech Stack & Requirements

* **Language:** Python 3.12+
* **Libraries:** 
  * `pandas` (For data manipulation and DataFrame structures)
  * `openpyxl` (Excel file engine integration)

---

## 🚀 How to Run Locally

Follow these quick steps to set up and run the script on your machine:

### 1. Clone the Repository
```bash
git clone (https://github.com/akhtar-lab-io/school-enrollment-automation)
cd school_enrollment_automation

### 2. Install dependecies
pip install pandas openpyxl

### 3. Run the Progam
python input_siswa.py
