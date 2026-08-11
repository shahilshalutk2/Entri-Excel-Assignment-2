# Entri-Excel-Assignment-2
# Excel Power Query Data Cleaning & Transformation

## 📌 Project Overview

This project demonstrates data cleaning, transformation, formatting, and conditional formatting using **Microsoft Excel and Power Query**.

The dataset contains product information such as Product Name, Brand Name, Price, Quantity, Category, Manufacturing Date, and Country Code.

The main objective of this assignment is to clean the raw dataset and prepare it for further analysis.

---

## 🛠️ Tools Used

* Microsoft Excel
* Power Query
* Excel Conditional Formatting

---

## 📊 Dataset

The dataset contains the following information:

* Manufacturing Date
* Country Code
* Product Name
* Brand Name
* Price ($)
* Quantity
* Category

The original Product ID contained information about the day, month, and country code. This was split and transformed into separate columns.

---

## 🔧 Data Cleaning & Transformation

### 1. Handling Missing Values

Missing values were identified in the **Price** column.

Three missing price values were identified and replaced using the average price.

Missing categories were also identified and assigned based on the product type and existing product-category relationships.

---

### 2. Correcting Inconsistent Data

Product names contained inconsistent capitalization, such as:

* `laptop`
* `Laptop`
* `headphones`
* `Headphones`

The text was standardized using **Capitalize Each Word**.

A category typo was also identified:

`Electroni` → `Electronics`

The typo was corrected using **Find and Replace**.

---

### 3. Removing Duplicates

Duplicate rows were identified by selecting all columns in the dataset.

The **Remove Duplicates** option in Power Query was used to remove duplicate records.

---

### 4. Splitting and Merging Data

The **Product ID** was split using the `-` delimiter into:

* Day
* Month
* Country Code

The Day and Month values were then combined with the year to create the **Manufacturing Date**.

The **Brand Name** and **Product Name** columns were merged using a space separator to create the **Product Brand** column.

Example:

`Dell + Laptop → Dell Laptop`

---

### 5. Number Formatting

The **Price** column was changed to the **Decimal Number** data type in Power Query and formatted as currency in Excel.

The **Manufacturing Date** column was converted to Date format and displayed as:

`DD-MM-YYYY`

Example:

* `28-01-2026`
* `15-02-2026`
* `03-03-2026`

---

### 6. Conditional Formatting

Data Bars were applied to the **Price** column to visually compare product prices.

A custom conditional formatting rule was applied to the **Category** column to highlight Electronics:

```excel
=G2="Electronics"
```

This highlights all cells where the category is **Electronics**.

---

## 📁 Project Structure

```text
Excel-PowerQuery-Data-Cleaning/
│
├── Assignment 2 Entri.pdf
├── README.md
└── Excel Dataset.xlsx
```

---

## 🎯 Learning Outcomes

Through this project, I learned how to:

* Import and transform data using Power Query
* Identify and handle missing values
* Standardize inconsistent text
* Correct data-entry errors
* Remove duplicate records
* Split columns using delimiters
* Merge multiple columns
* Change data types
* Format dates and currency
* Apply conditional formatting
* Prepare a dataset for analysis

---

## 👤 Author

**Shahil shalu TK**

This project was completed as part of an Excel and Power Query data-cleaning assignment.
