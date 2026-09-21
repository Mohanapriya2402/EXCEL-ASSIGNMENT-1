# 📊 Excel Data Exploration & Analysis – Product Dataset

## 📌 Project Overview

This project is a beginner-level **Data Analysis project using Microsoft Excel**. The objective is to perform basic data exploration, summarization, conditional analysis, and text manipulation on a Product Dataset.

As an aspiring **Data Analyst**, I used Excel formulas and functions to extract meaningful insights from product data and strengthen my foundational skills in data analysis.

---

## 🎯 Objectives

The main objectives of this project are:

* Perform basic data exploration using Excel
* Calculate summary statistics
* Apply logical functions
* Perform conditional calculations
* Extract information from text fields
* Categorize products based on price
* Build a foundation for further data analysis and visualization

---

## 📂 Dataset

The dataset contains information about different products.

### Dataset Columns

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Price ($)    | Price of the product               |
| Quantity     | Available quantity                 |
| Category     | Product category                   |

The dataset contains **34 populated product records**.

---

# 📝 Tasks Performed

## 1. Basic Data Exploration

### 🔹 Total Price

Used the `SUM` function to calculate the total price of all products.

```excel
=SUM(D2:D35)
```

**Result: $10,100**

### 🔹 Number of Products

Used the `COUNT` function to determine the number of products.

```excel
=COUNT(D2:D35)
```

**Result: 34 products**

### 🔹 Average Price

Used the `AVERAGE` function to calculate the average product price.

```excel
=AVERAGE(D2:D35)
```

**Result: $297.06**

---

## 2. Minimum and Maximum Price

### 🔹 Minimum Price

Used the `MIN` function.

```excel
=MIN(D2:D35)
```

**Result: $30**

### 🔹 Maximum Price

Used the `MAX` function.

```excel
=MAX(D2:D35)
```

**Result: $1,000**

---

## 3. Logical Function – IF

Created a new column named **Price Range** to categorize products based on their price.

### Classification Rule

* Price ≥ $500 → **High Price**
* Price < $500 → **Standard Price**

### Excel Formula

```excel
=IF(D2>=500,"High Price","Standard Price")
```

This formula was applied to all product records.

---

## 4. Conditional Functions – SUMIF and COUNTIF

### 🔹 Total Price of Electronics Products

Used the `SUMIF` function to calculate the total price of products belonging to the **Electronics** category.

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

**Result: $8,050**

### 🔹 Number of Products Below $100

Used the `COUNTIF` function to count products with a price less than $100.

```excel
=COUNTIF(D2:D35,"<100")
```

**Result: 11 products**

---

# 5. Text Manipulation

The `Product ID` column contains information in the following format:

```text
28-JAN-US
```

Excel text functions were used to extract different parts of the Product ID.

---

### 🔹 Day

Created a new column named **Day** using the `LEFT` function.

```excel
=LEFT(A2,2)
```

Example:

```text
28-JAN-US → 28
```

---

### 🔹 Country Code

Created a new column named **Country Code** using the `RIGHT` function.

```excel
=RIGHT(A2,2)
```

Example:

```text
28-JAN-US → US
```

---

### 🔹 Month

Created a new column named **Month** using the `MID` function.

```excel
=MID(A2,4,3)
```

Example:

```text
28-JAN-US → JAN
```

---

# 📊 Project Results

| Analysis                |  Result |
| ----------------------- | ------: |
| Total Product Price     | $10,100 |
| Number of Products      |      34 |
| Average Price           | $297.06 |
| Minimum Price           |     $30 |
| Maximum Price           |  $1,000 |
| Electronics Total Price |  $8,050 |
| Products Below $100     |      11 |

---

# 🛠️ Excel Functions Used

The following Excel functions were used in this project:

* `SUM`
* `COUNT`
* `AVERAGE`
* `MIN`
* `MAX`
* `IF`
* `SUMIF`
* `COUNTIF`
* `LEFT`
* `RIGHT`
* `MID`

---

# 💡 Skills Demonstrated

Through this project, I practiced the following Data Analyst skills:

* Data Exploration
* Data Summarization
* Excel Formulas
* Logical Functions
* Conditional Aggregation
* Text Manipulation
* Data Categorization
* Basic Data Cleaning and Preparation

---

# 📁 Project Structure

```text
Excel-Data-Exploration/
│
├── README.md
│
└── Excel Assignment 1 - Data Exploration.xlsx
```

---

# 🚀 Learning Outcome

This project helped me understand how Excel can be used to explore and summarize datasets using formulas and functions.

It is part of my learning journey toward becoming a **Data Analyst**, where I am building practical skills in:

**Excel → SQL → Python → Power BI → Data Analytics Projects**

---

## 👩‍💻 About Me

I am an aspiring **Data Analyst** building my skills through hands-on projects and practical datasets.

This repository documents my learning journey and projects as I develop my Data Analytics portfolio.
# EXCEL-ASSIGNMENT-1
