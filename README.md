# Excel Assignment 1: Data Exploration

This project contains an Excel-based data exploration assignment using a product dataset. The assignment demonstrates basic Excel functions for analyzing product prices, quantities, categories, and brands.

## 📊 Dataset

The dataset contains product information with the following columns:

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Product brand                      |
| Price ($)    | Price of the product               |
| Quantity     | Available quantity                 |
| Category     | Product category                   |

The dataset contains **34 products** across categories such as Electronics, Fashion, Kitchen, and Outdoor.

## 📝 Assignment Tasks

### 1. Sum, Count and Average

* Calculate the total price of all products.
* Count the number of products.
* Calculate the average product price.

### 2. Minimum and Maximum

* Find the minimum product price.
* Find the maximum product price.

### 3. IF Function

A new **Price Range** column is created using the `IF` function.

Products are categorized as:

* **High Price**: Price greater than or equal to $500
* **Standard Price**: Price below $500

Example formula:

```excel
=IF(D2>=500,"High Price","Standard Price")
```

### 4. SUMIF and COUNTIF

The assignment uses conditional Excel functions to analyze the dataset based on specific categories and conditions.

Example:

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

```excel
=COUNTIF(F2:F35,"Electronics")
```

## 🛠️ Excel Functions Used

* `SUM()`
* `COUNT()`
* `AVERAGE()`
* `MIN()`
* `MAX()`
* `IF()`
* `SUMIF()`
* `COUNTIF()`

## 📁 Files

```text
Excel-Assignment-1/
│
├── Excel Assignment 1 - Data Exploration.xlsx
└── README.md
```

## 🎯 Objective

The objective of this assignment is to understand and apply fundamental Excel functions for:

* Data exploration
* Statistical calculations
* Conditional classification
* Category-based analysis
* Product data analysis
