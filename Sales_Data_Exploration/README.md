# Excel Product Sales Data Exploration

## 📌 Project Overview

This project explores a product dataset using _**Microsoft Excel**_ and its built-in formulas and functions. The objective is to practice Data Exploration, summarize product information, classify products by price, analyze category-based results, and extract useful information from product IDs.

This project demonstrates foundational Excel data analysis skills through practical exercises using product prices, quantities, brands, and categories.

## 📂 Dataset Description

In this project, the primary dataset, named **`Product Dataset`**, was renamed to **`Sales_DB_Dataset`** in the `Excel_Product_Sales_Data_Exploration.xlsx` workbook. It contains **34 product records** with the following columns:

| Column | Description |
|---|---|
| Product ID | Unique product identifier containing day, month, and country-code information |
| Product Name | Name of the product |
| Brand Name | Brand associated with the product |
| Price ($) | Listed product price in US dollars |
| Quantity | Quantity recorded for the product |
| Category | Category to which the product belongs |

**Product categories:** Electronics, Fashion, Kitchen, Outdoor, and Accessories.

## 📑 Workbook Structure

| Worksheet | Description |
|---|---|
| `Sales_DB_Dataset` | Primary Product 	Dataset |
| `Sales_DB(Copy)` | Working copy with additional derived columns |
| `Aggregate_Functions` | Exercises using `SUM`, `COUNT`, `AVERAGE`, `MIN`, and `MAX` |
| `Logical_IF` | Product price classification using the `IF` function |
| `Conditional_Functions` | Category and price analysis using `SUMIF` and `COUNTIF` |
| `Text_Functions` | Extracting information from Product ID using `LEFT`, `RIGHT`, and `MID` |

## 🛠️ Excel Functions Used

### 🧮 1. Aggregate Functions

Aggregate functions help summarize numerical data and calculate useful statistics.

| Function | Purpose | Excel Formula |
|---|---|---|
| `SUM()` | Calculates the sum of listed product prices. | `=SUM(Sales_DB[Price ($)])` |
| `COUNTA()` | Counts the number of products in the dataset. | `=COUNTA(Sales_DB[Product Name])` |
| `AVERAGE()` | Calculates the average product price. | `=AVERAGE(Sales_DB[Price ($)])` |
| `MIN()` | Finds the minimum product price. | `=MIN(Sales_DB[Price ($)])` |
| `MAX()` | Finds the maximum product price. | `=MAX(Sales_DB[Price ($)])` |

**Results recorded in the workbook:**

| Analysis | Result |
|---|---:|
| Sum of listed product prices | $10,100 |
| Number of product records | 34 |
| Average listed product price | $297.06 |
| Minimum listed product price | $30 |
| Maximum listed product price | $1,000 |

> **Note:** The sum of listed product prices is not total sales revenue. To calculate revenue, multiply each product's price by its quantity and then sum the results.

### 🤔 2. Logical Function — IF()

The `IF()` function classifies products into price ranges based on a specified condition.

**Classification criteria:**
- Price greater than or equal to $500: **High Price**
- Price below $500: **Standard Price**

**Excel formula:**

```excel
=IF([@[Price ($)]]>=500,"High Price”,“Standard Price")
```

**Results recorded in the workbook:** 

**Example Result – Price Range Classification:**

| Product Name | Price ($) | Price Range |
|---|---:|---|
| Laptop | 1,000 | High Price |
| Sneakers | 80 | Standard Price |
| Coffee Maker | 130 | Standard Price |
| Smartphone | 900 | High Price |
| Backpack | 70 | Standard Price |
| Headphones | 200 | Standard Price |
| T-shirt | 30 | Standard Price |
| Blender | 90 | Standard Price |
| Tablet | 500 | High Price |
| Hiking Boots | 130 | Standard Price |

**Classification Summary (All 34 Records):**

| Price Range | Number of Products |
|---|---:|
| High Price | 8 |
| Standard Price | 26 |
| **Total Products** | **34** |

> **Note:** Out of 34 product records, 8 products are classified as High Price (≥ $500), while the remaining 26 products are classified as Standard Price (< $500).

### 🎯 3. Conditional Functions — SUMIF() and COUNTIF()

These functions calculate results based on specified conditions.

| Analysis | Result |
|---|---:|
| Sum of listed prices for Electronics | $8,050 |
| Number of products priced below $100 | 11 |

**Excel formulas:**

Calculates the sum of prices for products in the Electronics category.

```excel
=SUMIF(Sales_DB[Category],"Electronics",Sales_DB[Price ($)])
```

Counts products with prices below $100.

```excel
 =COUNTIF(Sales_DB[Price ($)],"<100")
```

> **Note:** Adjust the column references if your worksheet uses a different layout.*

### 🔤 4. Text Functions — LEFT(), RIGHT() and MID()

Text functions extract specific characters from product identifiers to separate useful information such as day, month, and country code.

**Excel Functions and Formulas:**

| Function | Purpose | Excel Formula |
|---|---|---|
| `LEFT()` | Extracts the first two characters as the day. | `=LEFT([@[Product ID]],2)` |
| `RIGHT()` | Extracts the last two characters as the country code. | `=RIGHT([@[Product ID]],2)` |
| `MID()` | Extracts the three-character month abbreviation. | `=MID([@[Product ID]],4,3)` |

**Example Result – Text Extraction:**

| Product ID | Day (`LEFT`) | Country Code (`RIGHT`) | Month (`MID`) |
|---|---|---|---|
| `28-JAN-US` | 28 | US | JAN |
| `15-FEB-US` | 15 | US | FEB |
| `03-MAR-US` | 03 | US | MAR |
| `11-APR-US` | 11 | US | APR |
| `22-MAY-US` | 22 | US | MAY |
| `07-JUN-UK` | 07 | UK | JUN |
| `19-JUL-UK` | 19 | UK | JUL |
| `23-AUG-UK` | 23 | UK | AUG |
| `05-SEP-UK` | 05 | UK | SEP |
| `14-OCT-UK` | 14 | UK | OCT |

**Example:**

For the Product ID `28-JAN-US`:

- **Day:** `28` — Extracted using `LEFT()`.
- **Month:** `JAN` — Extracted using `MID()`.
- **Country Code:** `US` — Extracted using `RIGHT()`.

> **Note:** The text functions successfully extract the day, month, and country code from Product IDs. This helps transform structured text into separate columns, making the dataset easier to organize and analyze.

### 📑 Final Column Structure

| Column | Field Name |
|---|---|
| A | `Product ID` |
| B | `Product Name` |
| C | `Brand Name` |
| D | `Category` |
| E | `Price ($)` |
| F | `Quantity` |
| G | `Price Range` |
| H | `Day` |
| I | `Country Code` |
| J | `Month` |

**Workbook Structure:** 34 Records | 10 Columns (A–J).

## 📝 Summary of Excel Functions

| Function | Description | Formula |
|---|---|---|
| 🧮 `SUM()` | Calculates the sum of product prices. | `=SUM(D2:D35)` |
| 🔢 `COUNT()` | Counts the number of numeric price entries. | `=COUNT(D2:D35)` |
| 📊 `AVERAGE()` | Calculates the average product price. | `=AVERAGE(D2:D35)` |
| 📉 `MIN()` | Finds the minimum product price. | `=MIN(D2:D35)` |
| 📈 `MAX()` | Finds the maximum product price. | `=MAX(D2:D35)` |
| 🤔 `IF()` | Classifies products based on a price condition. | `=IF(D2>=500,"High Price","Standard Price")` |
| 🎯 `SUMIF()` | Calculates the sum of prices for a specified category. | `=SUMIF(F2:F35,"Electronics",D2:D35)` |
| 🔍 `COUNTIF()` | Counts products priced below $100. | `=COUNTIF(D2:D35,"<100")` |
| 🔤 `LEFT()` | Extracts the first two characters of a Product ID. | `=LEFT(A2,2)` |
| 🔡 `RIGHT()` | Extracts the last two characters of a Product ID. | `=RIGHT(A2,2)` |
| ✂️ `MID()` | Extracts the three-character month abbreviation from a Product ID. | `=MID(A2,4,3)` |

### 🎯 Overall Summary

These Excel functions help summarize numerical data, classify products, perform conditional calculations, and extract information from text. They support efficient product data exploration and meaningful insights using Microsoft Excel.

## 📊 Key Observations

- The workbook contains **34 product records**.
- **Electronics** is the largest category, with 15 records.
- Product prices range from **$30 to $1,000**.
- The average listed product price is approximately **$297.06**.
- The sum of listed product prices is **$10,100**.
- Excel formulas can simplify calculations, classifications, conditional analysis, and text extraction.

## 💡 Skills Demonstrated

- Microsoft Excel Data Exploration
- Data summarization using aggregate functions
- Logical conditions and product classification
- Conditional calculations using `SUMIF()` and `COUNTIF()`
- Text manipulation using `LEFT()`, `RIGHT()`, and `MID()`
- Formula-based analysis and reporting
- Organizing data across multiple worksheets

## 🧰 Tools Used

- **Microsoft Excel**
- Excel formulas and functions

## 🚀 How to Use This Project

1. Download or clone the **Excel-Projects** repository.
2. Go to `Sales_Data_Exploration` Folder
3. In that, Open `Excel_Product_Sales_Data_Exploration.xlsx` in Microsoft Excel.
4. Explore the original dataset in the `Sales_DB_Dataset` worksheet.
5. Review the derived columns in `Sales_DB(Copy)`.
6. Explore the function-specific worksheets to understand the formulas and results.
7. Modify formula criteria or add sample records to practice Excel data analysis.

## 🔮 Future Improvements

Potential enhancements for this project include:

- Adding a **Revenue** column calculated as `Price × Quantity`.
- Creating PivotTables to compare products and categories.
- Building charts to visualize product prices and category distribution.
- Applying conditional formatting to highlight important values.
- Developing an interactive Excel dashboard with KPIs and slicers.
- Expanding the analysis with additional business questions and insights.

## 📁 Project Files & Documents

`Excel_Product_Sales_Data_Exploration.xlsx`
`Excel_Product_Sales_Data_Exploration_Document.pdf`

## 👤 Author

**Manikandan S.**

Aspiring Data Analyst | Learning Excel and developing practical data analysis projects.

---

## 📚 Acknowledgement

*This project is developed for **educational and learning purposes** to demonstrates foundational Excel data exploration and formula-based analysis skills.*
 
