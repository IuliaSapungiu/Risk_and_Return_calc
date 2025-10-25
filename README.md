# Risk and Return Analysis – Bittnet Systems S.A. (2020–2024)\
###### Undergraduate Bachelor Thesis Project – Financial Analysis of Bittnet Systems S.A. (2020–2024). 
###### All rights reserved © Iulia Sapungiu

## Table of contents

 - [Project Overview](#project-overview)
 - [Features](#features)
 - [Technologies](#technologies)
 - [Installation](#installation)
 - [Usage](#usage)
 - [Project Structure](#project-structure)


### Project Overview

This project was developed as part of my Bachelor Thesis, focusing on a quantitative financial analysis of Bittnet Systems S.A. over the 2020–2024 period.
It implements a Python-based analytical pipeline for computing, comparing, and exporting financial ratios derived from the company’s balance sheet and income statement data.

The script uses structured data processing with **pandas** and **numpy**, ensuring computational consistency through shared intermediate variables (e.g., ```total_assets```, ```net_income```,  ```inventories```).
All expense or cost values are automatically normalized to absolute values to maintain mathematical correctness.
The final output is a multi-sheet Excel report generated via **openpyxl**, where each sheet corresponds to a specific analytical category of ratios.


### Features

The project performs a structured and automated financial analysis workflow:
- Loads Bittnet’s financial statements into a   ``` pandas.DataFrame ``` for structured data processing
- Normalizes numerical fields and ensures consistent data types across periods
- Defines intermediate variables (e.g.,   ```total_assets```,   ```net_income ```,   ```inventories```) once and reuses them across all ratio calculations
- Converts all expense and cost values to their absolute values to preserve mathematical correctness
- Applies deterministic, vectorized formulas using   ```NumPy``` for optimized multi-year computations
- Aggregates computed ratios into grouped DataFrame objects, maintaining a standardized layout for temporal comparison
- Exports results to a multi-sheet Excel file via   ```openpyxl```, where each sheet represents a distinct analytical category


### Technologies
- Python 3.8 or later
- pandas – Data ingestion and transformation
- numpy – Vectorized financial computations
- openpyxl – Structured Excel export


### Installation

OPTIONAL: Git installed on your machine (optional, for cloning the repository).

1. Clone the repository: **https://github.com/IuliaSapungiu/Risk_and_Return_calc.git**
2. Navigate to the project directory: **cd Data_Science**
3. Create a virtual environment: **python -m venv env_name**
4. Activate the virtual environment:

    - **On Windows:**
  
      ```
      source env_name\Scripts\activate
      ```

    - **On macOS and Linux:**

      ```
      source env_name/bin/activate
      ```

5. Install the required dependencies:

      ```
      pip install -r requirements.txt
      ```

6. Run the project:
    ```
    python bittnet_ratios.py  
    ```


### Usage

1. Once you run the project, the script automatically:

- Loads and processes the dataset
- Computes all ratio categories
- Exports results to an Excel file named ```ratios.xlsx``` inside the ```outputs/``` directory

  ! Before running, make sure a folder named ```outputs``` exists in the project root (create it manually if necessary). !

2. Open ```outputs/ratios.xlsx```

  Each sheet corresponds to a financial ratio category, and each column displays annual values for 2020–2024.

  
### Project Structure

```
Risk_and_Return_calc/
│
├── .gitignore
├── README.md
│
└── ratios/
    │
    ├── BI visualizations/
    ├── reports/
    ├── outputs/              # creat manual, pentru fisierele generate
    │
    ├── bittnet_ratios.py
    ├── requirements.txt
```
 
