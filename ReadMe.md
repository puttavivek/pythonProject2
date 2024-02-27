# Inventory Assistant

## Introduction
This is a simple inventory assistant implemented in Python using the Tkinter GUI library and Pandas for data manipulation. The system allows users to perform various operations such as adding, updating, searching, and printing inventory data.

## Requirements
- Python 3.x
- Dependencies (Install using `pip install -r requirements.txt`):
  - matplotlib
  - pandas
  - openpyxl
  - numpy

## Installation
1. **Clone the repository:**
   Clone this repository to your local machine
   ```bash
   git clone https://github.com/your_username/inventory-management.git
   cd inventory-management
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the main application file:**
   ```bash
   python main.py
   ```

## User Input Details

### Add Inventory:
- Part Name: String
- Part No: String
- Model: String
- Stock Location: Integer (Between 0 and 100)
- Quantity: Integer (Between 1 and 10000)

### Edit Inventory:
- Part No: String
- Quantity: Integer (Between 1 and 10000)

### Search Inventory:
- Part No: String

### Bulk Input:
Select a file containing bulk inventory data (Excel format).

Follow the on-screen instructions to perform inventory management tasks.

## Features:
- Add new inventory items or update existing ones.
- Search for items based on part numbers.
- View the entire inventory list.
- Print the inventory data to a PDF file.
- Bulk entry of inventory data from a file.

