ExcelToDatabaseImporter
This project is a command-line tool written in C# (.NET) that reads an Excel file containing product data and inserts it into a SQL Server database.

 Features
 Prompts the user for:

Excel file path

SQL Server connection string

Column indexes for CodeBar, Product Name, Buy Price, and Sale Price

 Reads product data from .xlsx files using EPPlus

 Validates and sanitizes data, including:

Removing scientific notation in barcodes

Skipping invalid or incomplete rows

 Avoids duplicates:

Checks if a product already exists by CodeBar or ProductName

 Inserts new products using parameterized SQL into the Products table

 Designed for use in retail or inventory management systems

 Use Case
Useful for:

Quickly importing product catalogs from Excel into your store management system

Initializing a product database during development or setup

Avoiding manual entry of thousands of items

 Tech Stack
.NET 6 / 7 / 8 Console App

System.Data.SqlClient for direct SQL execution

EPPlus for reading Excel .xlsx files
