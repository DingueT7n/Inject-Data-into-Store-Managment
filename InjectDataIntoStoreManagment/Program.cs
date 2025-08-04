using System.Globalization;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace ExcelToDatabase
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("Enter the path to your Excel file:");
            string filePath = Console.ReadLine()?.Trim();

            Console.WriteLine("Enter the SQL Server connection string:");
            string connectionString = Console.ReadLine()?.Trim();

            Console.WriteLine("Enter the Excel column index for CodeBar (1-based):");
            int codeBarCol = int.Parse(Console.ReadLine() ?? "1");

            Console.WriteLine("Enter the Excel column index for Product Name:");
            int nameCol = int.Parse(Console.ReadLine() ?? "2");

            Console.WriteLine("Enter the Excel column index for Buy Price:");
            int buyPriceCol = int.Parse(Console.ReadLine() ?? "3");

            Console.WriteLine("Enter the Excel column index for Sale Price:");
            int salePriceCol = int.Parse(Console.ReadLine() ?? "4");

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File not found.");
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    Console.WriteLine("No worksheet found.");
                    return;
                }

                int rowCount = worksheet.Dimension.Rows;

                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string codeBar = worksheet.Cells[row, codeBarCol].Text?.Trim();
                        string productName = worksheet.Cells[row, nameCol].Text?.Trim();
                        string buyPriceStr = worksheet.Cells[row, buyPriceCol].Text?.Trim();
                        string salePriceStr = worksheet.Cells[row, salePriceCol].Text?.Trim();

                        if (string.IsNullOrWhiteSpace(codeBar) || codeBar.Length < 5 || string.IsNullOrWhiteSpace(productName))
                            continue;

                        // Convert scientific notation to plain format
                        if (double.TryParse(codeBar, out double parsed))
                            codeBar = parsed.ToString("F0", CultureInfo.InvariantCulture);

                        if (!decimal.TryParse(buyPriceStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal buyPrice))
                            buyPrice = 0;
                        if (!decimal.TryParse(salePriceStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal salePrice))
                            salePrice = 0;

                        // Check for duplicates
                        string checkQuery = @"
                            SELECT COUNT(1)
                            FROM Products
                            WHERE CodeBar = @CodeBar OR ProductName = @ProductName";

                        using (var checkCmd = new SqlCommand(checkQuery, connection))
                        {
                            checkCmd.Parameters.AddWithValue("@CodeBar", codeBar);
                            checkCmd.Parameters.AddWithValue("@ProductName", productName);

                            int exists = (int)checkCmd.ExecuteScalar();
                            if (exists > 0)
                            {
                                Console.WriteLine($"Row {row} skipped (already exists): {productName}");
                                continue;
                            }
                        }

                        // Insert the product
                        string insertQuery = @"
                            INSERT INTO Products (
                                ProductName, CodeBar, Quentity, BuyPrice, SalePrice, SalePriceSeconde, 
                                WholesalePrice, TaxValue, SalePriceTax, IsTaxed, DateExperation, MinQty, 
                                MaxDiscount, NbrProductInColis, ColisageBareCode, ProductImage, CategoryId
                            ) VALUES (
                                @ProductName, @CodeBar, 0, @BuyPrice, @SalePrice, @SalePrice, 
                                @SalePrice, 0, @SalePrice, 0, GETDATE(), 0, 
                                0, 0, NULL, NULL, 1
                            )";

                        using (var insertCmd = new SqlCommand(insertQuery, connection))
                        {
                            insertCmd.Parameters.AddWithValue("@ProductName", productName);
                            insertCmd.Parameters.AddWithValue("@CodeBar", codeBar);
                            insertCmd.Parameters.AddWithValue("@BuyPrice", buyPrice);
                            insertCmd.Parameters.AddWithValue("@SalePrice", salePrice);
                            insertCmd.ExecuteNonQuery();
                        }

                        Console.WriteLine($"Row {row} inserted: {productName}");
                    }
                }
            }

            Console.WriteLine("Import complete.");
        }
    }
}
