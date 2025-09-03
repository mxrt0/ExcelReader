using ExcelReader.Context;
using ExcelReader.Models;
using OfficeOpenXml;
namespace ExcelReader.Services;

public class ProductReader
{
    private ProductsDbContext _dbContext;
    private string excelFilePath = @"..\..\..\products.xlsx";
    public ProductReader(ProductsDbContext dbContext)
    {
        _dbContext = dbContext;
        _dbContext.Database.EnsureCreated();
    }

    public async Task Run()
    {
        var products = ReadProductsFromSpreadsheet(excelFilePath);
        await SaveProductsToDatabase(products);

    }
    public IEnumerable<Product> ReadProductsFromSpreadsheet(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];

        int rowCount = worksheet.Dimension.End.Row;
        int colCount = worksheet.Dimension.End.Column;

        var products = new List<Product>();
        for (int row = 2; row <= rowCount; row++)
        {
            yield return new Product(
                worksheet.Cells[row, 1].Text,                       // Name
                decimal.Parse(worksheet.Cells[row, 2].Text),        // UnitPrice
                int.Parse(worksheet.Cells[row, 3].Text),            // Quantity
                decimal.Parse(worksheet.Cells[row, 4].Text),        // TotalPrice
                decimal.Parse(worksheet.Cells[row, 5].Text)         // TotalPriceWithVAT
            );
        }
    }

    public async Task SaveProductsToDatabase(IEnumerable<Product> products)
    {
        foreach (var product in products)
        {
            _dbContext.Products.Add(product);
        }
        await _dbContext.SaveChangesAsync();
    }
}
