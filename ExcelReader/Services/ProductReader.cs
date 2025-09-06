using ExcelReader.Context;
using ExcelReader.Models;
using OfficeOpenXml;
namespace ExcelReader.Services;

public class ProductReader
{
    private ProductsDbContext _dbContext;
    private GenericDbContext? _genericContext;
    private string excelFilePath = @"..\..\..\products.xlsx";
    public ProductReader(ProductsDbContext dbContext)
    {
        AnsiConsole.MarkupLine("[aqua]Creating [lime][bold]Database Table[/][/]...[/]\n");
        _dbContext = dbContext;
        _dbContext.Database.EnsureDeleted();
        _dbContext.Database.EnsureCreated();
        AnsiConsole.MarkupLine("[aqua]Created [lime][bold]Database Table[/][/]![/]\n");
    }

    public async Task Run()
    {
        await SeedExcelData();
        var products = ReadProductsFromSpreadsheet(excelFilePath);
        await SaveProductsToDatabase(products);
    }

    public async Task RunGeneric()
    {
        await SeedExcelData();
        var data = ReadDataFromSpreadsheet(excelFilePath);
        await SaveDataToDatabase(data);
    }

    public async Task SeedExcelData()
    {
        int productsCount = 300_000;
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Products");

        worksheet.Cells[1, 1].Value = "ProductName";
        worksheet.Cells[1, 2].Value = "UnitPrice";
        worksheet.Cells[1, 3].Value = "Quantity";
        worksheet.Cells[1, 4].Value = "TotalPrice";
        worksheet.Cells[1, 5].Value = "TotalPriceWithVAT";
        worksheet.Cells[1, 6].Value = "RandomHeader";

        int row = 2;
        AnsiConsole.Progress()
            .Start(ctx =>
            {
                var task = ctx.AddTask("[aqua]Generating [lime][bold]Excel[/][/] Worksheet[/]\n", maxValue: productsCount);
                foreach (var product in GenerateProducts(productsCount))
                {
                    worksheet.Cells[row, 1].Value = product.ProductName;
                    worksheet.Cells[row, 2].Value = product.UnitPrice;
                    worksheet.Cells[row, 3].Value = product.Quantity;
                    worksheet.Cells[row, 4].Value = product.TotalPrice;
                    worksheet.Cells[row, 5].Value = product.TotalPriceWithVAT;
                    worksheet.Cells[row, 6].Value = product.Random;
                    row++;
                    task.Increment(1);
                }
            });

        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

        await package.SaveAsAsync(new FileInfo(excelFilePath));

        AnsiConsole.MarkupLine("[aqua]Generated [lime][bold]Excel[/][/] Worksheet![/]\n");
    }
    public IEnumerable<Product> ReadProductsFromSpreadsheet(string filePath)
    {
        AnsiConsole.MarkupLine("[aqua]Opening [lime][bold]Excel[/][/] worksheet...[/]\n");
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];
        AnsiConsole.MarkupLine("[aqua]Opened [lime][bold]Excel[/][/] worksheet![/]");

        int rowCount = worksheet.Dimension.End.Row;
        int colCount = worksheet.Dimension.End.Column;

        var products = new List<Product>();
        AnsiConsole.Progress()
            .Start(ctx =>
            {
                var task = ctx.AddTask("[aqua]Reading Data From [lime][bold]Excel[/][/] Worksheet[/]", maxValue: rowCount - 1);

                for (int row = 2; row <= rowCount; row++)
                {
                    products.Add(new Product(
                        worksheet.Cells[row, 1].Text,                       // Name
                        decimal.Parse(worksheet.Cells[row, 2].Text),        // UnitPrice
                        int.Parse(worksheet.Cells[row, 3].Text),            // Quantity
                        decimal.Parse(worksheet.Cells[row, 4].Text),        // TotalPrice
                        decimal.Parse(worksheet.Cells[row, 5].Text)         // TotalPriceWithVAT
                    ));
                    task.Increment(1);
                }
            });
        return products;
    }

    public IEnumerable<Models.ExcelRow> ReadDataFromSpreadsheet(string filePath)
    {
        AnsiConsole.MarkupLine("[aqua]Opening [lime][bold]Excel[/][/] worksheet...[/]\n");
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];
        AnsiConsole.MarkupLine("[aqua]Opened [lime][bold]Excel[/][/] worksheet![/]");

        int rowCount = worksheet.Dimension.End.Row;
        int colCount = worksheet.Dimension.End.Column;

        var headers = new List<string>();
        for (int col = 1; col <= colCount; col++)
        {
            var header = worksheet.Cells[1, col].Text.Trim();
            headers.Add(string.IsNullOrEmpty(header) ? $"Column{col}" : header);
        }
        _genericContext = new GenericDbContext(headers);
        _genericContext.Database.EnsureDeleted();
        _genericContext.Database.EnsureCreated();
        var rows = new List<Models.ExcelRow>();
        AnsiConsole.Progress()
            .Start(ctx =>
            {
                var task = ctx.AddTask("[aqua]Reading Data From [lime][bold]Excel[/][/] Worksheet[/]", maxValue: rowCount - 1);

                for (int row = 2; row <= rowCount; row++) // start after header
                {
                    var entity = new Models.ExcelRow();

                    var entry = _genericContext.Entry(entity);
                    for (int col = 1; col <= colCount; col++)
                    {
                        var header = headers[col - 1];
                        var colValue = worksheet.Cells[row, col].Text.Trim();
                        entry.Property(header).CurrentValue = colValue;
                    }
                    task.Increment(1);
                    rows.Add(entity);
                }
            });
        return rows;
    }

    public async Task SaveProductsToDatabase(IEnumerable<Product> products)
    {
        AnsiConsole.MarkupLine("[aqua]Saving [lime][bold]Products[/][/] To Database...[/]\n");
        foreach (var product in products)
        {
            _dbContext.Products.Add(product);
        }
        await _dbContext.SaveChangesAsync();
        AnsiConsole.MarkupLine("[aqua]Saved [lime][bold]Products[/][/] To Database![/]");
    }
    public async Task SaveDataToDatabase(IEnumerable<Models.ExcelRow> items)
    {
        AnsiConsole.MarkupLine("[aqua]Saving [lime][bold]Data[/][/] To Database...[/]\n");
        foreach (var item in items)
        {
            _genericContext!.Items.Add(item);
        }
        await _genericContext!.SaveChangesAsync();
        AnsiConsole.MarkupLine("[aqua]Saved [lime][bold]Data[/][/] To Database![/]");
    }
    private IEnumerable<Product> GenerateProducts(int count)
    {
        var rand = new Random();

        for (int i = 0; i < count; i++)
        {
            string name = $"Product {i + 1}";
            decimal unitPrice = Math.Round((decimal)(rand.NextDouble() * 10 + 0.5), 2);
            int quantity = rand.Next(1, 20);
            string random = $"Random {i + 1}";
            yield return new Product(name, unitPrice, quantity, random);
        }
    }
}
