global using Spectre.Console;
using ExcelReader.Context;
using ExcelReader.Services;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
namespace ExcelReader;

public class Program
{
    static async Task Main()
    {
        ExcelPackage.License.SetNonCommercialPersonal("mxrt0");
        var options = new DbContextOptionsBuilder<ProductsDbContext>().UseSqlite(@"Data Source=..\..\..\products.db").Options;
        var dbContext = new ProductsDbContext(options);
        var reader = new ProductReader(dbContext);
        await reader.RunGeneric();
        Console.ReadKey();
    }
}
