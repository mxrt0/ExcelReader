global using Spectre.Console;
using ExcelReader.Services;
using OfficeOpenXml;
namespace ExcelReader;

public class Program
{
    static async Task Main()
    {
        ExcelPackage.License.SetNonCommercialPersonal("mxrt0");
        var filePath = AnsiConsole.Prompt(new TextPrompt<string>("[aqua]Please insert the file path you wish to seed into/read from: [/]"));
        var reader = new ProductReader(filePath.Trim('"') ?? @"..\..\..\data.xslx");
        await reader.Run();
        Console.ReadKey();
    }
}
