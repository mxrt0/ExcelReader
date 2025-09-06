using ExcelReader.Context;
using OfficeOpenXml;
namespace ExcelReader.Services;

public class ProductReader
{
    private GenericDbContext? _genericContext;
    private string excelFilePath;
    public ProductReader(string filePath)
    {
        excelFilePath = filePath;
    }

    public async Task Run()
    {
        var data = ReadDataFromSpreadsheet(excelFilePath);
        await SaveDataToDatabase(data);
        DisplayData(_genericContext!.Items);
    }

    public IEnumerable<Models.ExcelRow> ReadDataFromSpreadsheet(string filePath)
    {
        AnsiConsole.MarkupLine("[aqua]\nOpening [lime][bold]Excel[/][/] worksheet...[/]\n");
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

        AnsiConsole.MarkupLine("[aqua]\nCreating [lime][bold]Database Table[/][/]...[/]\n");
        _genericContext = new GenericDbContext(headers);
        _genericContext.Database.EnsureDeleted();
        _genericContext.Database.EnsureCreated();
        AnsiConsole.MarkupLine("[aqua]Created [lime][bold]Database Table[/][/]![/]");

        var rows = new List<Models.ExcelRow>();
        AnsiConsole.Progress()
            .Start(ctx =>
            {
                var task = ctx.AddTask("[aqua]Reading Data From [lime][bold]Excel[/][/] Worksheet[/]", maxValue: rowCount - 1);

                for (int row = 2; row <= rowCount; row++)
                {
                    var entity = new Models.ExcelRow();

                    var entry = _genericContext.Entry(entity);
                    for (int col = 1; col <= colCount; col++)
                    {
                        var header = headers[col - 1];
                        var colValue = worksheet.Cells[row, col].Text.Trim() ?? "<>";
                        entry.Property(header).CurrentValue = colValue;
                    }
                    task.Increment(1);
                    rows.Add(entity);
                }
            });
        return rows;
    }

    public async Task SaveDataToDatabase(IEnumerable<Models.ExcelRow> items)
    {
        AnsiConsole.MarkupLine("[aqua]Saving [lime][bold]Data[/][/] To Database...[/]\n");
        foreach (var item in items)
        {
            _genericContext!.Items.Add(item);
        }
        await _genericContext!.SaveChangesAsync();
        AnsiConsole.MarkupLine("[aqua]Saved [lime][bold]Data[/][/] To Database!\n\n[/]");
    }

    public void DisplayData(IEnumerable<Models.ExcelRow> items)
    {
        var table = new Table().RoundedBorder().Centered().ShowRowSeparators();
        var headers = _genericContext!.GetHeaders();
        foreach (string header in headers)
        {
            table.AddColumn(header).Centered();
        }
        var entities = items.Select(e => _genericContext!.Entry(e)).Where(e =>
        {
            foreach (var header in headers)
            {
                if (e.Property(header).CurrentValue is null || string.IsNullOrEmpty(e.Property(header)?.CurrentValue?.ToString()))
                {
                    return false;
                }
            }
            return true;
        });
        foreach (var entity in entities)
        {
            string[] rowData = new string[headers.Count()];
            int index = 0;
            foreach (string header in headers)
            {
                rowData[index++] = entity.Property(header).CurrentValue!.ToString()!;
            }
            table.AddRow(rowData).Centered();
        }
        AnsiConsole.MarkupLine($"[aqua]Data:[/]\n");
        AnsiConsole.Write(table);
    }
}
