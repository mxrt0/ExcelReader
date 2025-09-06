using ExcelReader.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelReader.Context;

public class GenericDbContext : DbContext
{
    private readonly IEnumerable<string> _headers;

    public GenericDbContext(IEnumerable<string> headers)
    {
        _headers = headers;
    }

    public DbSet<ExcelRow> Items { get; set; }

    public IEnumerable<string> GetHeaders() => _headers;

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlite(@"Data Source=..\..\..\products.db");
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        var entity = modelBuilder.Entity<ExcelRow>();

        foreach (var header in _headers)
        {
            entity.Property<string>(header);
        }
    }

}
