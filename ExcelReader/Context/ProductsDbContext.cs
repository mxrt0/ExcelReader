using ExcelReader.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelReader.Context;

public class ProductsDbContext : DbContext
{
    public ProductsDbContext(DbContextOptions options) : base(options)
    {

    }
    public DbSet<Product> Products { get; set; }
}
