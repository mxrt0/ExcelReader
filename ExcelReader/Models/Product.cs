namespace ExcelReader.Models;

public class Product
{
    public Product(string productName, decimal unitPrice, int quantity, string random)
    {
        ProductName = productName;
        UnitPrice = unitPrice;
        Quantity = quantity;
        TotalPrice = UnitPrice * Quantity;
        TotalPriceWithVAT = TotalPrice * 1.2m;
        Random = random;
    }

    public Product(string productName, decimal unitPrice, int quantity, decimal totalPrice, decimal totalPriceWithVAT)
    {
        ProductName = productName;
        UnitPrice = unitPrice;
        Quantity = quantity;
        TotalPrice = totalPrice;
        TotalPriceWithVAT = totalPriceWithVAT;
    }

    public int Id { get; set; }
    public string ProductName { get; set; }
    public decimal UnitPrice { get; set; }
    public int Quantity { get; set; }
    public decimal TotalPrice { get; set; }
    public decimal TotalPriceWithVAT { get; set; }

    public string Random { get; set; }
}
