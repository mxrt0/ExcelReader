namespace ExcelReader.Models;

public class Product
{
    public Product(string productName, decimal unitPrice, int quantity, decimal totalPrice, decimal totalPriceWithVAT)
    {
        ProductName = productName;
        UnitPrice = unitPrice;
        Quantity = quantity;
        TotalPrice = totalPrice;
        TotalPriceWithVAT = totalPriceWithVAT;
    }

    public string ProductName { get; set; }
    public decimal UnitPrice { get; set; }
    public int Quantity { get; set; }
    public decimal TotalPrice { get; set; }
    public decimal TotalPriceWithVAT { get; set; }
}
