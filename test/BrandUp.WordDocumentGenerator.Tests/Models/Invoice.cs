namespace BrandUp.DocxGenerator.Models
{
    internal class Invoice
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public Organization OurCompany { get; set; }
        public Organization Counterparty { get; set; }
        public InvoiceContent Content { get; set; }
    }

    internal class Organization
    {
        public string INN { get; set; }
        public string KPP { get; set; }
        public string Name { get; set; }
        public string BankName { get; set; }
        public string BankAccount { get; set; }
        public string BIC { get; set; }
        public string CorrespondentAccount { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string Link { get; set; }
    }

    internal class InvoiceContent
    {
        public List<Product> Products { get; set; }
        public int TotalSum { get; set; }
        public int Tax { get; set; }
    }
    internal class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal Amount { get; set; }
    }
}
