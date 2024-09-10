namespace Zeyt.ExcelDocument.ConsoleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var customerList = new List<Customer>
            {
                new Customer { FirstName = "James", LastName = "Butt", Age = 23, Email = "james@james.com" },
                new Customer { FirstName = "Art", LastName = "Venere", Age = 33, Email = "art@venere.com" },
                new Customer { FirstName = "Sage", LastName = "Wieser", Age = 57, Email = null },
                new Customer { FirstName = "Minna", LastName = "Amigon", Age = 45, Email = "minna@amigon.com" },
                new Customer { FirstName = "Blair", LastName = "Malet", Age = 66, Email = null },
            };

            var customerExcelData = new ExcelDocumentWriter<Customer, CustomerExcelMap>("Customer Information").Write(customerList);
            File.WriteAllBytes(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customer.xlsx"), customerExcelData);
        }
    }

    public class CustomerExcelMap : ExcelDocumentMap<Customer>
    {
        public CustomerExcelMap()
        {
            base.Map(x => x.FirstName).Name("Full Name").Width(30).WriteUsing(x => $"{x.FirstName} {x.LastName}");
            base.Map(x => x.Age).Name("Age").Width(10);
            base.Map(x => x.Email).Name("Email").Width(50).Default("EMPTY");
        }
    }

    public class Customer
    {
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public int? Age { get; set; } = 0;
        public string? Email { get; set; }
    }
}