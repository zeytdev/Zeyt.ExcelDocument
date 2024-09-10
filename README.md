# Zeyt.ExcelDocument

**Zeyt.ExcelDocument** is a powerful and easy-to-use C# library designed for generating Excel documents from object lists. The library allows you to map object properties to Excel columns with custom names, widths, and even default values for fields with missing data. It also provides an easy way to export this data to `.xlsx` format.

## Features

- **Customizable Mappings**: Easily map object properties to Excel columns.
- **Fluent API**: Set column names, widths, and default values for null fields.
- **Quick Export**: Convert object lists into Excel files in just a few lines of code.
- **Null Field Handling**: Define default values for fields with missing data.

## Installation

You can install the package via NuGet Package Manager:

```bash
Install-Package Zeyt.ExcelDocument
```

Or via .NET CLI:

```bash
dotnet add package Zeyt.ExcelDocument
```

## Example Usage

Here is a basic example of how to use the **Zeyt.ExcelDocument** library:

```csharp
using Zeyt.ExcelDocument;

var customerList = new List<Customer>
{
    new Customer { FirstName = "James", LastName = "Butt", Age = 23, Email = "james@james.com" },
    new Customer { FirstName = "Art", LastName = "Venere", Age = 33, Email = "art@venere.com" },
    new Customer { FirstName = "Sage", LastName = "Wieser", Age = 57, Email = null },
    new Customer { FirstName = "Minna", LastName = "Amigon", Age = 45, Email = "minna@amigon.com" },
    new Customer { FirstName = "Blair", LastName = "Malet", Age = 66, Email = null },
};

// Map customer data to Excel and export it to a file
var customerExcelData = new ExcelDocumentWriter<Customer, CustomerExcelMap>("Customer Information").Write(customerList);
File.WriteAllBytes(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customer.xlsx"), customerExcelData);
```

### Mapping Configuration

The mapping is done using a separate `ExcelDocumentMap<T>` class. In this example, we create a custom mapping for the `Customer` class:

```csharp
public class CustomerExcelMap : ExcelDocumentMap<Customer>
{
    public CustomerExcelMap()
    {
        Map(x => x.FirstName).Name("Full Name").Width(30).WriteUsing(x => $"{x.FirstName} {x.LastName}");
        Map(x => x.Age).Name("Age").Width(10);
        Map(x => x.Email).Name("Email").Width(50).Default("EMPTY");
    }
}
```

### The `Customer` Class

This is the class we are mapping to Excel columns:

```csharp
public class Customer
{
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public int? Age { get; set; } = 0;
    public string? Email { get; set; }
}
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---