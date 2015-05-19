# Exporting to Excel like a boss

This package eases the pain of exporting data from your application as an Excel file. By wrapping the [ClosedXML](https://closedxml.codeplex.com/) in an easy to use API you can also make wonderful Excel spreadsheets, just like your manager!

## Installing the bits

Soon on NuGet, hopefully, but for now you'll have to clone this repo and build it yourself.

## Using it

In order to export your data to excel you need to decide what to export and how to format it. You do this by creating a class that will represent a row of data in the excel sheet. The properties of the class will be the cells in the spreadsheet. Two "modes" are available:

- **Attribute based** - Use the `ExcelCellAttribute` to decorate the properties that will be rendered as cell values in the spreadsheet
- **POCO** - Objects of the class are [Plain Old CLR Objects](http://en.wikipedia.org/wiki/Plain_Old_CLR_Object)

Only non-inherited public properties are considered in either mode. 

### Attribute based data class

This is an example of a data class that uses attributes to specify how properties are exported to the spreadsheet cells. 

```csharp
public class MyData {
    // The First column (1), has the heading "Namn"
    [ExcelCell(Column = 1, Heading = "Namn")]
    public string Name { get; set; }

    // The second heading is a date formatted with the 'D' format specifier
    [ExcelCell(Column = 2, Heading = "Datum för händelse", Format = "D")]
    public DateTime Date { get; set; }

    // The third column is a decimal number formatted with two decimals
    [ExcelCell(Column = 3, Heading = "Lön", Format = "0.00")]
    public decimal SomeNumber { get; set; }

    // The fourth column is always treated as text even when it is numeric
    [ExcelCell(Column = 4, Heading = "Produkt-id", TreatAsText = true)]
    public string ProductId { get; set; }

    // This property has no attribte an will not be exported
    public int Index { get; set; }
}
```

### POCO class

A poco class does not use attributes (it could do, but the exporter doesn't give a damn) which results in an undeterministic order of columns, unformatted data and the headings use the de-camelized names of the properties. I.e. a property named `FooBarBaz` becomes the heading **Foo Bar Baz**.

### Do the export

The export is handled by the `ExcelExporter` class. But you probably want to use one of the `IEnumerable` extension methods to export a sequence of your carefully crafted data objects.

This example shows some pseudo-ish code that reads data from a database and exports it to Excel.

```csharp
public void DoTheExport() {
    using(var db = Db.OpenConnection()) {
        // Read data from db 
        var data = db.Query<MyData>(@"
            select ProductId, Name, Date, SomeNumber 
            from Products
            where SomeNumber <> 42");

        // Open a file stream
        using(var file = File.Create(@"c:\exports\non42products.xlsx")) {
            // Use the WriteExcel extension method to write
            // the data sequence to the file
            data.WriteExcel( file );
        }
    }
}
```

Super Simple™, amirite?!

The `IEnumerable` extension methods available are `ToExcel(bool=false)`, which takes the sequence of data objects and returns a `byte[]` containing the Excel data.
 
The other method, `WriteExcel(Stream, bool=false)` shown in the above example, writes the data sequence to the specified stream. The boolean flag indicates whether the exporter should treat the objects as POCOs or read the `ExcelCellAttribute`s of properties.

## Requirements

Built for .Net 4.0 and higher.
