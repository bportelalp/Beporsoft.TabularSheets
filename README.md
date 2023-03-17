# Tabular data sheets

Tools to create data tables on file formats like Excel or Csv.

## Introduction and Goal

`Beporsoft.Spreadsheet` will allow to create simple Excel data tables based on a collection of items of type `T`.

The package uses `DocumentFormat.OpenXml` to create OpenXml Spreadsheets document. The aim of the package is to simplify the creation of a simple spreadsheet to store the information of instances of `T`, where each row represent an instance and the columns is specified using predicates.

This is one example of a simple table.

```
List<Product> productos = GetProducts();
var table = new TabularSpreadsheet<Product>();

table.SetSheetTitle("List of products");

table.SetColumn("Product", p => p.Name);
table.SetColumn("Cost", p => p.CostPerUnit);
table.SetColumn("In Stock", p => p.HasStock ? "Yes":"No");
table.SetColumn("Provider", p => p.ProviderName);

table.CreateExcel("ProductList.xlsx");
```
