# Tabular data sheets


Tools to create data tables on file formats like Excel or Csv.

## Introduction

`Beporsoft.TabularSheets` will allow to create simple Excel data tables based on a collection of items of type `T`.

The package uses `DocumentFormat.OpenXml` to create OpenXml Spreadsheets document. The aim of the package is to simplify the creation of a simple spreadsheet to store the information of instances of `T`, where each row represent an instance and the columns is specified using predicates.

This is one example of a simple table.

```
List<Product> productos = GetProducts();
var table = new TabularSpreadsheet<Product>();

table.SetSheetTitle("List of products");

table.AddColumn("Product", p => p.Name);
table.AddColumn("Cost", p => p.CostPerUnit);
table.AddColumn("In Stock", p => p.HasStock ? "Yes":"No");
table.AddColumn("Provider", p => p.ProviderName);

table.CreateExcel("ProductList.xlsx");
```
