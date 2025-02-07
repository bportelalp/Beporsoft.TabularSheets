# TabularSheets

[![Nuget](https://img.shields.io/nuget/v/Beporsoft.TabularSheets?logo=nuget)](https://www.nuget.org/packages/Beporsoft.TabularSheets/) 
[![Github Actions](https://github.com/bportelalp/Beporsoft.TabularSheets/actions/workflows/dotnet.yml/badge.svg)](https://github.com/bportelalp/Beporsoft.TabularSheets/actions/workflows/dotnet.yml)
[![Wiki](https://img.shields.io/badge/Docs-GitHub%20wiki-brightgreen)](https://github.com/bportelalp/Beporsoft.TabularSheets/wiki)

Create spreadsheets from your dotnet object collections quickly. With `Beporsoft.TabularSheets` your collection of `T` instances can be written on Excel Workbook easily.

## Introduction

`Beporsoft.TabularSheets` allow to create simple spreadsheet data tables based on a collection of items of type `T`.

The package uses `DocumentFormat.OpenXml` to create OpenXml Spreadsheets document. The aim of the package is to simplify the creation of a simple spreadsheet to store the information of instances of `T`, where each row represent an instance and the columns is specified using predicates.

This is one example of a simple table.

```csharp
List<Product> productList = GetProducts();

// Create table
var table = new TabularSheet<Product>(productList);
table.Title = "List of products";

// Configure columns
table.AddColumn("Product", p => p.Name);
table.AddColumn("Cost", p => p.CostPerUnit);
table.AddColumn("In Stock", p => p.HasStock ? "Yes":"No");
table.AddColumn("Provider", p => p.ProviderName);

// Add some style
table.BodyStyle.Border.SetBorderType(BorderStyle.BorderType.Thin);
table.HeaderStyle.Border.SetBorderType(BorderStyle.BorderType.Thin);
table.HeaderStyle.Fill.BackgroundColor = Color.LightCyan;

// Export document
table.CreateExcel("ProductList.xlsx");
```

That's all! Each item of your collection will be placed on a row whose cells are filled with the provided expressions.

## Documentation

-  [GitHub Wiki](https://github.com/bportelalp/Beporsoft.TabularSheets/wiki)

## Features

- Manipulate collections direclty from `TabularSheet<T>` since it implements `List<T>`.
- Handle cell content with **expression delegates**.
- Support for basic styling:
    - Apply **general** style for header and body separately.
    - Apply **specific** styling for each column.
    - **Styling features**:
      - **Fonts**: font style, color, size, bold, italic, underlined.
      - **Fill**: color
      - **Border**: color and independent style for left, right, top and bottom.
      - **Alignment**: horizontal and vertical alignment and text wrapping.
      - **Numbering**: apply direct Numbering formats. See [ECMA-376. Part 1. 18.8.30](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- Create excel books with **more than one sheet** combining multiple `TabularSheet<T>` with or without the same type.
- Additional support for **CSV** creation.
- Additional support for **Markdown Tables** creation.

## Download

`Beporsoft.TabularSheets` is available as [NuGet Package](https://www.nuget.org/packages/Beporsoft.TabularSheets/)

[![Nuget](https://img.shields.io/nuget/v/Beporsoft.TabularSheets?logo=nuget)](https://www.nuget.org/packages/Beporsoft.TabularSheets/) 

## Samples

You can clone this repo and see some use cases on [samples folder](./samples)

- [`RestCountries`](./samples/Beporsoft.TabularSheets.Samples.RestCountries/): a dotnet console application which create tables with some basic data from world countries.
  - A request to [RestCountries API](https://restcountries.com/) is performed to retrieve basic information.
  - Data is arranged on a `TabularSheet<T>`.
  - A file is created which can be an Excel Workbook, CSV or Markdown file.
- [`CurrencyExchange`](./samples/Beporsoft.TabularSheets.Samples.CurrencyExchange/): a Windows Forms application in `net48` which gets echange rates of [European Central Bank](https://www.ecb.europa.eu/stats/policy_and_exchange_rates/euro_reference_exchange_rates/html/index.en.html) using [Frankfurter API](https://frankfurter.dev/). It's a similar case like the previous one but with some UI manipulation.

## License

This software is released under [MIT License](./LICENSE).

