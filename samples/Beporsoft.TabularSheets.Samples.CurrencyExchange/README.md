# Sample: Currency exchanges.

This sample uses [Frankfurter API](https://www.frankfurter.app/) to retrieve the exchange between currencies, including time series.

The example shows a Windows Forms App written on .NET Framework 4.8 `.net48` and demonstrate the use of `TabularSheet` to export the time serie of conversion betweeen currencies along time

## Model
The class `Exchange` models the conversion between one origin currency to target currency. The class `ExchangeRecord` models the exchange between one origin currency to multiple target currencies at a specific date.

## How it works.

By means of the *Frankfurter API* it is retrieved a `List<ExchangeRecord>`. Notice that `ExchangeRecord` contains inside another list.
The sample demostrate how to make a table that each row is one item of `List<ExchangeRecord>` but columns has multiple origins:

* Properties from `ExchangeRecord` like `Date`.
* Properties from the nested list `ExchangeRecord.Exchanges`. It is possible to combine a `TabularSheet.AddColumn()` call inside a `foreach` with `Linq` queries.