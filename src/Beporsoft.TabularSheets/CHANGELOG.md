# Beporsoft.TabularSheets Changelog

## 1.4.0 - 2025-01-16

- Target `net9.0`

- NEW: Added extension methods of `TabularData<T>` to create markdown formatted tables.

## 1.3.0 - 2023-12-15

- Target `net8.0`.

- NEW: New element `TabularBook` which allows to export multiple sheets inside a single spreadsheet.

- NEW: New method for create on existent stream with `TabularSheet`, and for the extension method `ToCsv`.


## 1.2.2 - 2023-10-31

- FIX: `TabularSheet<T>.Options` is now `public`.
---

## 1.2.1 - 2023-10-16

### Performance improvements on style collections (Internal notes)

Improved performance on handling larger collection of styles. The collection of styles previously handled the styles as `List<TStyle>`, where the index was attached to cells. Before register a new style, it checks if exists one equivalent to the given one with `List<TStyle>.Contains`. This check is performed by iteration. However, a performance improvement has been achieved using `HashSet<TStyle>` instead. In this case, the `HashSet<TStyle>.Contains` is performed by indexing, so it is much faster.

Due to every style contains internally the index (adopted on the registration), it is not required that the collection saves the items ordered because the index is built-in on the style instance with `HashSet<TStyle>`, instead of use the index of the instance inside the previous `List<TStyle>`.

---

## 1.2.0 - 2023-09-28

### Support for export as CSV
- Extension methods `ToCsv<T>()`.
- `CsvOptions` class to configure some csv-related settings.
---

## 1.1.0 - 2023-08-07

### Column width

- Configure column width or calculate based on content.
- Added `AutoColumnWidth` and `FixedColumnWidth` to configure the column width method.
---

## 1.0.0 - 2023-07-17

### First major version
- Solved performance issues with large collections of shared strings.
- Added more style settings.
- Make private DocumentFormat.OpenXml API