**Beporsoft.TabularSheets** makes it easy to export generic collections into spreadsheets. It provides a simple API for specifying the contents of the object to be displayed in each column.

A common case of use is the export of one `ICollection<T>` where each row represents a `T` object and each column includes some of the properties of `T` or combinations between them.

The package uses `DocumentFormat.OpenXml` SDK to create MS Excel Documents. 