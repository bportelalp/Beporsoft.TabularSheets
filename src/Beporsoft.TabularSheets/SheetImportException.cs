using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    public class SheetImportException : Exception
    {
        internal SheetImportException(string message) : base(message)
        {

        }

        internal SheetImportException(string message, Exception innerException) : base(message, innerException)
        {

        }

        internal static SheetImportException FromSheetNotFound(string sheetName)
            => new SheetImportException($"Document does not contains any sheet called {sheetName}");

        internal static SheetImportException FromPropertyNotWrittable(Type targetType, PropertyInfo targetProperty)
            => new SheetImportException($"Property {targetProperty.Name} of {targetType.Name} is not writtable");


        internal static SheetImportException FromColumnExpressionInvalid<T>(TabularDataColumn<T> column)
            => new SheetImportException($"Unable to import column \"{column.Title}\" data using expression \"{column.CellContent.Body}\". " +
                $"Only member expressions can be used to import data.");

    }
}
