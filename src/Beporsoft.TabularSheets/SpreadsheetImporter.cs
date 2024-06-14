using Beporsoft.TabularSheets.Builders.SheetBuilders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    public class SpreadsheetImporter<T> where T : class, new()
    {
        public SpreadsheetImporter(TabularSheet<T> table)
        {
            Table = table;
        }

        public TabularSheet<T> Table { get; }
        public SheetData Data { get; private set; } = null!;
        private SharedStringTable SharedStrings { get; set; } = null!;

        public void ImportContent(string spreadsheetPath)
        {
            using var fs = new FileStream(spreadsheetPath, FileMode.Open, FileAccess.Read);

            using var spreadsheet = SpreadsheetDocument.Open(fs, false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;

            Sheet sheet = GetSheetByTableName(workbookPart);

            WorksheetPart? worksheetPart = workbookPart.GetPartById(sheet.Id?.Value!) as WorksheetPart;

            Worksheet worksheet = worksheetPart?.Worksheet!;

            Data = worksheet.Descendants<SheetData>()!.Single();
            SharedStrings = workbookPart.SharedStringTablePart!.SharedStringTable;

            Dictionary<TabularDataColumn<T>, int> filledColumns = GetFillableColumns();
            IEnumerable<Row> bodyRows = GetBodyRows();
            List<T> values = [];
            foreach (Row? row in bodyRows)
            {
                T rowValue = new();
                foreach (var columnRelation in filledColumns)
                {
                    
                    Cell cell = row.Descendants<Cell>().ElementAt(columnRelation.Value);

                    
                }
            }
        }

        

        private Sheet GetSheetByTableName(WorkbookPart workbookPart)
        {
            string sheetName = Table.Title;
            Sheet? sheet = (workbookPart.Workbook.Sheets?
                .Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == Table.Title)) ?? throw new SheetImportException($"Document does not contains any sheet called {sheetName}");

            return sheet;
        }

        private Dictionary<TabularDataColumn<T>, int> GetFillableColumns()
        {
            Dictionary<TabularDataColumn<T>, int> columns = [];
            Row headerRow = GetHeaderRow();
            foreach (Cell cell in headerRow.Descendants<Cell>())
            {
                (int Row, int Col) cellRef = CellRefBuilder.GetIndexes(cell.CellReference!);
                string columnName;
                if (cell.DataType!.Value == CellValues.SharedString)
                    columnName = GetSharedString(int.Parse(cell.CellValue!.InnerText));
                else
                    columnName = cell.CellValue!.InnerText;

                List<TabularDataColumn<T>> namedColumns = Table.Columns.Where(c => c.Title == columnName).ToList();
                if (namedColumns.Count == 0)
                    continue;
                else if (namedColumns.Count == 1)
                    columns[namedColumns.Single()] = cellRef.Col;
                else //TODO varias columnas mismo nombre¿?
                    throw new SheetImportException("TODO varias columnas mismo nombre");
            }
            return columns;
        }

        private Row GetHeaderRow()
        {
            //TODO Configurar header row.
            Row firstRow = Data.Descendants<Row>().FirstOrDefault();
            return firstRow;
        }

        private IEnumerable<Row> GetBodyRows()
        {
            // TODO seleccionar filas.
            IEnumerable<Row> bodyRows = Data.Descendants<Row>().Skip(1);
            return bodyRows;
        }

        private string GetSharedString(int index)
        {
            string value = SharedStrings.ElementAt(index).InnerText;
            return value;
        }

        private void FillCellValue(T instance, TabularDataColumn<T> column, object value)
        {
            var memberSelectorExpression = column.CellContent.Body as MemberExpression;
            var property = memberSelectorExpression.Member as PropertyInfo;
            property.SetValue(instance, value);
        }
    }
}
