using Beporsoft.TabularSheets.Builders.SheetBuilders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.Import
{
    public class SpreadsheetImporter<T> where T : class, new()
    {
        public SpreadsheetImporter(TabularSheet<T> table)
        {
            Table = table;
        }

        public TabularSheet<T> Table { get; }
        private SheetData Data { get; set; } = null!;
        private SharedStringTable SharedStrings { get; set; } = null!;
        private CellParser CellParser { get; set; } = default!;

        public void Import(string spreadsheetPath)
        {
            using var spreadsheet = SpreadsheetDocument.Open(spreadsheetPath, false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;

            Sheet sheet = GetSheetByTableName(workbookPart);

            WorksheetPart? worksheetPart = workbookPart.GetPartById(sheet.Id?.Value!) as WorksheetPart;

            Worksheet worksheet = worksheetPart?.Worksheet!;

            Data = worksheet.Descendants<SheetData>()!.Single();
            SharedStrings = workbookPart.SharedStringTablePart!.SharedStringTable;
            CellParser = new CellParser(SharedStrings);

            Dictionary<TabularDataColumn<T>, int> filledColumns = GetFillableColumns();
            IEnumerable<Row> bodyRows = GetBodyRows();
            List<T> values = [];
            foreach (Row? row in bodyRows)
            {
                T rowValue = new();
                foreach (var columnRelation in filledColumns)
                {
                    Cell cell = row.Descendants<Cell>().ElementAt(columnRelation.Value);
                    PopulateWithCellValue(rowValue, columnRelation.Key, cell);
                }
                values.Add(rowValue);
            }
            Table.AddRange(values);
        }

        private Sheet GetSheetByTableName(WorkbookPart workbookPart)
        {
            string sheetName = Table.Title;
            Sheet? sheet = (workbookPart.Workbook.Sheets?
                .Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == Table.Title)) ?? throw SheetImportException.FromSheetNotFound(sheetName);
            return sheet;
        }

        private Dictionary<TabularDataColumn<T>, int> GetFillableColumns()
        {
            Dictionary<TabularDataColumn<T>, int> columns = [];
            Row headerRow = GetHeaderRow();
            foreach (Cell cell in headerRow.Descendants<Cell>())
            {
                (int Row, int Col) cellRef = CellRefBuilder.GetIndexes(cell.CellReference!);
                string? columnName = (string?)CellParser.GetValue(cell, typeof(string));

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

        private void PopulateWithCellValue(T instance, TabularDataColumn<T> column, Cell cell)
        {
            Expression expr = column.CellContent.Body;
            if (expr is UnaryExpression unaryExpresion && unaryExpresion.NodeType == ExpressionType.Convert)
            {
                expr = unaryExpresion.Operand;
            }
            else if (expr is MemberExpression memberExpression)
            {
                expr = memberExpression;
            }
            var memberSelector = expr as MemberExpression;
            if (memberSelector is not null)
            {
                PropertyInfo? property = memberSelector.Member as PropertyInfo;

                if (property is not null)
                {
                    if (!property.CanWrite)
                        throw SheetImportException.FromPropertyNotWrittable(typeof(T), property);
                    object? value = CellParser.GetValue(cell, property.PropertyType);
                    property.SetValue(instance, value);
                }
            }
            else
            {
                throw SheetImportException.FromColumnExpressionInvalid(column);
            }
        }
    }
}
