using Beporsoft.TabularSheet.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Beporsoft.TabularSheet.Spreadsheets;

namespace Beporsoft.TabularSheet
{
    public class TabularSpreadsheet<T> : TabularData<T>
    {
        private const string _defaultExtension = ".xlsx";

        public TabularSpreadsheet()
        {
        }

        public TabularSpreadsheet(IEnumerable<T> items)
        {
            Items = items.ToList();
        }

        public string Title { get; set; } = "Sheet";
        public HeaderOptions Header { get; set; } = new();


        #region Configure Table
        public void SetSheetTitle(string title) => Title = title;
        #endregion

        #region Create
        public void Create(string path)
        {
            string pathCorrected = FileHelpers.VerifyPath(path, _defaultExtension);
            using (var fs = new FileStream(pathCorrected, FileMode.Create))
            using (MemoryStream ms = Create())
            {
                ms.Seek(0, SeekOrigin.Begin);
                ms.CopyTo(fs);
            }
        }
        public MemoryStream Create()
        {
            var ms = new MemoryStream();
            using (var spreadSheet = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadSheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                AppendWorksheetPart(ref workbookPart);
                workbookPart.Workbook.Save();
                spreadSheet.Close();
            }
            return ms;
        }
        #endregion

        #region Build Sheet

        /// <summary>
        /// This is for append more than one sheet inside a Worbook. Externally we can pass a WorkbookPart and this method
        /// will add the respective sheet which represent <see langword="this"/>.
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetId"></param>
        private void AppendWorksheetPart(ref WorkbookPart workbookPart)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            //Add Sheets to the Workbook if there aren't
            Sheets sheets;
            if (workbookPart.Workbook.Sheets is null)
                sheets = workbookPart!.Workbook.AppendChild(new Sheets());
            else
                sheets = workbookPart.Workbook.Sheets;

            UInt32Value sheetIdValue = FindSuitableSheetId(sheets);
            string nameSheet = BuildSuitableSheetName(sheets);

            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetIdValue,
                Name = nameSheet
            };
            sheets.Append(sheet);

            SheetData sheetData = BuildSheetData();
            worksheetPart.Worksheet.AppendChild(sheetData);
        }

        private SheetData BuildSheetData()
        {
            var sheetData = new SheetData();
            //Header
            var row = new Row();
            foreach (var column in _columns)
                row.Append(CreateCell(column.Title, CellValues.String));
            sheetData.AppendChild(row);

            //Rows
            foreach (var item in Items)
            {
                row = new Row();
                foreach (var column in _columns)
                {
                    object value = column.Apply(item);
                    if (value is null)
                        row.Append(new Cell());
                    else
                        row.Append(CreateCell(value));
                }
                sheetData.AppendChild(row);
            }
            return sheetData;
        }

        /// <summary>
        /// Automatic find a SheetId based on current sheets ids, creating a new incremental value
        /// </summary>
        /// <param name="sheets"></param>
        /// <returns></returns>
        private static UInt32Value FindSuitableSheetId(Sheets sheets)
        {
            UInt32Value? lastId = sheets.Select(s => s as Sheet).Max(s => s?.SheetId);
            UInt32Value sheetIdValue = lastId is null ? 1 : lastId + 1;
            return sheetIdValue;
        }

        /// <summary>
        /// Automatic find a suitable name for sheet. If there is a sheet with the same name, look for a suitable
        /// name based on {name}{incremental}
        /// </summary>
        /// <param name="sheets"></param>
        /// <returns></returns>
        private string BuildSuitableSheetName(Sheets sheets)
        {
            string nameSheet = string.IsNullOrWhiteSpace(Title) ? typeof(T).Name : Title;
            if (sheets.Select(s => s as Sheet).Any(s => s?.Name == nameSheet))
            {
                // Look for the last numeric value
                IEnumerable<StringValue?> sameNamesStarted = sheets
                    .Select(s => s as Sheet)
                    .Select(s => s!.Name)
                    .Where(n => n!.Value!.StartsWith(nameSheet))
                    .OrderBy(s => s);
                StringValue? highestValue = sameNamesStarted.LastOrDefault();

                Regex regex = new(@"\d{1,}$");
                Match matches = regex.Match(highestValue!);
                if (matches.Success)
                    nameSheet += Convert.ToInt32(matches.Value) + 1;
                else
                    nameSheet += "1";
            }
            return nameSheet;
        }
        #endregion

        #region Data handling
        protected static Cell CreateCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        protected static Cell CreateCell(object value)
        {
            CellValue cellContent;
            Type type = value.GetType();
            CellValues dataType;
            if (type == typeof(string))
            {
                dataType = CellValues.String;
                cellContent = new CellValue(Convert.ToString(value) ?? string.Empty);
            }
            else if (type == typeof(int))
            {
                dataType = CellValues.Number;
                cellContent = new CellValue(Convert.ToInt32(value));
            }
            else if (type == typeof(double))
            {
                dataType = CellValues.Number;
                cellContent = new CellValue(Convert.ToDouble(value));
            }
            else if (type == typeof(float))
            {
                dataType = CellValues.Number;
                cellContent = new CellValue(Convert.ToDecimal(value));
            }
            else if (type == typeof(bool))
            {
                dataType = CellValues.Boolean;
                cellContent = new CellValue(Convert.ToBoolean(value));
            }
            else if (type == typeof(DateTime))
            {
                dataType = CellValues.Date;
                cellContent = new CellValue(Convert.ToString(value) ?? string.Empty);
            }
            else
            {
                dataType = CellValues.String;
                cellContent = new CellValue(Convert.ToString(value) ?? string.Empty);
            }

            return new Cell()
            {
                CellValue = cellContent,
                DataType = dataType
            };
        }
        #endregion
    }
}
