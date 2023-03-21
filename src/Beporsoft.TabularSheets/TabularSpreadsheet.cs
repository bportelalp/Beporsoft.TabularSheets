using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Beporsoft.TabularSheets.Tools;
using Beporsoft.TabularSheets.Style;
using DocumentFormat.OpenXml.Validation;
using System.Xml;
using Beporsoft.TabularSheets.Builders.StyleBuilders;

namespace Beporsoft.TabularSheets
{
    public class TabularSpreadsheet<T> : TabularData<T>
    {
        private static readonly string[] _defaultExtensions = new string[] { ".xlsx", ".xls" };
        private readonly StylesheetBuilder _styleBuilder = new();
        public TabularSpreadsheet()
        {
        }

        public TabularSpreadsheet(IEnumerable<T> items)
        {
            Items = items.ToList();
        }

        /// <summary>
        /// The title of the current sheet
        /// </summary>
        public string Title { get; set; } = "Sheet";
        public HeaderOptions Header { get; set; } = new();


        #region Configure Table
        public void SetSheetTitle(string title) => Title = title;
        #endregion

        #region Create
        public void Create(string path)
        {
            string pathCorrected = FileHelpers.VerifyPath(path, _defaultExtensions);
            using var fs = new FileStream(pathCorrected, FileMode.Create);
            using MemoryStream ms = Create();
            ms.Seek(0, SeekOrigin.Begin);
            ms.CopyTo(fs);
        }
        public MemoryStream Create()
        {
            var ms = new MemoryStream();
            using var spreadSheet = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            FillSpreadsheetData(spreadSheet);
            return ms;
        }
        #endregion

        #region Build Sheet

        /// <summary>
        /// Fill the given document with the respective parts of OpenXML Sheets
        /// </summary>
        /// <param name="spreadsheet"></param>
        private void FillSpreadsheetData(SpreadsheetDocument spreadsheet)
        {
            WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
            BuildStyleSheet(ref workbookPart);
            workbookPart.Workbook = new Workbook();
            AppendWorksheetPart(ref workbookPart);
            workbookPart.Workbook.Save();
            ValidateSpreadSheet(spreadsheet);
            spreadsheet.Close();
        }

        /// <summary>
        /// This is for append more than one sheet inside a Worbook. Externally we can pass a WorkbookPart and this method
        /// will add the respective sheet which represent <see langword="this"/>.
        /// </summary>
        /// <param name="workbookPart">The workbookPart where it will be added a new worksheetPart</param>
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

        /// <summary>
        /// Build the <see cref="SheetData"/> which contains the data represented by <see langword="this"/>
        /// </summary>
        /// <returns></returns>
        private SheetData BuildSheetData()
        {
            var sheetData = new SheetData();
            //Header
            var row = new Row();
            foreach (var column in _columns)
            {
                var cell = CreateCell(column.Title, CellValues.String);
                cell.StyleIndex = 0;
                row.Append(cell);
            }
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

        #region Validation
        private static void ValidateSpreadSheet(SpreadsheetDocument spreadsheet)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            IEnumerable<ValidationErrorInfo> errors = validator.Validate(spreadsheet);
            if (errors.Any())
                throw new XmlException("Errors validating Xml");
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
            Type type = value.GetType();
            Cell cell = new Cell();
            if (type == typeof(string))
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(Convert.ToString(value) ?? string.Empty);
            }
            else if (type == typeof(int))
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToInt32(value));
            }
            else if (type == typeof(double))
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToDouble(value));
            }
            else if (type == typeof(float))
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToDecimal(value));
            }
            else if (type == typeof(bool))
            {
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(Convert.ToBoolean(value));
            }
            else if (type == typeof(DateTime))
            {
                cell.DataType = CellValues.Number;
                cell.StyleIndex = 1;
                cell.CellValue = new CellValue(Convert.ToDateTime(value).ToOADate().ToString());
            }
            else
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(Convert.ToString(value) ?? string.Empty);
            }

            return cell;
        }
        #endregion

        #region Style Handling
        private void BuildStyleSheet(ref WorkbookPart workbookPart)
        {
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet()
            {
                
                Fonts = new Fonts(new Font()),
                Fills = new Fills(),
                Borders = new Borders(new Border()),
                CellStyleFormats = new CellStyleFormats(new CellFormat()),
                CellFormats =
                new CellFormats(
                    new CellFormat()
                    {
                        FillId = 0,
                        ApplyFill = true,
                    },
                    new CellFormat
                    {
                        NumberFormatId = 14,
                        FillId = 0,
                        ApplyNumberFormat = true
                    })
            };
        }
        #endregion
    }
}
