using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Beporsoft.TabularSheets.Builders
{
    /// <summary>
    /// A class which build a spreadsheets from instance(s) of <see cref="TabularSheet{T}"/>
    /// </summary>
    internal sealed class SpreadsheetBuilder
    {
        private StylesheetBuilder _styleBuilder;
        private SharedStringBuilder _sharedStringBuilder;

        /// <summary>
        /// Build spreadsheets.
        /// </summary>
        public SpreadsheetBuilder()
        {
            _styleBuilder = new StylesheetBuilder();
            _sharedStringBuilder = new SharedStringBuilder();
        }

        /// <summary>
        /// Build spreadsheets using a shared <see cref="StylesheetBuilder"/>. This is ideal when build spreadsheets with more than one table.
        /// </summary>
        /// <param name="stylesheetBuilder"></param>
        /// <param name="sharedStringBuilder"></param>
        public SpreadsheetBuilder(StylesheetBuilder stylesheetBuilder, SharedStringBuilder sharedStringBuilder)
        {

            _styleBuilder = stylesheetBuilder;
            _sharedStringBuilder = sharedStringBuilder;
        }

        public StylesheetBuilder StyleBuilder => _styleBuilder;
        public SharedStringBuilder SharedStringBuilder => _sharedStringBuilder;

        #region Create Spreadsheet
        public void Create(string path, params ISheet[] tables)
        {
            string pathCorrected = FileHelpers.VerifyPath(path, SpreadsheetsFileExtension.AllowedExtensions);
            using var fs = new FileStream(pathCorrected, FileMode.Create);
            using MemoryStream ms = Create(tables);
            ms.Seek(0, SeekOrigin.Begin);
            ms.CopyTo(fs);
        }
        public MemoryStream Create(params ISheet[] tables)
        {
            MemoryStream stream = new();
            using var spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            foreach (var table in tables)
            {
                AppendWorksheetPart(ref workbookPart, table);
            }
            AppendWorkbookStylePart(ref workbookPart);
            AppendSharedStringTablePart(ref workbookPart);
            workbookPart.Workbook.Save();
            ValidateSpreadSheet(spreadsheet);
            return stream;
        }
        #endregion

        #region Append OpenXml Parts

        /// <summary>
        /// Append a <see cref="WorksheetPart"/> to <paramref name="workbookPart"/> based on the
        /// content of the provided <see cref="ISheet"/>.<br/><br/>
        /// The <see cref="SharedStringItem"/>s located and the <see cref="CellFormat"/> of each cell when it is required will be 
        /// registered on <see cref="SharedStringBuilder"/> and <see cref="StyleBuilder"/> to include them subsequently inside a 
        /// common <see cref="SharedStringTable"/> and <see cref="Stylesheet"/>, respectively.<br/><br/>
        /// This architecture allows to include more than one <see cref="ISheet"/> on the same <see cref="SpreadsheetDocument"/> 
        /// using shared resources.
        /// </summary>
        /// <param name="workbookPart">A reference to the <see cref="WorkbookPart"/> where to append the <see cref="WorkbookStylesPart"/></param>
        /// <param name="table">The <see cref="TabularSheet{T}"/> which will populate the <see cref="SheetData"/></param>
        public void AppendWorksheetPart(ref WorkbookPart workbookPart, ISheet table)
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
            string nameSheet = string.IsNullOrWhiteSpace(table.Title) ? table.ItemType.Name : table.Title;
            nameSheet = BuildSuitableSheetName(sheets, nameSheet);

            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetIdValue,
                Name = nameSheet
            };
            sheets.Append(sheet);

            Columns columns = table.BuildColumns();
            worksheetPart.Worksheet.Append(columns);

            SheetData sheetData = table.BuildSheetContext(_styleBuilder, _sharedStringBuilder);
            worksheetPart.Worksheet.AppendChild(sheetData);
        }


        /// <summary>
        /// Append a  <see cref="WorkbookStylesPart"/> to <paramref name="workbookPart"/> based on the
        /// content of <see cref="StyleBuilder"/>
        /// </summary>
        /// <param name="workbookPart">A reference to the <see cref="WorkbookPart"/> where to append the <see cref="WorkbookStylesPart"/></param>
        private void AppendWorkbookStylePart(ref WorkbookPart workbookPart)
        {
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var stylesheet = new Stylesheet();
            stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat());

            // Protect the SpreadsheetML Schema adding containers only if there is any registered item
            if (StyleBuilder.RegisteredFills > 0)
                stylesheet.Fills = StyleBuilder.GetFills();
            if (StyleBuilder.RegisteredFonts > 0)
                stylesheet.Fonts = StyleBuilder.GetFonts();
            if (StyleBuilder.RegisteredBorders > 0)
                stylesheet.Borders = StyleBuilder.GetBorders();
            if (StyleBuilder.RegisteredNumberingFormats > 0)
                stylesheet.NumberingFormats = StyleBuilder.GetNumberingFormats();
            if (StyleBuilder.RegisteredFormats > 0)
                stylesheet.CellFormats = StyleBuilder.GetFormats();

            stylesPart.Stylesheet = stylesheet;
        }

        /// <summary>
        /// Append a <see cref="SharedStringTablePart"/> to <paramref name="workbookPart"/> based on the content
        /// of <see cref="SharedStringBuilder"/>
        /// </summary>
        /// <param name="workbookPart">A reference to the <see cref="WorkbookPart"/> where to append the <see cref="SharedStringTablePart"/></param>
        private void AppendSharedStringTablePart(ref WorkbookPart workbookPart)
        {
            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            var sharedStringTable = _sharedStringBuilder.GetSharedStringTable();
            sharedStringTablePart.SharedStringTable = sharedStringTable;
        }

        #endregion

        #region Helpers for Sheet Naming
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
        /// <param name="nameSheet"></param>
        /// <returns></returns>
        private string BuildSuitableSheetName(Sheets sheets, string nameSheet)
        {
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

        #region OpenXml Validation
        private static void ValidateSpreadSheet(SpreadsheetDocument spreadsheet)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            IEnumerable<ValidationErrorInfo> errors = validator.Validate(spreadsheet);
            if (errors.Any())
                throw new XmlException("Errors validating Xml");
        }
        #endregion
    }
}
