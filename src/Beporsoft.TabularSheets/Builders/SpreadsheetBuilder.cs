﻿using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.Exceptions;
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

namespace Beporsoft.TabularSheets.Builders
{
    /// <summary>
    /// A class which build a spreadsheets from instance(s) of <see cref="TabularSheet{T}"/>
    /// </summary>
    internal sealed class SpreadsheetBuilder
    {

        /// <summary>
        /// Build spreadsheets.
        /// </summary>
        public SpreadsheetBuilder()
        {
            StyleBuilder = new StylesheetBuilder();
            SharedStringBuilder = new SharedStringBuilder();
        }

        /// <summary>
        /// Build spreadsheets using a shared <see cref="StylesheetBuilder"/>. This is ideal when build spreadsheets with more than one table.
        /// </summary>
        /// <param name="stylesheetBuilder"></param>
        /// <param name="sharedStringBuilder"></param>
        public SpreadsheetBuilder(StylesheetBuilder stylesheetBuilder, SharedStringBuilder sharedStringBuilder)
        {

            StyleBuilder = stylesheetBuilder;
            SharedStringBuilder = sharedStringBuilder;
        }

        public StylesheetBuilder StyleBuilder { get; }
        public SharedStringBuilder SharedStringBuilder { get; }

        #region Create Spreadsheet
        public void Create(string path, params ITabularSheet[] tables)
        {
            string pathCorrected = FileHelpers.VerifyPath(path, SpreadsheetFileExtension.AllowedExtensions);
            using var fs = new FileStream(pathCorrected, FileMode.Create);
            using MemoryStream ms = Create(tables);
            ms.Seek(0, SeekOrigin.Begin);
            ms.CopyTo(fs);
        }

        public MemoryStream Create(params ITabularSheet[] tables)
        {
            MemoryStream stream = new();
            Create(stream, tables);
            return stream;
        }

        public void Create(Stream stream, params ITabularSheet[] tables)
        {
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
        }
        #endregion

        #region Append OpenXml Parts

        /// <summary>
        /// Append a <see cref="WorksheetPart"/> to <paramref name="workbookPart"/> based on the
        /// content of the provided <see cref="ITabularSheet"/>.<br/><br/>
        /// The <see cref="SharedStringItem"/>s located and the <see cref="CellFormat"/> of each cell when it is required will be 
        /// registered on <see cref="SharedStringBuilder"/> and <see cref="StyleBuilder"/> to include them subsequently inside a 
        /// common <see cref="SharedStringTable"/> and <see cref="Stylesheet"/>, respectively.<br/><br/>
        /// This architecture allows to include more than one <see cref="ITabularSheet"/> on the same <see cref="SpreadsheetDocument"/> 
        /// using shared resources.
        /// </summary>
        /// <param name="workbookPart">A reference to the <see cref="WorkbookPart"/> where to append the <see cref="WorkbookStylesPart"/></param>
        /// <param name="table">The <see cref="TabularSheet{T}"/> which will populate the <see cref="SheetData"/></param>
        public void AppendWorksheetPart(ref WorkbookPart workbookPart, ITabularSheet table)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            Sheets sheets;
            if (workbookPart.Workbook.Sheets is null)
                sheets = workbookPart!.Workbook.AppendChild(new Sheets());
            else
                sheets = workbookPart.Workbook.Sheets;

            UInt32Value sheetIdValue = FindSuitableSheetId(sheets);
            string nameSheet = string.IsNullOrWhiteSpace(table.Title) ? table.RowType.Name : table.Title;
            nameSheet = BuildSuitableSheetName(sheets, nameSheet);

            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetIdValue,
                Name = nameSheet
            };
            sheets.Append(sheet);

            Worksheet ws = table.BuildWorksheet(StyleBuilder, SharedStringBuilder);
            worksheetPart.Worksheet = ws;

        }


        /// <summary>
        /// Append a  <see cref="WorkbookStylesPart"/> to <paramref name="workbookPart"/> based on the
        /// content of <see cref="StyleBuilder"/>
        /// </summary>
        /// <param name="workbookPart">A reference to the <see cref="WorkbookPart"/> where to append the <see cref="WorkbookStylesPart"/></param>
        private void AppendWorkbookStylePart(ref WorkbookPart workbookPart)
        {
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var stylesheet = new Stylesheet
            {
                CellStyleFormats = new CellStyleFormats(new CellFormat())
            };

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
            var sharedStringTable = SharedStringBuilder.GetSharedStringTable();
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
        private static string BuildSuitableSheetName(Sheets sheets, string nameSheet)
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
    }
}
