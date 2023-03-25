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
using Beporsoft.TabularSheets.Builders.Interfaces;
using System.IO;
using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Builders.SheetBuilders;

namespace Beporsoft.TabularSheets
{
    public class TabularSpreadsheet<T> : TabularData<T>, ISheet
    {
        public TabularSpreadsheet()
        {
        }

        public TabularSpreadsheet(IEnumerable<T> items)
        {
            Items = items.ToList();
        }
        #region Properties
        /// <summary>
        /// The title of the current sheet
        /// </summary>
        public string Title { get; set; } = "Sheet";
        public HeaderStyle HeaderStyle { get; private set; } = new();

        #endregion

        #region Configure Table
        public void SetSheetTitle(string title) => Title = title;
        #endregion

        #region Create
        public void Create(string path)
        {
            SpreadsheetBuilder builder = new();
            builder.Create(path, this);
        }

        public MemoryStream Create()
        {
            SpreadsheetBuilder builder = new();
            MemoryStream stream = builder.Create(this);
            return stream;
        }
        #endregion

        #region ISheet
        Type ISheet.ItemType => typeof(T);

        SheetData ISheet.BuildData(ref StylesheetBuilder stylesheetBuilder)
        {
            SheetBuilder<T> builder = new(this, stylesheetBuilder);
            return builder.BuildSheetData();
        }
        #endregion
    }
}
