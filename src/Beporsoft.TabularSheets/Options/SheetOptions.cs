﻿using Beporsoft.TabularSheets.CellStyling;
using System;
using System.Globalization;

namespace Beporsoft.TabularSheets.Options
{
    /// <summary>
    /// Settings for table and the building of sheet
    /// </summary>
    public class SheetOptions
    {
        /// <summary>
        /// The default Date time format applied to cells which contains a <see cref="DateTime"/> or <see cref="DateTimeOffset"/> value.
        /// <br/>
        /// By default the pattern is [ShortDatePattern LongTimePattern] regarding the current culture
        /// <br/>
        /// <br/>
        /// NB: OpenXml commonly stores the datetime as double and use the <see cref="DocumentFormat.OpenXml.Spreadsheet.CellValues.Number"/> 
        /// as the default datatype. This conversion is achieved using <see cref="DateTime.ToOADate"/> internally. The
        /// representation over the cell is commonly exposed using minimal stylesheet to show the common format. This field
        /// is required to apply as the default <see cref="Style.NumberingPattern"/> to every cell which contains a <see cref="DateTime"/> or 
        /// <see cref="DateTimeOffset"/> type, if no more specific style is applied on the styling properties.
        /// </summary>
        public string DateTimeFormat { get; set; } = DefaultDateTimeFormat;

        /// <summary>
        /// The default time format applied to cells which contains a <see cref="TimeSpan"/> value. <br/>
        /// By default the pattern is [HH]:mm:ss, where days value are showed as hours.
        /// <br/>
        /// <br/>
        /// NB: OpenXml commonly stores the time span as double and use the <see cref="DocumentFormat.OpenXml.Spreadsheet.CellValues.Number"/> 
        /// as the default datatype. The value saved on the cell is <see cref="TimeSpan.TotalDays"/> and the
        /// representation over the cell is commonly exposed using minimal stylesheet to show the common format. This field
        /// is required to apply as the default <see cref="Style.NumberingPattern"/> to every cell which contains a <see cref="TimeSpan"/> type, 
        /// if no more specific style is applied on the styling properties.
        /// </summary>
        public string TimeSpanFormat { get; set; } = "[HH]:mm:ss";

        /// <summary>
        /// When <see langword="true"/>, <see cref="TabularSheet{T}.HeaderStyle"/> <see langword="null"/> properties will be overrided from
        /// the equivalent <see cref="TabularSheet{T}.BodyStyle"/> properties
        /// </summary>
        public bool InheritHeaderStyleFromBody { get; set; } = false;

        /// <summary>
        /// Gets or sets if spreadsheet should include a autofilter
        /// </summary>
        internal bool UseAutofilter { get; set; } = false;

        /// <summary>
        /// The default column options for those which haven't configured their own options
        /// </summary>
        public ColumnOptions ColumnOptions { get; set; } = new();

        internal static string DefaultDateTimeFormat { get; set; } = 
            $"{CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern} {CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern}";

    }
}
