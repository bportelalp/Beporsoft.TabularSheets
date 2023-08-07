using Beporsoft.TabularSheets.Options;
using Beporsoft.TabularSheets.Options.ColumnWidth;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// A class to handle the measurement of the contents by column to estimate the ideal width
    /// </summary>
    internal class ColumnMeasurer
    {
        private Dictionary<int, ColMeasure> ColumnMeasurements { get; set; } = new();

        /// <summary>
        /// Initialize the memory of columns width and enable those columns which must be auto measured.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        internal void Initialize<T>(TabularSheet<T> table)
        {
            ColumnMeasurements.Clear();
            //bool autoWidthDefault = table.Options.DefaultColumnOptions.AutoWidth ?? false;
            IColumnWidth? widthDefault = table.Options.ColumnOptions.Width;
            foreach (var column in table.Columns)
            {
                ColMeasure? measure = null;
                if (column.Options.Width is FixedColumnWidth fixedWidth)
                {
                    measure = new()
                    {
                        MaxContentWidth = fixedWidth.Width,
                        AutoWidth = false
                    };
                }
                else if (column.Options.Width is AutoColumnWidth autoWidth)
                {
                    measure = new()
                    {
                        AutoWidth = true,
                        FontFactor = autoWidth.ScaleFactor
                    };
                }
                else if (widthDefault is FixedColumnWidth fixedWidthDefault)
                {
                    measure = new()
                    {
                        MaxContentWidth = fixedWidthDefault.Width,
                        AutoWidth = false
                    };
                }
                else if (widthDefault is AutoColumnWidth autoWidthDefault)
                {
                    measure = new()
                    {
                        AutoWidth = true,
                        FontFactor = autoWidthDefault.ScaleFactor
                    };
                }

                if (measure is not null)
                    ColumnMeasurements.Add(column.Index, measure);
            }
        }

        /// <summary>
        /// Evaluate the column content trying to measure with the specified format and save the results.
        /// This operation is not performed if the respective column has no autowidth
        /// </summary>
        /// <param name="col">Index of the column</param>
        /// <param name="content">Content of the cell</param>
        /// <param name="format">Numbering format of the cell</param>
        /// <param name="fontSize"></param>
        internal void MeasureContent(int col, object? content, string? format, double? fontSize)
        {
            if (!ColumnMeasurements.ContainsKey(col) || ColumnMeasurements[col].AutoWidth is false)
                return;

            int length = BuildHelpers.DefaultColumnWidth;
            string formatted;
            if (content is not null)
            {
                Type type = content.GetType();
                try
                {
                    if (BuildHelpers.NumericTypes.Contains(type))
                    {
                        if (format is null)
                        {
                            formatted = content.ToString()!;
                            if (string.IsNullOrWhiteSpace(formatted) || formatted.Length > BuildHelpers.DefaultColumnWidth)
                                length = 10;
                            else
                                length = formatted.Length;
                        }
                        else
                        {
                            double doub = Convert.ToDouble(content);
                            formatted = doub.ToString(format);
                            length = formatted.Length;
                        }
                    }
                    else if (BuildHelpers.TimeSpanTypes.Contains(type))
                    {
                        length = ((TimeSpan)content).ToString("c").Length;
                    }
                    else if (BuildHelpers.DateTimeTypes.Contains(type))
                    {
                        length = ((DateTime)content).ToString(format).Length;
                    }
                    else
                    {
                        length = content.ToString()?.Length ?? 0;
                    }
                }
                catch (FormatException)
                {
                    // Omit format exceptions
                }
            }

            if (ColumnMeasurements.ContainsKey(col))
                ColumnMeasurements[col].EvaluateWidth(length, fontSize);
        }

        /// <summary>
        /// Calculate the ideal width of che column, or <see langword="null"/> if no width to configure.
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        public double? EstimateColumnWidth(int col)
        {
            if (!ColumnMeasurements.ContainsKey(col))
                return null;

            ColMeasure measure = ColumnMeasurements[col];
            const double coefficient = 1.4; // Extra increment to reach the value of excel
            double sizeFactor = measure.MaxFontSize / BuildHelpers.DefaultFontSize;

            double result = (measure.MaxContentWidth + coefficient) * sizeFactor * measure.FontFactor;
            return result;
        }


        /// <summary>
        /// Container of the measurement of the column
        /// </summary>
        private class ColMeasure
        {
            /// <summary>
            /// If <see langword="true"/>, <see cref="MaxContentWidth"/> will be updating with new values.
            /// Otherwise, <see cref="MaxContentWidth"/> is the default value on the initialization and it doesn't change.
            /// </summary>
            public bool AutoWidth { get; set; }

            /// <summary>
            /// Value of width to apply to the column
            /// </summary>
            public double MaxContentWidth { get; set; }

            /// <summary>
            /// Max font size applied to any cell column
            /// </summary>
            public double MaxFontSize { get; set; } = BuildHelpers.DefaultFontSize;

            /// <summary>
            /// A scale applied for fonts different from calibri
            /// </summary>
            public double FontFactor { get; set; } = 1.0;

            /// <summary>
            /// Check a new value and save if it is greather than the current.
            /// </summary>
            /// <param name="width"></param>
            /// <param name="fontSize"></param>
            public void EvaluateWidth(double width, double? fontSize)
            {
                if (!AutoWidth)
                    return;
                if (MaxContentWidth < width)
                    MaxContentWidth = width;
                if (fontSize is not null && fontSize > MaxFontSize)
                    MaxFontSize = fontSize.Value;
            }
        }
    }
}
