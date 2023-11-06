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
        private Dictionary<int, ContentMeasure> ColumnMeasurements { get; set; } = new();

        /// <summary>
        /// Initialize the memory of columns width and enable those columns which must be auto measured.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        internal void Initialize<T>(TabularSheet<T> table)
        {
            ColumnMeasurements.Clear();
            IColumnWidth? widthTable = table.Options.ColumnOptions.Width;
            foreach (var column in table.Columns)
            {
                IColumnWidth? widthColumn = column.Options.Width;
                ContentMeasure? measure = null;

                if (widthColumn is not null)
                    measure = widthColumn.InitializeContentMeasure();
                else if (widthTable is not null)
                    measure = widthTable.InitializeContentMeasure();

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

            int length = default;
            if (content is not null)
            {
                Type type = content.GetType();
                if (BuildHelpers.NumericTypes.Contains(type))
                    length = ColumnMeasurer.MeasureNumericContent(content, format);
                else if (BuildHelpers.TimeSpanTypes.Contains(type))
                    length = ColumnMeasurer.MeasureTimeSpanContent(content, format);
                else if (BuildHelpers.DateTimeTypes.Contains(type))
                    length = ColumnMeasurer.MeasureDateTimeContent(content, format);
                else
                    length = MeasureContentAsString(content, format);
            }

            length = length > BuildHelpers.DefaultColumnWidth ? length : BuildHelpers.DefaultColumnWidth;
            ColumnMeasurements[col].EvaluateWidth(length, fontSize);
        }

        /// <summary>
        /// Calculate the ideal width of che column, or <see langword="null"/> if no width to configure.
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        internal double? EstimateColumnWidth(int col)
        {
            if (!ColumnMeasurements.ContainsKey(col))
                return null;

            ContentMeasure measure = ColumnMeasurements[col];
            const double coefficient = 1.4; // Extra increment to reach the value of excel
            double sizeFactor = measure.MaxFontSize / BuildHelpers.DefaultFontSize;

            double result = (measure.MaxContentWidth + coefficient) * sizeFactor * measure.FontFactor;
            return result;
        }

        #region Measure Content
        private static int MeasureNumericContent(object content, string? format)
        {
            try
            {
                // On ausence of format, handle as double value (most compatible, because include decimals and can be measure the string length)
                string? formatted = format is null ? Convert.ToDouble(content).ToString() : content.ToString();
                int length = formatted?.Length ?? default;
                return length;
            }
            catch (FormatException)
            {
                return default;
            }
        }

        private static int MeasureTimeSpanContent(object content, string? format)
        {
            try
            {
                int length = ((TimeSpan)content).ToString("c").Length;
                return length;
            }
            catch (FormatException)
            {
                return default;
            }
        }

        private static int MeasureDateTimeContent(object content, string? format)
        {
            try
            {
                int length = ((DateTime)content).ToString(format).Length;
                return length;
            }
            catch (FormatException)
            {
                return default;
            }
        }

        private int MeasureContentAsString(object content, string? format)
            => content.ToString()?.Length ?? default;
        #endregion
    }
}
