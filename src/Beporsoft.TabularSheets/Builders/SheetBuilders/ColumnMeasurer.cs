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
            bool autoWidthDefault = table.Options.DefaultColumnOptions.AutoWidth ?? false;
            double? widthDefault = table.Options.DefaultColumnOptions.Width;
            foreach (var column in table.Columns)
            {
                ColMeasure? measure = null;
                if (column.Options.Width.HasValue)
                {
                    measure = new()
                    {
                        Width = column.Options.Width.Value,
                        AutoWidth = false
                    };
                }
                else if (column.Options.AutoWidth is true)
                {
                    measure = new() { AutoWidth = true };
                }
                else if (widthDefault.HasValue)
                {
                    measure = new()
                    {
                        Width = widthDefault!.Value,
                        AutoWidth = false
                    };
                }
                else if (autoWidthDefault is true)
                {
                    measure = new() { AutoWidth = true };
                }

                if (measure is not null)
                    ColumnMeasurements.Add(column.Index, measure);
            }
        }

        /// <summary>
        /// Evaluate the column content trying to measure with the specified format and save the results.
        /// This operation is not performed if the respective column has no autowidth
        /// </summary>
        /// <param name="col"></param>
        /// <param name="content"></param>
        /// <param name="format"></param>
        internal void MeasureContent(int col, object? content, string? format)
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
                        //length = ((TimeSpan)content).ToString(format).Length;
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
                ColumnMeasurements[col].EvaluateWidth(length);
        }

        /// <summary>
        /// Calculate the ideal width of che column, or <see langword="null"/> if no width to configure.
        /// </summary>
        /// <param name="col"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        public double? EstimateColumnWidth(int col, double? fontSize)
        {
            if (!ColumnMeasurements.ContainsKey(col))
                return null;

            ColMeasure measure = ColumnMeasurements[col];
            //const double calibri11Width = 7.0;
            const double calibri11 = 11.0; // Points of Calibri 11, treated as the default.
            const double coefficient = 1.4; // Extra increment to reach the value of excel
            double inputWidth = measure.Width < 10 ? 10 : measure.Width;
            double size = fontSize is null ? calibri11 : fontSize.Value;
            double sizeFactor = size / calibri11;

            double result = (inputWidth + coefficient) * sizeFactor;
            return result;
        }


        /// <summary>
        /// Container of the measurement of the column
        /// </summary>
        private class ColMeasure
        {
            /// <summary>
            /// If <see langword="true"/>, <see cref="Width"/> will be updating with new values.
            /// Otherwise, <see cref="Width"/> is the default value on the initialization and it doesn't change.
            /// </summary>
            public bool AutoWidth { get; set; }

            /// <summary>
            /// Value of width to apply to the column
            /// </summary>
            public double Width { get; set; }

            /// <summary>
            /// Check a new value and save if it is greather than the current.
            /// </summary>
            /// <param name="width"></param>
            public void EvaluateWidth(double width)
            {
                if (!AutoWidth)
                    return;
                if (Width < width)
                    Width = width;
            }
        }
    }
}
