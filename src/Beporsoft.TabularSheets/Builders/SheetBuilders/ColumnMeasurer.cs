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
        private Dictionary<int, ColMeasure> ContentColumnMaxLength { get; set; } = new();

        internal void Initialize<T>(TabularSheet<T> table)
        {
            ContentColumnMaxLength.Clear();
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

                if(measure is not null)
                    ContentColumnMaxLength.Add(column.Index, measure);
            }
        }

        internal void MeasureContent(int col, object? content, string? format)
        {
            int length = BuildHelpers.DefaultColumnWidth;
            string formatted;
            if (content is not null)
            {
                Type type = content.GetType();
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

            if (ContentColumnMaxLength.ContainsKey(col))
                ContentColumnMaxLength[col].EvaluateWidth(length);
        }

        public double? EstimateColumnWidth(int col, double? fontSize)
        {
            if (!ContentColumnMaxLength.ContainsKey(col))
                return null;

            ColMeasure measure = ContentColumnMaxLength[col];
            //const double calibri11Width = 7.0;
            const double calibri11 = 11.0; // Points of Calibri 11, treated as the default.
            const double coefficient = 1.4; // Extra increment to reach the value of excel
            double inputWidth = measure.Width < 10 ? 10 : measure.Width;
            double size = fontSize is null ? calibri11 : fontSize.Value;
            double sizeFactor = size / calibri11;

            double result = (inputWidth + coefficient) * sizeFactor;
            return result;
        }

        private class ColMeasure
        {
            public bool AutoWidth { get; set; }
            public double Width { get; set; }

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
