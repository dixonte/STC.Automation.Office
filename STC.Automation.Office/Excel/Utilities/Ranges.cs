using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace STC.Automation.Office.Excel.Utilities
{
    /// <summary>
    /// A collection of functions to aid in converting between range references and indexes.
    /// </summary>
    public sealed class Ranges
    {
        private static Regex _cell, _range;

        static Ranges()
        {
            _cell = new Regex(@"\$?([A-Za-z]+)\$?([0-9]+)");
            _range = new Regex(@"(\$?[A-Za-z]+\$?[0-9]+):(\$?[A-Za-z]+\$?[0-9]+)");
        }

        /// <summary>
        /// Parse cell reference into row and column indexes.  For example, "B5" = row 5, column 2.
        /// </summary>
        /// <param name="cell">The cell reference to parse (e.g. "B5")</param>
        /// <param name="row">The row index (e.g. 5)</param>
        /// <param name="col">The column index (e.g. 2)</param>
        public static void Parse(string cell, out int row, out int col)
        {
            Match m = _cell.Match(cell);
            if (m != null && m.Success)
            {
                row = Convert.ToInt32(m.Groups[2].Value);
                col = Columns.Parse(m.Groups[1].Value.ToUpper());
            }
            else
            {
                throw new ArgumentException("Cell reference must be characters followed by numbers.  For example, \"B5\" or \"XY235\".");
            }
        }

        /// <summary>
        /// Parse cell range into row and column indexes.  For example, "B5:AD235" = row 5, column 2 to row 235, column 30.
        /// </summary>
        /// <param name="cellRange">The cell range to parse (e.g. "B5:C10")</param>
        /// <param name="row1">The first row index (e.g. 5)</param>
        /// <param name="col1">The first column index (e.g. 2)</param>
        /// <param name="row2">The last row index (e.g. 10)</param>
        /// <param name="col2">The last column index (e.g. 3)</param>
        public static void Parse(string cellRange, out int row1, out int col1, out int row2, out int col2)
        {
            Match m = _range.Match(cellRange);
            if (m != null && m.Success)
            {
                Parse(m.Groups[1].Value, out row1, out col1);
                Parse(m.Groups[2].Value, out row2, out col2);
            }
            else
            {
                throw new ArgumentException("Cell range must be two cell references seperated by a colon.  For example, \"B5:C10\".");
            }
        }

        /// <summary>
        /// Format a cell range by row and column index.  For example, row 5, column 2 returns "B5".
        /// </summary>
        /// <param name="row">The row index, 1-based.</param>
        /// <param name="col">The column index, 1-based.</param>
        /// <returns>The cell reference.</returns>
        public static string Format(int row, int col)
        {
            return String.Concat(Columns.Format(col), row);
        }

        /// <summary>
        /// Format a range by row and column index.  For example (5,2) to (235,30) returns "B5:AD235".
        /// </summary>
        /// <param name="row1">The first row index, 1-based.</param>
        /// <param name="col1">The first column index, 1-based.</param>
        /// <param name="row2">The last row index, 1-based.</param>
        /// <param name="col2">The last column index, 1-based.</param>
        /// <returns>The range.</returns>
        public static string Format(int row1, int col1, int row2, int col2)
        {
            return String.Concat(Columns.Format(col1), row1, ':', Columns.Format(col2), row2);
        }

        /// <summary>
        /// Convert from a .NET string format into Excel NumberFormat.  Uses default culture.
        /// </summary>
        /// <param name="dotNetFormat"></param>
        /// <returns></returns>
        public static string ConvertFormat(string dotNetFormat)
        {
            if (String.IsNullOrEmpty(dotNetFormat))
                return null;

            // http://msdn.microsoft.com/en-us/library/dwhawy9k.aspx
            var regex = new Regex(@"^([CDEFGNPRX])(\d{0,2})$");
            var m = regex.Match(dotNetFormat.ToUpper());
            if (m != null && m.Success)
            {
                string formatSpecifier = m.Groups[1].Value;
                int precisionSpecifier = -1;
                if (!String.IsNullOrEmpty(m.Groups[2].Value))
                    precisionSpecifier = Int32.Parse(m.Groups[2].Value); // Number

                switch (formatSpecifier) // Letter
                {
                    case "C": // currency
                        return string.Concat("$", GetFormat(precisionSpecifier < 0 ? 2 : precisionSpecifier, true));

                    case "D": // decimal, integral only
                        return new string('0', Math.Max(1, precisionSpecifier));

                    case "E": // exponential
                        return String.Concat(GetFormat(precisionSpecifier, false), "E+000");

                    case "F": // fixed-point
                        return GetFormat(precisionSpecifier, false);

                    case "N": // numeric
                        return GetFormat(precisionSpecifier, true);

                    case "P": // percent
                        return String.Concat(GetFormat((precisionSpecifier < 0 ? 2 : precisionSpecifier), true), "%");

                    case "G": // general
                    case "R": // round-trip
                    case "X": // hexadecimal
                    default:
                        return string.Empty;
                }
            }

            // http://msdn.microsoft.com/en-us/library/az4se3k1.aspx
            switch (dotNetFormat)
            {
                case "d":
                    return "d/m/yyyy";

                case "D":
                    return "dddd, mmmm dd, yyyy";

                case "f":
                    return "dddd, mmmm dd, yyyy h:mm AM/PM";

                case "F":
                case "U":
                    return "dddd, mmmm dd, yyyy h:mm:ss AM/PM";

                case "g":
                    return "d/m/yyyy h:mm AM/PM";

                case "G":
                    return "d/m/yyyy h:mm:ss AM/PM";

                case "m":
                case "M":
                    return "mmmm dd";

                case "o":
                case "O":
                case "s":
                    return "yyyy-mm-ddThh:mm:ss";

                case "r":
                case "R":
                    return "ddd, dd mmm yyyy hh:mm:ss";

                case "t":
                    return "h:mm AM/PM";

                case "T":
                    return "h:mm:ss AM/PM";

                case "u":
                    return "yyyy-mm-dd hh:mm:ssZ";

                case "y":
                case "Y":
                    return "mmmm, yyyy";

                default:
                    return dotNetFormat;
            }
        }

        private static string GetFormat(int precision, bool thousands)
        {
            precision = Math.Max(0, precision);

            if (precision == 0)
                return (thousands ? "#,##0" : "0");
            
            return String.Concat((thousands ? "#,##0" : "0"), '.', new string('0', precision));
        }
    }
}
