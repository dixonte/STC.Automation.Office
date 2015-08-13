using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace STC.Automation.Office.Excel.Utilities
{
    /// <summary>
    /// A collection of functions to aid in converting between column names and indexes.
    /// </summary>
    public sealed class Columns
    {
        /// <summary>
        /// Calculate the index (1-based) of the given column name.  Column names should be alpha characters only (e.g. "A", "BC", "DEF").
        /// </summary>
        /// <remarks>
        /// Logic used when working out formula for function:
        /// (where A = 1 and Z = 26)
        /// Z = 26
        /// AA = 27 (26 + 1) 			= 26 * 1 + 1
        /// AB = 28 (26 + 2) 			= 26 * 1 + 2
        /// BA = 53 (26 + 26 + 1)		= 26 * 2 + 1
        /// BC = 55 (26 + 26 + 3)		= 26 * 2 + 3
        /// CD = 82 (26 + 26 + 26 + 4)	= 26 * 3 + 4
        /// ...
        /// YW = 26 * 25 + 23
        /// ZW = 26 * 26 + 23
        /// ZZ = 26 * 26 + 26
        /// AAA = 26 * 27 + 1 = 26 * (26 + 1) + 1, where (26 + 1) is AA
        /// AAB = 26 * (26 + 1) + 2
        /// ABA = 26 * (26 + 2) + 1, where (26 + 2) is AB
        /// ABC = 26 * (26 + 2) + 3
        /// BAA = 26 * (26 * 2 + 1) + 1, where (26 * 2 + 1) is BA
        /// BCA = 26 * (26 * 2 + 3) + 1
        /// CDA = 26 * (26 * 3 + 4) + 1
        /// </remarks>
        /// <param name="column">The column name whose index will be calculated.</param>
        /// <returns>The 1-based index of the column.</returns>
        /// <exception cref="System.ArgumentException">Thrown if the column is null, empty or contains characters other than alpha.</exception>
        public static int Parse(string column)
        {
            Regex regex = new Regex("^[A-Za-z]+$");
            if (regex.IsMatch(column))
                return DecodeColumnRecursive(column.ToUpper());
            else
                throw new ArgumentException("Column should only contain alpha characters.");
        }

        private static int DecodeColumnRecursive(string column)
        {
            if (column.Length == 1)
                return (int)column[0] - 64;
            else
                return 26 * DecodeColumnRecursive(column.Substring(0, column.Length - 1)) + DecodeColumnRecursive(column[column.Length - 1].ToString());
        }

        /// <summary>
        /// Calculate the alphabetic representation of a column using a 1-based index.  For example, 1 = "A", 2 = "B", 26 = "Z", 27 = "AA", 704 = "AAB".
        /// </summary>
        /// <param name="columnIndex">The 1-based index of a column.</param>
        /// <returns>A string containing the alphabetic representation of the column index.</returns>
        public static string Format(int columnIndex)
        {
            if (columnIndex >= 1)
                return EncodeColumnRecursive(columnIndex);
            else
                throw new ArgumentException("Column index should be 1-based and greater than or equal to 1.");
        }

        private static string EncodeColumnRecursive(int columnIndex)
        {
            int whole = (int)Math.Floor((columnIndex - 1) / 26.0);
            int remainder = (columnIndex - 1) % 26;

            if (whole <= 0)
                return ((char)(65 + remainder)).ToString();
            else
                return EncodeColumnRecursive(whole) + EncodeColumnRecursive(remainder + 1);
        }


        /// <summary>
        /// Given a column range string, return an array of column indexes.  Indexes are not repeated for overlapping ranges and indexes are not sorted.  For example, "A:A" = { 1 }; "B:D" = { 2,3,4 }; "A:C,E:F,M:M" = { 1,2,3, 5,6, 13 }; "A:C,B:D" = { 1,2,3,4 };
        /// </summary>
        /// <param name="range">The string representation of a range.</param>
        /// <returns>An array of column indexes for all columns in the given range.</returns>
        public static int[] ParseRange(string range)
        {
            var indexes = new List<int>();

            Regex r = new Regex(@"([A-Za-z]+):([A-Za-z]+)");
            foreach (Match m in r.Matches(range))
            {
                // safe to index directly as our Regex takes care of validating input
                int from = Parse(m.Groups[1].Value);
                int to = Parse(m.Groups[2].Value);

                for (int i = from; i <= to; i++)
                {
                    if (!indexes.Contains(i))
                        indexes.Add(i);
                }
            }

            return indexes.ToArray();
        }
    }
}
