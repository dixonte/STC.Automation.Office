using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// A collection of Cell objects in a table column, table row, selection, or range.
    /// </summary>
    [WrapsCOM("Word.Cells", "0002094A-0000-0000-C000-000000000046")]
    public class Cells : ComWrapper
    {
        internal Cells(object cellsObj)
            : base(cellsObj)
        {
        }

        /// <summary>
        /// Returns the number of items in the cells collection.
        /// </summary>
        public long Count
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Merges the specified table cells with one another. The result is a single table cell.
        /// </summary>
        public void Merge()
        {
            InternalObject.GetType().InvokeMember("Merge", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Splits a range of table cells.
        /// </summary>
        /// <param name="numRows">The number of rows that the cell or group of cells is to be split into.</param>
        /// <param name="numColumns">The number of columns that the cell or group of cells is to be split into.</param>
        /// <param name="mergeBeforeSplit">True to merge the cells with one another before splitting them.</param>
        public void Split(int numRows, int numColumns, bool mergeBeforeSplit)
        {
            InternalObject.GetType().InvokeMember("Split", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { numRows, numColumns, mergeBeforeSplit });
        }
    }
}
