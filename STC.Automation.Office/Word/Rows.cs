using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// A collection of Row objects that represent the rows in a table.
    /// </summary>
    [WrapsCOM("Word.Columns", "0002094C-0000-0000-C000-000000000046")]
    public class Rows : ComWrapper
    {
        internal Rows(object rowsObj)
            : base(rowsObj)
        {
        }

        /// <summary>
        /// Returns the number of items in the rows collection.
        /// </summary>
        public long Count
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Deletes the specified table rows.
        /// </summary>
        public void Delete()
        {
            InternalObject.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }
    }
}
