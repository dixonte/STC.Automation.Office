using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// A collection of Column objects that represent the columns in a table.
    /// </summary>
    [WrapsCOM("Word.Columns", "0002094B-0000-0000-C000-000000000046")]
    public class Columns : ComWrapper
    {
        internal Columns(object columnsObj)
            : base(columnsObj)
        {
        }

        /// <summary>
        /// Returns the number of items in the columns collection.
        /// </summary>
        public long Count
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Deletes the specified table columns.
        /// </summary>
        public void Delete()
        {
            InternalObject.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }
    }
}
