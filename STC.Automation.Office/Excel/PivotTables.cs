using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// A collection of all the PivotTable objects in the specified workbook.
    /// </summary>
    ///
    [WrapsCOM("Excel.PivotTables", "00020873-0000-0000-C000-000000000046")]
    public class PivotTables : ComWrapper
    {
        internal PivotTables(object pivotTablesObj)
            : base(pivotTablesObj)
        {
        }

        /// <summary>
        /// Index the Pivot Tables
        /// </summary>
        /// <param name="id">pivot table index</param>
        /// <returns></returns>
        public PivotTable this[int id]
        {
            get
            {
                return new PivotTable(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { id }));
            }
        }

        /// <summary>
        /// Index the Pivot Tables
        /// </summary>
        /// <param name="name">pivot table name</param>
        /// <returns></returns>
        public PivotTable this[string name]
        {
            get
            {
                return new PivotTable(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { name }));
            }
        }
    }
}
