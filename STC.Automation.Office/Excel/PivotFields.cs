using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// A collection of all the PivotField objects in a PivotTable report.
    /// </summary>
    [WrapsCOM("Excel.PivotFields", "00020875-0000-0000-C000-000000000046")]
    public class PivotFields : ComWrapper
    {
        internal PivotFields(object pivotfields)
            : base(pivotfields)
        {
        }

        /// <summary>
        /// Index the Pivotfields
        /// </summary>
        /// <param name="id">Pivot Field index</param>
        /// <returns></returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public PivotField this[int id]
        {
            get
            {
                return new PivotField(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { id }));
            }
        }
    }
}
