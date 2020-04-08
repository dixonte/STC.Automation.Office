using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    [WrapsCOM("Excel.PivotField", "00020874-0000-0000-C000-000000000046")]
    public class PivotField : ComWrapper
    {
        internal PivotField(object pivotfield)
            : base(pivotfield)
        {
        }

        /// <summary>
        /// Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the visible and hidden items (a PivotItems object) in the specified field. Read-only.
        /// </summary>
        /// <param name="index">The number of the field to be returned.</param>
        /// <returns></returns>
        public PivotItem PivotItems(int index)
        {
            return new PivotItem(InternalObject.GetType().InvokeMember("PivotItems", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
        }

        /// <summary>
        /// Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the visible and hidden items (a PivotItems object) in the specified field. Read-only.
        /// </summary>
        /// <param name="index">The name of the field to be returned.</param>
        /// <returns></returns>
        public PivotItem PivotItems(string index)
        {
            return new PivotItem(InternalObject.GetType().InvokeMember("PivotItems", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
        }
    }
}
