using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    //ToDo Add ole class guid thingy
    /// <summary>
    /// Represents a PivotTable report on a worksheet.
    /// </summary>
    public class PivotTable: ComWrapper
    {
        internal PivotTable(object pivottableobj)
            : base(pivottableobj)
        {
        }

        /// <summary>
        /// Changes the PivotCache object of the specified PivotTable.
        /// </summary>
        /// <param name="bstr">A PivotTable or PivotCache object that represents the new PivotCache for the specified PivotTable.</param>
        public void ChangePivotCache(PivotTable bstr)
        {
            InternalObject.GetType().InvokeMember("ChangePivotCache", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { bstr.InternalObject });
        }

        /// <summary>
        /// Changes the PivotCache object of the specified PivotTable.
        /// </summary>
        /// <param name="bstr">A PivotTable or PivotCache object that represents the new PivotCache for the specified PivotTable.</param>
        public void ChangePivotCache(PivotCache bstr)
        {
            InternalObject.GetType().InvokeMember("ChangePivotCache", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { bstr.InternalObject });
        }

        /// <summary>
        /// Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of both the visible and hidden fields (a PivotFields object) in the PivotTable report. Read-only.
        /// </summary>
        /// <param name="index">The number of the field to be returned.</param>
        /// <returns></returns>
        public PivotField PivotFields(int index)
        {
            return new PivotField(InternalObject.GetType().InvokeMember("PivotFields", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
        }

        /// <summary>
        /// Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of both the visible and hidden fields (a PivotFields object) in the PivotTable report. Read-only.
        /// </summary>
        /// <param name="index">The name of the field to be returned.</param>
        /// <returns></returns>
        public PivotField PivotFields(string index)
        {
            return new PivotField(InternalObject.GetType().InvokeMember("PivotFields", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
        }
    }
}
