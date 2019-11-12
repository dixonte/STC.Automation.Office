using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    [WrapsCOM("Excel.PivotFields", "00020876-0000-0000-C000-000000000046")]
    public class PivotItem : ComWrapper
    {
        internal PivotItem(object pivotitem)
            : base(pivotitem)
        {
        }

        /// <summary>
        /// Returns or sets a Long value that represents the position of the item in its field, if the item is currently showing.
        /// </summary>
        public long Position
        {
            get
            {
                return (long)InternalObject.GetType().InvokeMember("Position", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Position", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
