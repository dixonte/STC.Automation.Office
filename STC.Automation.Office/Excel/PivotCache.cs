using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    [WrapsCOM("Excel.PivotCache", "0002441C-0000-0000-C000-000000000046")]
    public class PivotCache : ComWrapper
    {
        internal PivotCache(object pivotcacheObj)
            : base(pivotcacheObj)
        {
        }

        /// <summary>
        /// Causes the specified chart to be redrawn immediately.
        /// </summary>
        public void Refresh()
        {
            InternalObject.GetType().InvokeMember("Refresh", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// True if the PivotTable cache is automatically updated each time the workbook is opened. The default value is False. Read/write Boolean.
        /// </summary>
        public bool RefreshOnFileOpen
        {
            get
            {
                return (bool)(InternalObject.GetType().InvokeMember("RefreshOnFileOpen", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("RefreshOnFileOpen", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
