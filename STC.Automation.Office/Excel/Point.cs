using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    [WrapsCOM("Excel.PivotCache", "0002086A-0000-0000-C000-000000000046")]
    public class Point: ComWrapper
    {
        internal Point(object intObj)
            : base(intObj)
        {
        }

        /// <summary>
        /// Returns a DataLabel object that represents the data label associated with the point. Read-only.
        /// </summary>
        public DataLabel DataLabel
        {
            get
            {
                return new DataLabel(InternalObject.GetType().InvokeMember("DataLabel", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }
    }
}
