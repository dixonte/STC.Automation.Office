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


    }
}
