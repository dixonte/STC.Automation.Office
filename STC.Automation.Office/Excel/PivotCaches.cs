using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Excel.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    [WrapsCOM("Excel.PivotCache", "0002441D-0000-0000-C000-000000000046")]
    public class PivotCaches : ComWrapper
    {
        internal PivotCaches(object pivotcachesObj)
            : base(pivotcachesObj)
        {
        }

        /// <summary>
        /// Creates a new PivotCache.
        /// </summary>
        /// <param name="SourceType">SourceType can be one of these XlPivotTableSourceType constants: xlConsolidation, xlDatabase, or xlExternal.</param>
        /// <param name="SourceData">The data for the new PivotTable cache.</param>
        /// <param name="Version">Version of the PivotTable. Version can be one of the XlPivotTableVersionList constants.</param>
        public PivotCache Create(XlPivotTableSourceType SourceType, object SourceData, XlPivotTableVersionList Version)
        {
            return new PivotCache(InternalObject.GetType().InvokeMember("Create", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, ComArguments.Prepare(SourceType, SourceData, Version)));
        }
    }
}
