using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the source of the report data.
    /// </summary>
    public enum XlPivotTableSourceType
    {
        /// <summary>
        /// Multiple consolidation ranges.
        /// </summary> 

        /// <summary>
        /// Microsoft Excel list or database.
        /// </summary>
        xlConsolidation = 3,

        /// <summary>
        /// Data from another application.
        /// </summary>
        xlDatabase = 1,

        /// <summary>
        /// Data from another application.
        /// </summary>
        xlExternal = 2,

        /// <summary>
        /// Same source as another PivotTable report.
        /// </summary>
        xlPivotTable = -4148,

        /// <summary>
        /// Data is based on scenarios created using the Scenario Manager.
        /// </summary>
        xlScenario
    }
}
