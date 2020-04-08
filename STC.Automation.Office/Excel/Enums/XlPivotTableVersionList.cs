using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// behave in the same manner as they did in the corresponding version of Excel.
    /// </summary>
    public enum XlPivotTableVersionList
    {
        /// <summary>
        /// Excel 2000
        /// </summary>
        xlPivotTableVersion2000 = 0,

        /// <summary>
        /// Excel 2002
        /// </summary>
        xlPivotTableVersion10 = 1,

        /// <summary>
        /// Excel 2003
        /// </summary>
        xlPivotTableVersion11 = 2,

        /// <summary>
        /// Excel 2007
        /// </summary>
        xlPivotTableVersion12 = 3,

        /// <summary>
        /// Excel 2010
        /// </summary>
        xlPivotTableVersion14 = 4,

        /// <summary>
        /// Excel 2013
        /// </summary>
        xlPivotTableVersion15 = 5,

        /// <summary>
        /// Excel 6
        /// </summary>
        xlPivotTableVersion16 = 6,

        /// <summary>
        /// Provided only for backward compatibility
        /// </summary>
        xlPivotTableVersionCurrent = -1
    }
}
