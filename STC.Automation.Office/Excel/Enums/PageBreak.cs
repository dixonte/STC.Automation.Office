using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the page orientation when the worksheet is printed (requires Excel 2007)
    /// </summary>
    public enum PageBreak
    {
        /// <summary>
        ///Excel will automatically add page breaks.
        /// </summary>
        PageBreakAutomatic = -4105,

        /// <summary>
        /// Page breaks are manually inserted.
        /// </summary>
        PageBreakManual = -4135,

        /// <summary>
        /// Page breaks are not inserted in the worksheet.
        /// </summary>
        PageBreakNone = -4142
    }
}
