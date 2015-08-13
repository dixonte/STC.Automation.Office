using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the worksheet type.
    /// </summary>
    public enum SheetType
    {
        /// <summary>
        /// Chart
        /// </summary>
        Chart = -4109,
        /// <summary>
        /// Dialog sheet
        /// </summary>
        Dialog = -4116,
        /// <summary>
        /// Excel version 4 international macro sheet
        /// </summary>
        Excel4IntlMacro = 4,
        /// <summary>
        /// Excel version 4 macro sheet
        /// </summary>
        Excel4Macro = 3,
        /// <summary>
        /// Worksheet
        /// </summary>
        Worksheet = -4167
    }
}
