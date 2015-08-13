using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the reference style.
    /// </summary>
    public enum XLReferenceStyle
    {
        /// <summary>
        /// Use xlA1 to return an A1-style reference.
        /// </summary>
        xlA1 = 1,

        /// <summary>
        /// Use xlR1C1 to return an R1C1-style reference.
        /// </summary>
        xlR1C1 = -4150

    }
}
