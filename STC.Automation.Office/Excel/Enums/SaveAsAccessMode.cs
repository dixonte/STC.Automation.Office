using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the access mode for the Save As function.
    /// </summary>
    public enum SaveAsAccessMode
    {
        /// <summary>
        /// Default (does not change the access mode)
        /// </summary>
        NoChange = 1,
        /// <summary>
        /// Share list
        /// </summary>
        Shared = 2,
        /// <summary>
        /// Exclusive mode
        /// </summary>
        Exclusive = 3
    }
}
