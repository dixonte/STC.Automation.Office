using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies the resolution and quality of the exported document.
    /// </summary>
    public enum ExportOptimizeFor
    {
        /// <summary>
        /// Export for print, which is higher quality and results in a larger file size.
        /// </summary>
        Print = 0,

        /// <summary>
        /// Export for screen, which is a lower quality and results in a smaller file size.
        /// </summary>
        OnScreen=1
    }
}
