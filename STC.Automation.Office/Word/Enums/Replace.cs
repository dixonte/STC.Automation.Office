using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies the number of replacements to be made when find and replace is used.
    /// </summary>
    public enum Replace
    {
        /// <summary>
        /// Replace all occurrences.
        /// </summary>
        All = 2,
        /// <summary>
        /// Replace no occurrences.
        /// </summary>
        None = 0,
        /// <summary>
        /// Replace the first occurrence encountered.
        /// </summary>
        One = 1
    }
}
