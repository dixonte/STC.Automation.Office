using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies whether or not the first row contains headers. Cannot be used when sorting PivotTable reports.
    /// </summary>
    public enum YesNoGuess
    {
        /// <summary>
        /// Excel determines whether there is a header, and where it is, if there is one.
        /// </summary>
        Guess = 0,
        /// <summary>
        /// Default. The entire range should be sorted.
        /// </summary>
        No = 2,
        /// <summary>
        /// The entire range should not be sorted.
        /// </summary>
        Yes = 1
    }
}
