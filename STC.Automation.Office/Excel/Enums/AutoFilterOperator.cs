using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies whether or not the first row contains headers. Cannot be used when sorting PivotTable reports.
    /// </summary>
    public enum AutoFilterOperator
    {
        /// <summary>
        /// Logical AND of Criteria1 and Criteria2.
        /// </summary>
        And = 1,

        /// <summary>
        /// Lowest-valued items displayed (number of items specified in Criteria1).
        /// </summary>
        Bottom10Items = 4,

        /// <summary>
        /// Lowest-valued items displayed (percentage specified in Criteria1).
        /// </summary>
        Bottom10Percent = 6,

        /// <summary>
        /// Color of the cell
        /// </summary>
        CellColor = 8,

        /// <summary>
        /// Dynamic filter
        /// </summary>
        Dynamic = 11,

        /// <summary>
        /// Color of the font
        /// </summary>
        FontColor = 9,

        /// <summary>
        /// Filter icon
        /// </summary>
        Icon = 10,

        /// <summary>
        /// Filter values
        /// </summary>
        Values = 7,

        /// <summary>
        /// Logical OR of Criteria1 or Criteria2.
        /// </summary>
        Or = 2,

        /// <summary>
        /// Highest-valued items displayed (number of items specified in Criteria1).
        /// </summary>
        Top10Items = 3,

        /// <summary>
        /// Highest-valued items displayed (percentage specified in Criteria1).
        /// </summary>
        Top10Percent = 5
    }
}
