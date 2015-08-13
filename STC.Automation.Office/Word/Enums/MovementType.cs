using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies the way the selection is moved.
    /// </summary>
    public enum MovementType
    {
        /// <summary>
        /// The selection is collapsed to an insertion point and moved to the end of the specified unit.
        /// </summary>
        Move = 0,
        /// <summary>
        /// The end of the selection is extended to the end of the specified unit.
        /// </summary>
        Extend = 1
    }
}
