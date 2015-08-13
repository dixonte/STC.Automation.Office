using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies how to shift cells to replace deleted cells.
    /// </summary>
    public enum DeleteShiftDirection
    {
        /// <summary>
        /// Cells are shifted to the left.
        /// </summary>
        ToLeft = -4159,
        /// <summary>
        /// Cells are shifted up.
        /// </summary>
        Up = -4162
    }
}
