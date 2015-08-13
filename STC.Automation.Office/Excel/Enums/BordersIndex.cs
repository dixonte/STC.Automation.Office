using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the border to be retrieved.
    /// </summary>
    public enum BordersIndex
    {
        /// <summary>
        /// Border running from the upper left-hand corner to the lower right of each cell in the range.
        /// </summary>
        DiagonalDown = 5,
        /// <summary>
        /// Border running from the lower left-hand corner to the upper right of each cell in the range.
        /// </summary>
        DiagonalUp = 6,
        /// <summary>
        /// Border at the bottom of the range.
        /// </summary>
        EdgeBottom = 9,
        /// <summary>
        /// Border at the left-hand edge of the range.
        /// </summary>
        EdgeLeft = 7,
        /// <summary>
        /// Border at the right-hand edge of the range.
        /// </summary>
        EdgeRight = 10,
        /// <summary>
        /// Border at the top of the range.
        /// </summary>
        EdgeTop = 8,
        /// <summary>
        /// Horizontal borders for all cells in the range except borders on the outside of the range.
        /// </summary>
        InsideHorizontal = 12,
        /// <summary>
        /// Vertical borders for all the cells in the range except borders on the outside of the range.
        /// </summary>
        InsideVertical = 11
    }
}
