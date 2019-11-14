using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies a tri-state value.
    /// </summary>
    public enum MsoTriState
    {
        /// <summary>
        /// Not supported
        /// </summary>
        msoCTrue = 1,

        /// <summary>
        /// False
        /// </summary>
        msoFalse = 0,

        /// <summary>
        /// Not supported
        /// </summary>
        msoTriStateMixed = -2,

        /// <summary>
        /// Not supported
        /// </summary>
        msoTriStateToggle = -3,

        /// <summary>
        /// True
        /// </summary>
        msoTrue = -1
    }
}
