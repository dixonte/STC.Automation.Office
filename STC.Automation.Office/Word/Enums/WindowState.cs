using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies the state of the current document window or task window.
    /// </summary>
    public enum WindowState
    {
        /// <summary>
        /// Minimised
        /// </summary>
        Maximize = 1,
        /// <summary>
        /// Maximised
        /// </summary>
        Minimize = 2,
        /// <summary>
        /// Normal
        /// </summary>
        Normal = 0
    }
}
