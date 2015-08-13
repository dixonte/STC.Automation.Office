using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies how much of the document to export.
    /// </summary>
    public enum ExportRange
    {
        /// <summary>
        /// Exports the entire document.
        /// </summary>
        AllDocument = 0,

        /// <summary>
        /// Exports the current page.
        /// </summary>
        CurrentPage = 2,

        /// <summary>
        /// Exports the contents of a range using the starting and ending positions.
        /// </summary>
        ExportFromTo = 3,

        /// <summary>
        /// Exports the contents of the current selection.
        /// </summary>
        ExportSelection = 1
    }
}
