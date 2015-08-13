using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    public enum PrintOutRange
    {
        /// <summary>
        /// The entire document.
        /// </summary>
        PrintAllDocument = 0,
        /// <summary>
        /// The current page.
        /// </summary>
        PrintCurrentPage = 2,
        /// <summary>
        /// A specified range.
        /// </summary>
        PrintFromTo = 3,
        /// <summary>
        /// A specified range of pages.
        /// </summary>
        PrintRangeOfPages = 4,
        /// <summary>
        /// The current selection.
        /// </summary>
        PrintSelection = 1
    }
}
