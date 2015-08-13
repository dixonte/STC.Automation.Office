using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies whether to export the document with markup.
    /// </summary>
    public enum ExportItem
    {
        /// <summary>
        /// Exports the document without markup.
        /// </summary>
        Content = 0,

        /// <summary>
        /// Exports the document with markup.
        /// </summary>
        WithMarkup = 7
    }
}
