using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies whether to export bookmarks and the type of bookmarks to export.
    /// </summary>
    public enum ExportCreateBookmarks
    {
        /// <summary>
        /// Create a bookmark in the exported document for each Microsoft Office Word heading, which includes only headings within the main document and text boxes not within headers, footers, endnotes, footnotes, or comments.
        /// </summary>
        CreateHeadingBookmarks = 1,

        /// <summary>
        /// Do not create bookmarks in the exported document.
        /// </summary>
        CreateNoBookmarks = 0,

        /// <summary>
        /// Create a bookmark in the exported document for each Word bookmark, which includes all bookmarks except those contained within headers and footers.
        /// </summary>
        CreateWordBookmarks = 2
    }
}
