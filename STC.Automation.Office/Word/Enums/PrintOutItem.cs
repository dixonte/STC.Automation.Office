using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    public enum PrintOutItem
    {
        /// <summary>
        /// Autotext entries in the current document.
        /// </summary>
        PrintAutoTextEntries = 4,
        /// <summary>
        /// Comments in the current document.
        /// </summary>
        PrintComments = 2,
        /// <summary>
        /// Current document content.
        /// </summary>
        PrintDocumentContent = 0,
        /// <summary>
        /// Current document content including markup.
        /// </summary>
        PrintDocumentWithMarkup = 7,
        /// <summary>
        /// An envelope.
        /// </summary>
        PrintEnvelope = 6,
        /// <summary>
        /// Key assignments in the current document.
        /// </summary>
        PrintKeyAssignments = 5,
        /// <summary>
        /// Markup in the current document.
        /// </summary>
        PrintMarkup = 2,
        /// <summary>
        /// Properties in the current document.
        /// </summary>
        PrintProperties = 1,
        /// <summary>
        /// Styles in the current document.
        /// </summary>
        PrintStyles = 3
    }
}
