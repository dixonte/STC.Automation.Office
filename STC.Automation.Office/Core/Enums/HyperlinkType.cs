using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Core.Enums
{
    public enum HyperlinkType
    {
        /// <summary>
        /// Hyperlink applies to an inline shape. Used only with Microsoft Word.
        /// </summary>
        InlineShape = 2,

        /// <summary>
        /// Hyperlink applies to a Range object.
        /// </summary>
        Range = 0,

        /// <summary>
        /// Hyperlink applies to a Shape object.
        /// </summary>
        Shape = 1
    }
}
