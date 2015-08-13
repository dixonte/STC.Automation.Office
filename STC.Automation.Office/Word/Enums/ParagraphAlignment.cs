using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies the alignment of a paragraph.
    /// </summary>
    public enum ParagraphAlignment
    {
        /// <summary>
        /// Center-aligned.
        /// </summary>
        Center = 1,
        /// <summary>
        /// Paragraph characters are distributed to fill the entire width of the paragraph.
        /// </summary>
        Distribute = 4,
        /// <summary>
        /// Fully justified.
        /// </summary>
        Justify = 3,
        /// <summary>
        /// Justified with a high character compression ratio.
        /// </summary>
        JustifyHi = 7,
        /// <summary>
        /// Justified with a low character compression ratio.
        /// </summary>
        JustifyLow = 8,
        /// <summary>
        /// Justified with a medium character compression ratio.
        /// </summary>
        JustifyMed = 5,
        /// <summary>
        /// Left-aligned.
        /// </summary>
        Left = 0,
        /// <summary>
        /// Right-aligned.
        /// </summary>
        Right = 2,
        /// <summary>
        /// Justified according to Thai formatting layout.
        /// </summary>
        ThaiJustify = 9
    }
}
