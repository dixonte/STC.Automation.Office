using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies type of break.
    /// </summary>
    public enum BreakType
    {
        /// <summary>
        /// Column break at the insertion point.
        /// </summary>
        ColumnBreak = 8,
        /// <summary>
        /// Line break.
        /// </summary>
        LineBreak = 6,
        /// <summary>
        /// Line break.
        /// </summary>
        LineBreakClearLeft = 9,
        /// <summary>
        /// Line break.
        /// </summary>
        LineBreakClearRight = 10,
        /// <summary>
        /// Page break at the insertion point.
        /// </summary>
        PageBreak = 7,
        /// <summary>
        /// New section without a corresponding page break.
        /// </summary>
        SectionBreakContinuous = 3,
        /// <summary>
        /// Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
        /// </summary>
        SectionBreakEvenPage = 4,
        /// <summary>
        /// Section break on next page.
        /// </summary>
        SectionBreakNextPage = 2,
        /// <summary>
        /// Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
        /// </summary>
        SectionBreakOddPage = 5,
        /// <summary>
        /// Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.
        /// </summary>
        TextWrappingBreak = 11
    }
}
