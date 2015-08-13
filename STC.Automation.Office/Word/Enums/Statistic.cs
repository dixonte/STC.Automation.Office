using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies a statistic to return from a selection or item.
    /// </summary>
    public enum Statistic
    {
        /// <summary>
        /// Count of characters.
        /// </summary>
        Characters = 3,
        /// <summary>
        /// Count of characters including spaces.
        /// </summary>
        CharactersWithSpaces = 5,
        /// <summary>
        /// Count of characters for Asian languages.
        /// </summary>
        FarEastCharacters = 6,
        /// <summary>
        /// Count of lines.
        /// </summary>
        Lines = 1,
        /// <summary>
        /// Count of pages.
        /// </summary>
        Pages = 2,
        /// <summary>
        /// Count of paragraphs.
        /// </summary>
        Paragraphs = 4,
        /// <summary>
        /// Count of words.
        /// </summary>
        Words = 0
    }
}
