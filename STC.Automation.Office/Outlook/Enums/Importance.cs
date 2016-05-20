using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook.Enums
{
    /// <summary>
    /// Specifies the level of importance for an item marked by the creator of the item.
    /// </summary>
    public enum Importance
    {
        /// <summary>
        /// Item is marked as high importance.
        /// </summary>
        High = 2,
        /// <summary>
        /// Item is marked as low importance.
        /// </summary>
        Low = 0,
        /// <summary>
        /// Item is marked as medium importance.
        /// </summary>
        Normal = 1
    }
}
