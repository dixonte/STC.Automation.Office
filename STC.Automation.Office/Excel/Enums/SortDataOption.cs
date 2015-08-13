using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies how to sort text.
    /// </summary>
    public enum SortDataOption
    {
        /// <summary>
        /// Default. Sorts numeric and text data separately.
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Treat text as numeric data for the sort.
        /// </summary>
        TextAsNumbers = 1

    }
}
