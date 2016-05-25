using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook.Enums
{
    /// <summary>
    /// Specifies the folder display mode.
    /// </summary>
    public enum FolderDisplayMode
    {
        /// <summary>
        /// Only the contents of the selected folder are displayed.
        /// </summary>
        FolderOnly = 1,
        /// <summary>
        /// Folder contents are displayed but no navigation pane is shown.
        /// </summary>
        NoNavigation = 2,
        /// <summary>
        /// Folder is displayed with navigation pane on the left and folder contents on the right.
        /// </summary>
        Normal = 0
    }
}
