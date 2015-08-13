using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Core
{
    /// <summary>
    /// Wraps an Core.CommandBarComboBox COM object
    /// </summary>
    [WrapsCOM("Core.CommandBarComboBox", CommandBarComboBox.UUID)]
    public class CommandBarComboBox : CommandBarControl
    {
        public const string UUID = "000C030C-0000-0000-C000-000000000046";

        /// <summary>
        /// Wraps the given COM object as a CommandBarComboBox.
        /// </summary>
        /// <param name="commandBarComboBoxObj"></param>
        public CommandBarComboBox(object commandBarComboBoxObj)
            : base(commandBarComboBoxObj)
        {
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
            }

            base.Dispose(true);
        }

        #endregion
    }
}
