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
    /// Wraps an Core.CommandBarPopup COM object
    /// </summary>
    [WrapsCOM("Core.CommandBarPopup", CommandBarPopup.UUID)]
    public class CommandBarPopup : CommandBarControl
    {
        public const string UUID = "000C030A-0000-0000-C000-000000000046";

        private CommandBarControls _controls;

        /// <summary>
        /// Wraps the given COM object as a CommandBarPopup.
        /// </summary>
        /// <param name="commandBarPopupObj"></param>
        public CommandBarPopup(object commandBarPopupObj)
            : base(commandBarPopupObj)
        {
        }

        public CommandBarControls Controls
        {
            get
            {
                if (_controls == null || _controls.IsDisposed)
                    _controls = new CommandBarControls(InternalObject.GetType().InvokeMember("Controls", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _controls;
            }
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_controls != null && !_controls.IsDisposed)
                {
                    _controls.Dispose();
                    _controls = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
