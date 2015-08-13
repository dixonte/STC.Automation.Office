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
    /// Wraps an Core.CommandBar COM object
    /// </summary>
    [WrapsCOM("Core.CommandBar", "000C0304-0000-0000-C000-000000000046")]
    public class CommandBar : ComWrapper
    {
        private CommandBarControls _controls;

        /// <summary>
        /// Wraps the given COM object as a CommandBar.
        /// </summary>
        /// <param name="commandBarObj"></param>
        public CommandBar(object commandBarObj)
            : base(commandBarObj)
        {
        }

        public string Name
        {
            get
            {
                return (string)InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
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
