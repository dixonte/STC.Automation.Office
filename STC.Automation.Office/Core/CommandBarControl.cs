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
    /// Wraps an Core.CommandBarControl COM object
    /// </summary>
    [WrapsCOM("Core.CommandBarControl", "000C0308-0000-0000-C000-000000000046")]
    public class CommandBarControl : ComWrapper
    {
        public OfficeApplication _application;

        /// <summary>
        /// Wraps the given COM object as a CommandBarControl.
        /// </summary>
        /// <param name="commandBarControlObj"></param>
        public CommandBarControl(object commandBarControlObj)
            : base(commandBarControlObj)
        {
        }

        /// <summary>
        /// Gets the Application to which this Window object belongs. This object is internally cached and does not need to be manually disposed.
        /// </summary>
        public OfficeApplication Application
        {
            get
            {
                if (_application == null)
                {
                    _application = GetNewApplication();
                }

                return _application;
            }
        }

        /// <summary>
        /// Gets a copy of the Application to which this Window object belongs. This object needs to be manually disposed.
        /// </summary>
        public OfficeApplication GetNewApplication()
        {
            var app = InternalObject.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);

            if (SupportsInterface(app, new Guid(Access.Application.UUID)))
            {
                return new Access.Application(app);
            }
            else if (SupportsInterface(app, new Guid(Excel.Application.UUID)))
            {
                return new Excel.Application(app);
            }
            else if (SupportsInterface(app, new Guid(Word.Application.UUID)))
            {
                return new Word.Application(app);
            }
            else
            {
                throw new COMException("Unknown application type!");
            }
        }

        public string Caption
        {
            get
            {
                return (string)InternalObject.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_application != null)
                {
                    _application.Dispose();
                    _application = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
