using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Window COM object
    /// </summary>
    [WrapsCOM("Excel.Window", "00020893-0001-0000-C000-000000000046")]
    public class Window : ComWrapper
    {
        private Application _application;

        /// <summary>
        /// Wraps the given COM object as a Window.
        /// </summary>
        /// <param name="windowObj"></param>
        public Window(object windowObj)
            : base(windowObj)
        {
        }

        /// <summary>
        /// Gets the Application to which this Window object belongs. This object is internally cached and does not need to be manually disposed.
        /// </summary>
        public Application Application
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
        public Application GetNewApplication()
        {
            return new Application(InternalObject.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
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
