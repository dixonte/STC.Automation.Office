using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word.Window COM object
    /// </summary>
    [WrapsCOM("Word.Window", "00020962-0000-0000-C000-000000000046")]
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
        /// Activates the specified object.
        /// </summary>
        public void Activate()
        {
            InternalObject.GetType().InvokeMember("Activate", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Gets a copy of the Application to which this Window object belongs. This object needs to be manually disposed.
        /// </summary>
        public Application GetNewApplication()
        {
            return new Application(InternalObject.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        /// <summary>
        /// Returns or sets the state of the specified document window or task window.
        /// </summary>
        public WindowState WindowState
        {
            get
            {
                return (WindowState)InternalObject.GetType().InvokeMember("WindowState", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("WindowState", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
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
