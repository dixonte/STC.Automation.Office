using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps the Word.Application COM object
    /// </summary>
    [WrapsCOM("Word.Application", Application.UUID)]
    public class Application : OfficeApplication
    {
        public const string UUID = "00020970-0000-0000-C000-000000000046";

        private Documents _documents;
        private Selection _selection;

        /// <summary>
        /// Creates a new instance of Word for the purposes of automation
        /// </summary>
        public Application()
            : base()
        {
        }

        internal Application(object applicationObj)
            : base(applicationObj)
        {
        }

        /// <summary>
        /// Attempts to attach to an already running Word process.
        /// </summary>
        /// <param name="processToAttach">The Process object to which to attach.</param>
        /// <returns>An Application wrapper.</returns>
        public static Application FromProcess(Process processToAttach)
        {
            using (Window window = ComWrapper.FromProcess<Window>(processToAttach, "_WwG"))
            {
                if (window != null)
                {
                    return window.GetNewApplication();
                }
                else
                {
                    return null;
                }
            }
        }

        public static IList<Application> GetRunningApplications()
        {
            return Application.FromROT<Application>();
        }

        /// <summary>
        /// Activates the specified object.
        /// </summary>
        public void Activate()
        {
            InternalObject.GetType().InvokeMember("Activate", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Gets or sets the name of the active printer. Read/write String.
        /// </summary>
        public string ActivePrinter
        {
            get
            {
                return InternalObject.GetType().InvokeMember("ActivePrinter", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }

            set
            {
                InternalObject.GetType().InvokeMember("ActivePrinter", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Provides methods for dealing with workbooks (.xls files). This object is internally cached and does not require manual disposal.
        /// </summary>
        public Documents Documents
        {
            get
            {
                if (_documents == null)
                {
                    _documents = new Documents(InternalObject.GetType().InvokeMember("Documents", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _documents;
            }
        }

        /// <summary>
        /// Tells Excel to close itself. It may not actually close if you are still holding references to Excel objects; use of the using() clause is recommended.
        /// </summary>
        public void Quit()
        {
            InternalObject.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        public object Run(string proc, params object[] args)
        {
            List<object> inArgs = new List<object>(args);
            inArgs.Insert(0, proc);

            return InternalObject.GetType().InvokeMember("Run", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, inArgs.ToArray());
        }

        /// <summary>
        /// Gets the Selection object that represents a selected range or the insertion point.
        /// </summary>
        public Selection Selection
        {
            get
            {
                if (_selection == null)
                {
                    _selection = new Selection(InternalObject.GetType().InvokeMember("Selection", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _selection;
            }
        }

        /// <summary>
        /// Returns the Microsoft Word version number.
        /// </summary>
        public Version Version
        {
            get
            {
                return new Version(InternalObject.GetType().InvokeMember("Version", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString());
            }
        }

        /// <summary>
        /// Gets or sets the visibility of the Word program window
        /// </summary>
        public bool Visible
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
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
                if (_documents != null)
                {
                    _documents.Dispose();
                    _documents = null;
                }

                if (_selection != null)
                {
                    _selection.Dispose();
                    _selection = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
