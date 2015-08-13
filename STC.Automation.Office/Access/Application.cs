using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using STC.Automation.Office.Common;
using STC.Automation.Office.Excel.Events;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Core;

namespace STC.Automation.Office.Access
{
    /// <summary>
    /// Wraps the Access.Application COM object
    /// </summary>
    [WrapsCOM("Access.Application", Application.UUID)]
    public class Application : OfficeApplication
    {
        public const string UUID = "68CCE6C0-6129-101B-AF4E-00AA003F0F07";

        private CommandBars _commandBars;

        /// <summary>
        /// Creates a new instance of Access for the purposes of automation
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
        /// Attempts to attach to an already running Access process.
        /// </summary>
        /// <param name="processToAttach">The Process object to which to attach.</param>
        /// <returns>An Application wrapper.</returns>
        public static Application FromProcess(Process processToAttach)
        {
            var accesses = Application.FromROT<Application>();

            Application foundAccess = null;

            foreach (var access in accesses)
            {
                if (access.hWnd == processToAttach.MainWindowHandle)
                {
                    foundAccess = access;
                    continue;
                }

                access.Dispose();
            }
            accesses.Clear();

            return foundAccess;
        }

        public static IList<Application> GetRunningApplications()
        {
            return Application.FromROT<Application>();
        }

        public CommandBars CommandBars
        {
            get
            {
                if (_commandBars == null || _commandBars.IsDisposed)
                    _commandBars = new CommandBars(InternalObject.GetType().InvokeMember("CommandBars", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _commandBars;
            }
        }

        public IntPtr hWnd
        {
            get
            {
                return (IntPtr)(int)InternalObject.GetType().InvokeMember("hWndAccessApp", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
            }
        }

        /// <summary>
        /// Returns the Microsoft Access version number.
        /// </summary>
        public Version Version
        {
            get
            {
                return new Version(InternalObject.GetType().InvokeMember("Version", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString());
            }
        }

        /// <summary>
        /// Gets or sets the visibility of the Access program window
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
        /// Tells Access to close itself. It may not actually close if you are still holding references to Access objects; use of the using() clause is recommended.
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

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_commandBars != null && !_commandBars.IsDisposed)
                {
                    _commandBars.Dispose();
                    _commandBars = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
