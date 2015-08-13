using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Common
{
    public abstract class OfficeApplication : ComWrapper, IDisposable
    {
        public OfficeApplication()
            : base()
        {
        }

        internal OfficeApplication(object applicationObj)
            : base(applicationObj)
        {
        }

        /// <summary>
        /// Gets or sets the visibility of the main program window
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
        /// Tells the application to close itself. It may not actually close if you are still holding references to its objects; use of the using() clause is recommended.
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
    }
}
