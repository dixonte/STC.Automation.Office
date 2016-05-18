using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Wraps an Outlook.Folder object
    /// </summary>
    [WrapsCOM("Outlook.Folder")] //, "000208DA-0000-0000-C000-000000000046")]
    public class Folder : ComWrapper
    {
        internal Folder(object folderObj)
            : base(folderObj)
        {
        }

        /// <summary>
        /// Displays a new Explorer object for the folder.
        /// </summary>
        public void Display()
        {
            InternalObject.GetType().InvokeMember("Display", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { });
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
