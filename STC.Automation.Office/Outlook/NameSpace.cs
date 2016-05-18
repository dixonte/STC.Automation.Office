using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Wraps an Outlook.NameSpace object
    /// </summary>
    [WrapsCOM("Outlook.NameSpace")] //, "000208DA-0000-0000-C000-000000000046")]
    public class NameSpace : ComWrapper
    {
        internal NameSpace(object namespaceObj)
            : base(namespaceObj)
        {
        }

        /// <summary>
        /// Returns a Folder object that represents the default folder of the requested type for the current profile; for example, obtains the default Calendar folder for the user who is currently logged on.
        /// </summary>
        /// <param name="folderType">The type of default folder to return.</param>
        public Folder GetDefaultFolder(Enums.DefaultFolders folderType)
        {
            return new Folder(InternalObject.GetType().InvokeMember("GetDefaultFolder", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { folderType }));
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
