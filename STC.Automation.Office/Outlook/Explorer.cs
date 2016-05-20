using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Represents the window in which the contents of a folder are displayed.
    /// </summary>
    [WrapsCOM("Outlook.Explorer", "00063003-0000-0000-C000-000000000046")]
    public class Explorer : ComWrapper
    {
        internal Explorer(object explorerObj)
            : base(explorerObj)
        {
        }
        
        /// <summary>
        /// Returns or sets a Folder object that represents the current folder displayed in the explorer. Read/write.  This object must be disposed manually.
        /// </summary>
        public Folder CurrentFolder
        {
            get
            {
                return new Folder(InternalObject.GetType().InvokeMember("CurrentFolder", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
            set
            {
                InternalObject.GetType().InvokeMember("CurrentFolder", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value.InternalObject });
            }
        }


        /// <summary>
        /// Activates an explorer window by bringing it to the foreground and setting keyboard focus.
        /// </summary>
        public void Activate()
        {
            InternalObject.GetType().InvokeMember("Activate", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { });
        }
    }
}
