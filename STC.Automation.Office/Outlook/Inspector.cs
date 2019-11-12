using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Represents the window in which an Outlook item is displayed.
    /// </summary>
    [WrapsCOM("Outlook.Inspector", "00063005-0000-0000-C000-000000000046")]
    public class Inspector : ComWrapper
    {
        internal Inspector(object explorerObj)
            : base(explorerObj)
        {
        }

        /// <summary>
        /// Activates an inspector window by bringing it to the foreground and setting keyboard focus.
        /// </summary>
        public void Activate()
        {
            InternalObject.GetType().InvokeMember("Activate", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { });
        }
    }
}
