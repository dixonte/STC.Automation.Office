using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Represents a user or resource in Outlook, generally a mail or mobile message addressee.
    /// </summary>
    [WrapsCOM("Outlook.Recipient", "00063045-0000-0000-C000-000000000046")]
    public class Recipient : ComWrapper
        //where T : struct, IConvertible // Enum
    {
        internal Recipient(object recipientObj)
            : base(recipientObj)
        {
        }

        /// <summary>
        /// Returns or sets a long representing the type of recipient. Read/write.
        /// </summary>
        public long Type
        {
            get
            {
                return (long)InternalObject.GetType().InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
            set
            {
                InternalObject.GetType().InvokeMember("Type", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Attempts to resolve a Recipient object against the Address Book.
        /// </summary>
        /// <returns>True if the object was resolved; otherwise, False.</returns>
        public bool Resolve()
        {
            return (bool)InternalObject.GetType().InvokeMember("Resolve", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { });
        }
    }
}
