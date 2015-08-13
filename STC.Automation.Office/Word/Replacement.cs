using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents the replace criteria for a find-and-replace operation.
    /// The properties and methods of the Replacement object correspond to the options in the Find and Replace dialog box.
    /// </summary>
    [WrapsCOM("Word.Replacement", "000209B1-0000-0000-C000-000000000046")]
    public class Replacement : ComWrapper
    {
        internal Replacement(object replacementObj)
            : base(replacementObj)
        {
        }

        /// <summary>
        /// Removes text and paragraph formatting from the text specified in a replace operation.
        /// </summary>
        public void ClearFormatting()
        {
            InternalObject.GetType().InvokeMember("ClearFormatting", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Gets or sets the text to replace.
        /// </summary>
        public string Text
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }

            set
            {
                InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
