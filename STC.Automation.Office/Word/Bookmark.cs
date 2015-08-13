using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents a single bookmark in a document, selection, or range.
    /// </summary>
    [WrapsCOM("Word.Bookmark", "00020968-0000-0000-C000-000000000046")]
    public class Bookmark : ComWrapper
    {
        internal Bookmark(object bookmarkObj)
            : base(bookmarkObj)
        {
        }

        /// <summary>
        /// Deletes the specified bookmark.
        /// </summary>
        public void Delete()
        {
            InternalObject.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Gets the name of the bookmark.
        /// </summary>
        public string Name
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }
        }

        /// <summary>
        /// Selects the bookmark.
        /// </summary>
        public void Select()
        {
            InternalObject.GetType().InvokeMember("Select", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }
    }
}
