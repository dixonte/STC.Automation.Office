using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents a single header or footer.
    /// </summary>
    [WrapsCOM("Word.HeaderFooter", "00020985-0000-0000-C000-000000000046")]
    public class HeaderFooter : ComWrapper
    {
        internal HeaderFooter(object headerFooterObj)
            : base(headerFooterObj)
        {
        }

        /// <summary>
        /// Gets the index of this HeaderFooter.
        /// </summary>
        public long Index
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("Index", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// True if the specified HeaderFooter object is a header. Read-only Boolean.
        /// </summary>
        public bool IsHeader
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("IsHeader", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns a Range object that represents the portion of a document that is contained within the specified header or footer.
        /// This Range object is NOT internally cached and must be manually disposed.
        /// </summary>
        public Range Range
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
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
