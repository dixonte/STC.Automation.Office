using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// A collection of HeaderFooter objects that represent the headers or footers in the specified section of a document.
    /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
    /// </summary>
    [WrapsCOM("Word.HeadersFooters", "00020984-0000-0000-C000-000000000046")]
    public class HeadersFooters : ComWrapper, IEnumerable<HeaderFooter>
    {
        internal HeadersFooters(object headersFootersObj)
            : base(headersFootersObj)
        {
        }

        /// <summary>
        /// Gets the number of items in this collection.
        /// </summary>
        public long Count
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns an individual HeaderFooter object in the collection. The returned HeaderFooter must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the HeaderFooter. 1-indexed.</param>
        /// <returns>A HeaderFooter object.</returns>
        public HeaderFooter this[long index]
        {
            get
            {
                return new HeaderFooter(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
        }

        #region IEnumerable<HeaderFooter> Members

        /// <summary>
        /// Gets a generic IEnumerator of HeaderFooter objects.
        /// </summary>
        /// <returns>IEnumerator&lt;HeaderFooter&gt;</returns>
        public IEnumerator<HeaderFooter> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<HeaderFooter>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets a IEnumerator of HeaderFooter objects.
        /// </summary>
        /// <returns>IEnumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ComIEnumeratorWrapper<HeaderFooter>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion
    }
}
