using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word.ContentControls object. If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
    /// </summary>
    [WrapsCOM("Word.ContentControls", "804CD967-F83B-432D-9446-C61A45CFEFF0")]
    public class ContentControls : ComWrapper, IEnumerable<ContentControl>
    {
        internal ContentControls(object contentControlsObj)
            : base(contentControlsObj)
        {
        }

        /// <summary>
        /// Returns an individual ContentControl object in the collection. The returned ContentControl must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the ContentControl. 1-indexed.</param>
        /// <returns>A ContentControl object.</returns>
        public ContentControl this[long index]
        {
            get
            {
                return new ContentControl(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
        }

        /// <summary>
        /// Returns an individual ContentControl object in the collection. The returned ContentControl must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the ContentControl. 1-indexed.</param>
        /// <returns>A ContentControl object.</returns>
        public ContentControl this[string index]
        {
            get
            {
                return new ContentControl(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
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

        #region IEnumerable<ContentControl> Members

        /// <summary>
        /// Gets a generic IEnumerator of ContentControl objects.
        /// </summary>
        /// <returns>IEnumerator&lt;ContentControl&gt;</returns>
        public IEnumerator<ContentControl> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<ContentControl>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an IEnumerator of ContentControl objects.
        /// </summary>
        /// <returns>IEnumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ComIEnumeratorWrapper<ContentControl>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion
    }
}
