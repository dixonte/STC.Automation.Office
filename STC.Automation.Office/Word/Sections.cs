using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.IO;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word.Sections object. If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
    /// </summary>
    [WrapsCOM("Word.Sections", "0002095A-0000-0000-C000-000000000046")]
    public class Sections : ComWrapper, IEnumerable<Section>
    {
        internal Sections(object sectionsObj)
            : base(sectionsObj)
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
        /// Returns an individual Section object in the collection. The returned Section must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the Section. 1-indexed.</param>
        /// <returns>A Section object.</returns>
        public Section this[long index]
        {
            get
            {
                return new Section(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
        }

        #region IEnumerable<Section> Members

        /// <summary>
        /// Gets a generic IEnumerator of Section objects.
        /// </summary>
        /// <returns>IEnumerator&lt;Section&gt;</returns>
        public IEnumerator<Section> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Section>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an IEnumerator of Section objects.
        /// </summary>
        /// <returns>IEnumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Section>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion
    }
}
