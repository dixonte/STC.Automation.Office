using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// A collection of Field objects that represent all the fields in a selection, range, or document.
    /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
    /// </summary>
    [WrapsCOM("Word.Fields", "00020930-0000-0000-C000-000000000046")]
    public class Fields : ComWrapper, IEnumerable<Field>
    {
        internal Fields(object fieldsObj)
            : base(fieldsObj)
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
        /// Returns an individual Field object in the collection. The returned Field must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the Field. 1-indexed.</param>
        /// <returns>A Field object.</returns>
        public Section this[long index]
        {
            get
            {
                return new Section(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
        }

        #region IEnumerable<Section> Members

        /// <summary>
        /// Gets a generic IEnumerator of Field objects.
        /// </summary>
        /// <returns>IEnumerator&lt;Field&gt;</returns>
        public IEnumerator<Field> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Field>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an IEnumerator of Field objects.
        /// </summary>
        /// <returns>IEnumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Field>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

    }
}
