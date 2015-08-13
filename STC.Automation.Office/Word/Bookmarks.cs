using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word.Bookmarks object. If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
    /// </summary>
    [WrapsCOM("Word.Bookmarks", "00020967-0000-0000-C000-000000000046")]
    public class Bookmarks : ComWrapper, IEnumerable<Bookmark>
    {
        internal Bookmarks(object bookmarksObj)
            : base(bookmarksObj)
        {
        }

        /// <summary>
        /// Returns an individual Bookmark object in the collection. The returned Bookmark must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the Bookmark. 1-indexed.</param>
        /// <returns>A Bookmark object.</returns>
        public Bookmark this[long index]
        {
            get
            {
                return new Bookmark(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
        }

        /// <summary>
        /// Returns an individual Bookmark object in the collection. The returned Bookmark must be manually disposed.
        /// </summary>
        /// <param name="index">Name of the Bookmark.</param>
        /// <returns>A Bookmark object.</returns>
        public Bookmark this[string index]
        {
            get
            {
                return new Bookmark(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
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

        /// <summary>
        /// Determines whether the specified bookmark exists.
        /// </summary>
        /// <param name="name">A bookmark name.</param>
        /// <returns>Returns True if the bookmark exists.</returns>
        public bool Exists(string name)
        {
            return Convert.ToBoolean(InternalObject.GetType().InvokeMember("Exists", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { name }));
        }

        #region IEnumerable<Bookmark> Members

        /// <summary>
        /// Gets a generic IEnumerator of Bookmark objects.
        /// </summary>
        /// <returns>IEnumerator&lt;Bookmark&gt;</returns>
        public IEnumerator<Bookmark> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Bookmark>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an IEnumerator of Bookmark objects.
        /// </summary>
        /// <returns>IEnumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Bookmark>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion
    }
}
