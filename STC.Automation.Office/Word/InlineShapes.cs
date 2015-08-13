using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word.InlineShapes object. If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
    /// </summary>
    [WrapsCOM("Word.InlineShapes", "000209A9-0000-0000-C000-000000000046")]
    public class InlineShapes : ComWrapper, IEnumerable<InlineShape>
    {
        internal InlineShapes(object inlineShapesObj)
            : base(inlineShapesObj)
        {
        }

        /// <summary>
        /// Returns an individual InlineShape object in the collection. The returned InlineShape must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the InlineShape. 1-indexed.</param>
        /// <returns>An InlineShape object.</returns>
        public InlineShape this[long index]
        {
            get
            {
                return new InlineShape(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
        }

        /// <summary>
        /// Returns an individual InlineShape object in the collection. The returned InlineShape must be manually disposed.
        /// </summary>
        /// <param name="index">Index of the InlineShape. 1-indexed.</param>
        /// <returns>An InlineShape object.</returns>
        public InlineShape this[string index]
        {
            get
            {
                return new InlineShape(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
            }
        }

        /// <summary>
        /// Adds a picture to a document. Returns an InlineShape object that represents the picture.
        /// </summary>
        /// <param name="filename">The path and file name of the picture.</param>
        /// <param name="linkToFile">True to link the picture to the file from which it was created. False to make the picture an independent copy of the file. The default value is False.</param>
        /// <param name="saveWithDocument">True to save the linked picture with the document. The default value is False.</param>
        /// <param name="range">The location where the picture will be placed in the text. If the range isn't collapsed, the picture replaces the range; otherwise, the picture is inserted. If this argument is omitted, the picture is placed automatically.</param>
        /// <returns>InlineShape</returns>
        public InlineShape AddPicture(string filename, bool? linkToFile = null, bool? saveWithDocument = null, Range range = null)
        {
            if (string.IsNullOrEmpty(filename) || !System.IO.File.Exists(filename))
                throw new ArgumentException("Parameter 'filename' must be set, and must be for a file that exists.");

            var args = new List<object>();

            args.Add(filename);
            args.Add(linkToFile ?? (object)System.Reflection.Missing.Value);
            args.Add(saveWithDocument ?? (object)System.Reflection.Missing.Value);
            args.Add(range != null ? range.InternalObject : System.Reflection.Missing.Value);

            return new InlineShape(InternalObject.GetType().InvokeMember("AddPicture", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, args.ToArray()));
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

        #region IEnumerable<InlineShape> Members

        /// <summary>
        /// Gets a generic IEnumerator of InlineShape objects.
        /// </summary>
        /// <returns>IEnumerator&lt;InlineShape&gt;</returns>
        public IEnumerator<InlineShape> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<InlineShape>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an IEnumerator of InlineShape objects.
        /// </summary>
        /// <returns>IEnumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ComIEnumeratorWrapper<InlineShape>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion
    }
}
