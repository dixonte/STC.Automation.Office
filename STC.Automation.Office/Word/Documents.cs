using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.IO;
using STC.Automation.Office.Attributes;
using System.Text.RegularExpressions;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word.Documents object. If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
    /// </summary>
    [WrapsCOM("Word.Documents", "0002096C-0000-0000-C000-000000000046")]
    public class Documents : ComWrapper, IEnumerable<Document>
    {
        internal Documents(object documentsObj)
            : base(documentsObj)
        {
        }

        /// <summary>
        /// Creates a new Document. The returned document must be manually disposed.
        /// </summary>
        /// <returns>A reference to the new Document</returns>
        public Document Add()
        {
            return new Document(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null));
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
        /// Opens an existing Document. The returned document must be manually disposed.
        /// </summary>
        /// <param name="filename">The full path to the document file (.doc)</param>
        /// <returns>A reference to the Document</returns>
        /// <exception cref="FileNotFoundException">If the file does not exist</exception>
        public Document Open(string filename)
        {
            return Open(filename, null);
        }

        /// <summary>
        /// Opens an existing Document. The returned document must be manually disposed.
        /// </summary>
        /// <param name="filename">The full path to the document file (.doc)</param>
        /// <param name="readOnly">Open the document as read only</param>
        /// <returns>A reference to the Document</returns>
        /// <exception cref="FileNotFoundException">If the file does not exist</exception>
        public Document Open(string filename, bool? readOnly)
        {
            if (Regex.IsMatch(filename, STC.Automation.Office.Properties.Resources.FilepathRegex) && !File.Exists(filename))
                throw new FileNotFoundException();

            return new Document(InternalObject.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, 
                new object[] { 
                    filename, 
                    System.Reflection.Missing.Value, 
                    (readOnly.HasValue ? (object)readOnly : System.Reflection.Missing.Value) 
                }));
        }

        #region IEnumerable<Document> Members

        /// <summary>
        /// Gets a generic IEnumerator of Document objects.
        /// </summary>
        /// <returns>IEnumerator&lt;Document&gt;</returns>
        public IEnumerator<Document> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Document>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an IEnumerator of Document objects.
        /// </summary>
        /// <returns>IEnumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Document>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }

        #endregion
    }
}
