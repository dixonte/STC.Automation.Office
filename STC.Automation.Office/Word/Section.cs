using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using System.Runtime.InteropServices;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word.Section object
    /// </summary>
    public class Section : ComWrapper
    {
        private HeadersFooters _headers, _footers;

        internal Section(object sectionObj)
            : base(sectionObj)
        {
            // Check if supports interface _Document (reported by OleView.exe)
            if (!SupportsInterface(InternalObject, new Guid("00020959-0000-0000-C000-000000000046")))
            {
                throw new COMException("Problem wrapping Word.Section object; does not support interface {00020959-0000-0000-C000-000000000046}.");
            }
        }

        /// <summary>
        /// Gets the index of this Section.
        /// </summary>
        public long Index
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("Index", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets a collection of headers in this section. This collection object is internally cached and does not need to be manually disposed.
        /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
        /// </summary>
        public HeadersFooters Headers
        {
            get
            {
                if (_headers == null)
                {
                    _headers = new HeadersFooters(InternalObject.GetType().InvokeMember("Headers", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _headers;
            }
        }

        /// <summary>
        /// Gets a collection of footers in this section. This collection object is internally cached and does not need to be manually disposed.
        /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
        /// </summary>
        public HeadersFooters Footers
        {
            get
            {
                if (_footers == null)
                {
                    _footers = new HeadersFooters(InternalObject.GetType().InvokeMember("Footers", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _footers;
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
                if (_headers != null)
                {
                    _headers.Dispose();
                    _headers = null;
                }

                if (_footers != null)
                {
                    _footers.Dispose();
                    _footers = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
