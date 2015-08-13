using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents a contiguous area in a document. Each Range object is defined by a starting and ending character position.
    /// </summary>
    [WrapsCOM("Word.Range", "0002095E-0000-0000-C000-000000000046")]
    public class Range : ComWrapper
    {
        private ContentControls _contentControls;
        private Fields _fields;
        private InlineShapes _inlineShapes;

        internal Range(object rangeObj)
            : base(rangeObj)
        {
        }

        /// <summary>
        /// Collapses a range or selection to the starting or ending position. After a range or selection is collapsed, the starting and ending points are equal.
        /// </summary>
        /// <param name="direction">The direction in which to collapse the range or selection.</param>
        public void Collapse(Enums.CollapseDirection? direction = null)
        {
            List<object> args = new List<object>();
            args.Add(direction ?? (object)System.Reflection.Missing.Value);

            InternalObject.GetType().InvokeMember("Collapse", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, args.ToArray());
        }

        /// <summary>
        /// Gets a ContentControls collection that represents all the Content Controls in a range.
        /// This collection object is internally cached and does not need to be manually disposed.
        /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
        /// </summary>
        public ContentControls ContentControls
        {
            get
            {
                if (_contentControls == null)
                {
                    _contentControls = new ContentControls(InternalObject.GetType().InvokeMember("ContentControls", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _contentControls;
            }
        }

        /// <summary>
        /// Copies the specified range to the Clipboard.
        /// </summary>
        public void Copy()
        {
            InternalObject.GetType().InvokeMember("Copy", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Deletes the specified number of characters or words.
        /// </summary>
        /// <param name="unit">The unit by which the collapsed range is to be deleted.</param>
        /// <param name="count">The number of units to be deleted. To delete units after the range, collapse the range and use a positive number. To delete units before the range, collapse the range and use a negative number.</param>
        public void Delete(Enums.Units? unit = null, long? count = null)
        {
            var args = new List<object>();
            args.Add(unit ?? (object)System.Reflection.Missing.Value);
            args.Add(count ?? (object)System.Reflection.Missing.Value);

            InternalObject.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, args.ToArray());
        }

        /// <summary>
        /// Returns or sets the ending character position of a range. Read/write Long.
        /// </summary>
        /// <remarks>
        /// Range objects all have a starting position and an ending position. The ending position is the point farthest away from the beginning of the story. If this property is set to a value smaller than the Start property, the Start property is set to the same value (that is, the Start and End property are equal).
        /// This property returns the ending character position relative to the beginning of the story. The main document story (wdMainTextStory) begins with character position 0 (zero). You can change the size of a selection, range, or bookmark by setting this property.
        /// </remarks>
        public long End
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("End", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("End", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Expands the specified range or selection. Returns the number of characters added to the range or selection.
        /// </summary>
        /// <param name="unit">The unit by which to expand the range. Can be one of the following Units constants: Character, Word, Sentence, Paragraph, Section, Story, Cell, Column, Row, or Table.</param>
        /// <returns></returns>
        public long Expand(Enums.Units unit)
        {
            return Convert.ToInt64(InternalObject.GetType().InvokeMember("Expand", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { unit }));
        }

        /// <summary>
        /// Returns a Fields collection that represents all the fields in the range.
        /// This collection object is internally cached and does not need to be manually disposed.
        /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
        /// </summary>
        public Fields Fields
        {
            get
            {
                if (_fields == null)
                {
                    _fields = new Fields(InternalObject.GetType().InvokeMember("Fields", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _fields;
            }
        }

        /// <summary>
        /// Gets or sets a Font object that represents the character formatting of the specified object.
        /// This object must be manually disposed.
        /// </summary>
        public Font Font
        {
            get
            {
                return new Font(InternalObject.GetType().InvokeMember("Font", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Font", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns an InlineShapes collection that represents all the InlineShape objects in a range.
        /// This collection object is internally cached and does not need to be manually disposed.
        /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
        /// </summary>
        public InlineShapes InlineShapes
        {
            get
            {
                if (_inlineShapes == null)
                {
                    _inlineShapes = new InlineShapes(InternalObject.GetType().InvokeMember("InlineShapes", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _inlineShapes;
            }
        }

        /// <summary>
        /// Inserts a page, column, or section break.
        /// When you insert a page or column break, the break replaces the range. If you don't want to replace the range, use the Collapse method before using the InsertBreak method.
        /// </summary>
        /// <param name="type"></param>
        public void InsertBreak(Enums.BreakType type)
        {
            InternalObject.GetType().InvokeMember("InsertBreak", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { type });
        }

        /// <summary>
        /// Inserts the contents of the Clipboard at the specified range.
        /// </summary>
        public void Paste()
        {
            InternalObject.GetType().InvokeMember("Paste", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Returns or sets the starting character position of a range. Read/write Long.
        /// </summary>
        /// <remarks>
        /// Range objects have starting and ending character positions. The starting position refers to the character position closest to the beginning of the story. If this property is set to a value larger than that of the End property, the End property is set to the same value as that of Start property.
        /// This property returns the starting character position relative to the beginning of the story. The main text story (wdMainTextStory) begins with character position 0 (zero). You can change the size of a selection, range, or bookmark by setting this property.
        /// </remarks>
        public long Start
        {
            get
            {
                return Convert.ToInt64(InternalObject.GetType().InvokeMember("Start", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Start", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets the text in the specified range or selection.
        /// </summary>
        public string Text
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }

            set
            {
                InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_contentControls != null && !_contentControls.IsDisposed)
                {
                    _contentControls.Dispose();
                    _contentControls = null;
                }

                if (_fields != null && !_fields.IsDisposed)
                {
                    _fields.Dispose();
                    _fields = null;
                }

                if (_inlineShapes != null && !_inlineShapes.IsDisposed)
                {
                    _inlineShapes.Dispose();
                    _inlineShapes = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
