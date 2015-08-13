using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Core.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents a single InlineShape in a document.
    /// </summary>
    [WrapsCOM("Word.InlineShape", "000209A8-0000-0000-C000-000000000046")]
    public class InlineShape : ComWrapper
    {
        internal InlineShape(object inlineShapeObj)
            : base(inlineShapeObj)
        {
        }

        /// <summary>
        /// Deletes the specified InlineShape.
        /// </summary>
        public void Delete()
        {
            InternalObject.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Returns or sets a value that represents the height, in points, of the InlineShape.
        /// </summary>
        public Single Height
        {
            get
            {
                return Convert.ToSingle(InternalObject.GetType().InvokeMember("Height", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Height", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns a Range object that represents the portion of a document that is contained within an inline shape.
        /// This Range object is NOT internally cached and must be manually disposed.
        /// </summary>
        public Range Range
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        public TriState LockAspectRatio
        {
            get
            {
                return (TriState)Convert.ToUInt32(InternalObject.GetType().InvokeMember("LockAspectRatio", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("LockAspectRatio", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a value that represents the width, in points, of the InlineShape.
        /// </summary>
        public Single Width
        {
            get
            {
                return Convert.ToSingle(InternalObject.GetType().InvokeMember("Width", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Width", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
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
