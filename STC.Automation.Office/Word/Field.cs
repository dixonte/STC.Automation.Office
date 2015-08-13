using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents a field.
    /// </summary>
    [WrapsCOM("Word.Field", "0002092F-0000-0000-C000-000000000046")]
    public class Field : ComWrapper
    {
        internal Field(object fieldObj)
            : base(fieldObj)
        {
        }

        /// <summary>
        /// True if the specified field is locked. Read/write Boolean.
        /// </summary>
        public bool Locked
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                Convert.ToBoolean(InternalObject.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value }));
            }
        }

        /// <summary>
        /// Returns a Range object that represents a field's code.
        /// This Range object is NOT internally cached and does need to be manually disposed.
        /// </summary>
        public Range Code
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("Code", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns a Range object that represents a field's result.
        /// This Range object is NOT internally cached and does need to be manually disposed.
        /// </summary>
        public Range Result
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("Result", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
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
