using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel Shape object. Could contain a picture, chart, label, callout, etc.
    /// </summary>
    [WrapsCOM("Excel.AutoFilter", "00024432-0000-0000-C000-000000000046")]
    public class AutoFilter : ComWrapper
    {
        private Sort _sort;

        internal AutoFilter(object autoFilterObj)
            : base(autoFilterObj)
        {
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_sort != null)
                {
                    _sort.Dispose();
                    _sort = null;
                }
            }

            base.Dispose(true);
        }

        #endregion

        /// <summary>
        /// Gets or sets a String value representing the name of the Shape. This object is internally cached and does not require manual disposal.
        /// </summary>
        public Sort Sort
        {
            get
            {
                if (_sort == null)
                    _sort = new Sort(InternalObject.GetType().InvokeMember("Sort", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _sort;
            }
        }

        /// <summary>
        /// Returns or sets a value that represents the width, in points, of the object.
        /// </summary>
        public double Width
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Width", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Width", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a value that represents the height, in points, of the object.
        /// </summary>
        public double Height
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Height", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Height", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
