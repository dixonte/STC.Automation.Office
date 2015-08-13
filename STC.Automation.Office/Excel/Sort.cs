using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel Shape object. Could contain a picture, chart, label, callout, etc.
    /// </summary>
    [WrapsCOM("Excel.Sort", "000244AB-0000-0000-C000-000000000046")]
    public class Sort : ComWrapper
    {
        private SortFields _sortFields;

        internal Sort(object sortObj)
            : base(sortObj)
        {
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_sortFields != null)
                {
                    _sortFields.Dispose();
                    _sortFields = null;
                }
            }

            base.Dispose(true);
        }

        #endregion

        /// <summary>
        /// Provides methods for dealing with sort fields. This object is internally cached and does not require manual disposal.
        /// </summary>
        public SortFields SortFields
        {
            get
            {
                if (_sortFields == null)
                    _sortFields = new SortFields(InternalObject.GetType().InvokeMember("SortFields", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _sortFields;
            }
        }

        /// <summary>
        /// Specifies whether the first row contains header information.
        /// </summary>
        public YesNoGuess Header
        {
            get
            {
                return (YesNoGuess)InternalObject.GetType().InvokeMember("Header", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Header", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { (int)value });
            }
        }

        /// <summary>
        /// Set to True to perform a case-sensitive sort or set to False to perform non-case sensitive sort.
        /// </summary>
        public bool MatchCase
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("MatchCase", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("MatchCase", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Specifies the orientation for the sort.
        /// </summary>
        public SortOrientation Orientation
        {
            get
            {
                return (SortOrientation)InternalObject.GetType().InvokeMember("Orientation", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Orientation", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { (int)value });
            }
        }

        /// <summary>
        /// Sorts the range based on the currently applied sort states.
        /// </summary>
        public void Apply()
        {
            InternalObject.GetType().InvokeMember("Apply", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }
    }
}
