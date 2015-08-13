using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using System.Reflection;
using STC.Automation.Office.Excel.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Shapes object
    /// </summary>
    [WrapsCOM("Excel.SortFields", "000244AA-0000-0000-C000-000000000046")]
    public class SortFields : ComWrapper
    {
        internal SortFields(object sortFieldsObj)
            : base(sortFieldsObj)
        {
        }

        /// <summary>
        /// Creates a new sort field and returns a SortFields object.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public SortFields Add(Range key, SortOn sortOn = SortOn.Values, SortOrder order = SortOrder.Ascending, SortDataOption dataOption = SortDataOption.Normal)
        {
            InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, ComArguments.Prepare(key, sortOn, order, dataOption));

            // I'm returning this since the documentation says it returns 'SortFields' not 'SortField'
            return this;
        }

        /// <summary>
        /// Clears all the SortFields objects.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public void Clear()
        {
            InternalObject.GetType().InvokeMember("Clear", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }
    }
}
