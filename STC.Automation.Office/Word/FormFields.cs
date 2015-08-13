using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// A collection of FormField objects that represent all the form fields in a selection, range, or document.
    /// </summary>
    [WrapsCOM("Word.FormFields", "00020929-0000-0000-C000-000000000046")]
    public class FormFields : ComWrapper
    {
        internal FormFields(object formFieldsObj)
            : base(formFieldsObj)
        {
        }

        /// <summary>
        /// Returns a FormField object that represents a new form field added at a range.
        /// The returned object must be manually disposed.
        /// </summary>
        /// <param name="range">The range where you want to add the form field. If the range isn't collapsed, the form field replaces the range.</param>
        /// <param name="type">The type of form field to add.</param>
        /// <returns>FormField</returns>
        public FormField Add(Range range, FieldType type)
        {
            return new FormField(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { range.InternalObject, type }));
        }
    }
}
