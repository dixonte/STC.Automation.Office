using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents a single form field. The FormField object is a member of the FormFields collection.
    /// </summary>
    [WrapsCOM("Word.FormField", "00020928-0000-0000-C000-000000000046")]
    public class FormField : ComWrapper
    {
        internal FormField(object formFieldObj)
            : base(formFieldObj)
        {
        }
    }
}
